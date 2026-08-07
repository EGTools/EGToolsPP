// FxGroup.cpp — GROUPBY / PIVOTBY, 네이티브 인수 완전 정합판 (plan/21 R1).
//
//   GROUPBY(row_fields, values, function, [field_headers], [total_depth],
//           [sort_order], [filter_array], [field_relationship])
//   PIVOTBY(row_fields, col_fields, values, function, [field_headers],
//           [row_total_depth], [row_sort_order], [col_total_depth],
//           [col_sort_order], [filter_array], [relative_to])
//
// 공통 피벗 엔진 하나를 두고 GROUPBY는 "열 축이 없는 PIVOTBY"로 호출한다.
//
// 네이티브와의 의도적 차이(구조 제약, plan/21 D):
//   * function은 eta-lambda 대신 텍스트: "SUM"|"AVERAGE"|"COUNT"|"COUNTA"|
//     "MAX"|"MIN"|"PRODUCT"|"PERCENTOF".
//   * field_relationship(GROUPBY 8번째)은 수용하되 무시. PIVOTBY relative_to는
//     0(총합)/1(행합)/2(열합)만 지원, 3·4(부모)는 #VALUE.
// row_fields·col_fields·values 모두 다중 열 허용(네이티브 동일, 실측 정합):
//   * 다중 col_fields → 열 키 레벨별 헤더 행(Kc행), 다중 values → 열 그룹마다
//     값 열 V개 중첩. 표시 모드(2·3)에서는 열 필드명 행(", " 연결)과
//     행 필드명+값 이름 행이 추가된다.
//   * col_total_depth ±2 = 총계+부분합 열(첫 열필드 블록별, 열필드 2개 이상).
//   * 다중 값(V>1)의 합계/부분합 '열': 네이티브는 빈칸으로 두지만 EG는 값
//     열별로 집계해 채운다 — 의도적 차이(plan/21 D5, 사용자 결정).
// 그 외 인수 의미:
//   * field_headers: 0=헤더없음·표시안함, 1=있음·표시안함, 2=없음·생성표시,
//     3=있음·표시. 생략=자동(첫 행이 전부 텍스트이고 아래에 비텍스트가 있으면
//     헤더로 간주해 표시).
//   * total_depth: 0=없음, 1=총계, 2=총계+부분합(키 2열 이상), 음수=상단 배치.
//     생략=1(네이티브 기본).
//   * sort_order: ±열번호 또는 그 배열. 출력 열 기준(키들→값들). 생략=키 오름차순.
//   * filter_array: 행별 TRUE/FALSE.

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "../core/I18n.h"
#include "../core/Aggregate.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <string>
#include <unordered_map>
#include <vector>
#include <algorithm>
#include <numeric>
#include <limits>
#include <cwctype>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        std::wstring upper(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)std::towupper(c);
            return s;
        }

        bool asNum(const ExcelObj& o, double& d)
        {
            auto t = o.type();
            if (t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool)
                { d = o.get<double>(0.0); return true; }
            return false;
        }

        // 집계 — 공용 코어(core/Aggregate.h) 위임. grandTotal은 PERCENTOF 분모.
        // 지원 목록은 네이티브 GROUPBY/PIVOTBY와 동일한 16종
        // (core::isGroupByAggregator).
        ExcelObj aggregate(const std::wstring& fnUpper, const std::vector<const ExcelObj*>& vals,
                           double grandTotal = 0.0)
        {
            return egtools::core::aggregateObjs(fnUpper, vals, nullptr, grandTotal);
        }

        // 그룹 동등성 정규화 키(숫자=수치, 문자열=대소문자 무시, 타입 구분).
        void appendCanonical(std::wstring& out, const ExcelObj& v)
        {
            switch (v.type())
            {
            case ExcelType::Num:
            case ExcelType::Int:
            {
                wchar_t buf[40];
                swprintf_s(buf, L"#N%.17g", v.get<double>(0.0));
                out += buf;
                break;
            }
            case ExcelType::Bool:
                out += v.get<bool>(false) ? L"#B1" : L"#B0";
                break;
            case ExcelType::Str:
                out += L"#S";
                out += upper(v.toString());
                break;
            case ExcelType::Err:
                out += L"#X";
                out += v.toString();
                break;
            default:
                out += L"#E";
                break;
            }
            out += L'\x01';
        }

        // ---- 공통 피벗 엔진 -------------------------------------------------

        struct PivotArgs
        {
            const ExcelObj* rowF = nullptr;      // 필수 (다열 허용)
            const ExcelObj* colF = nullptr;      // PIVOTBY만 (다열 허용)
            const ExcelObj* values = nullptr;    // 필수 (다열 허용)
            const ExcelObj* fn = nullptr;        // 필수 (텍스트)
            const ExcelObj* fieldHeaders = nullptr;
            const ExcelObj* rowDepth = nullptr;
            const ExcelObj* rowSort = nullptr;
            const ExcelObj* colDepth = nullptr;  // PIVOTBY만
            const ExcelObj* colSort = nullptr;   // PIVOTBY만
            const ExcelObj* filter = nullptr;
            const ExcelObj* relativeTo = nullptr; // PIVOTBY만
        };

        struct Group
        {
            std::vector<ExcelObj> key;
            std::vector<ExcelArray::row_t> rows;
        };

        // ±열번호(단일/배열) 정렬 지정 읽기. 실패 시 false.
        bool readSortSpec(const ExcelObj* v, std::vector<std::pair<int, int>>& out)
        {
            if (!v || v->isMissing()) return true;
            auto add = [&](const ExcelObj& e) -> bool
            {
                double d;
                if (!asNum(e, d) || (int)d == 0) return false;
                out.emplace_back(abs((int)d), (int)d > 0 ? 1 : -1);
                return true;
            };
            if (v->isType(ExcelType::Multi))
            {
                ExcelArray a(*v);
                const size_t n = (size_t)a.nRows() * a.nCols();
                for (size_t k = 0; k < n; ++k)
                    if (!add(a.at(k))) return false;
                return true;
            }
            return add(*v);
        }

        int readInt(const ExcelObj* v, int def)
        {
            return (!v || v->isMissing()) ? def : v->get<int>(def);
        }

        ExcelObj* pivotEngine(const PivotArgs& in)
        {
            if (!in.rowF || in.rowF->isMissing() || !in.values || in.values->isMissing() ||
                !in.fn || in.fn->isMissing())
                return returnValue(CellError::Value);

            ExcelArray rk(*in.rowF), vv(*in.values);
            const size_t Rall = rk.nRows();
            const size_t K = rk.nCols();
            const size_t V = vv.nCols();
            if (Rall == 0 || K == 0 || V == 0 || vv.nRows() != Rall)
                return returnValue(CellError::Value);

            const bool hasCols = in.colF && !in.colF->isMissing();
            ExcelArray ckStore = hasCols ? ExcelArray(*in.colF) : ExcelArray(*in.rowF);
            const size_t Kc = hasCols ? (size_t)ckStore.nCols() : 0;
            if (hasCols && (ckStore.nRows() != Rall || Kc == 0))
                return returnValue(CellError::Value);

            const std::wstring fn = upper(in.fn->toString());
            // 네이티브 16종 + 파라미터형 확장 9종 (core/Aggregate.h).
            if (!egtools::core::isGroupByAggregatorEx(fn))
                return returnValue(CellError::Value);

            // relative_to 이중 용도:
            //   * PERCENTOF — 분모 모드 0 열합(기본)/1 행합/2 총합 (네이티브 정합)
            //   * 파라미터형 확장(LARGE/SMALL/PERCENTILE*/QUARTILE*/TEXTJOIN) —
            //     집계 파라미터(k/p/quart/구분자)를 그대로 전달
            const bool paramAgg = egtools::core::isParamAggregator(fn);
            const ExcelObj* aggParam =
                (paramAgg && in.relativeTo && !in.relativeTo->isMissing()) ? in.relativeTo : nullptr;
            int relativeTo = 0;
            if (fn == L"PERCENTOF")
            {
                relativeTo = readInt(in.relativeTo, 0);
                if (relativeTo < 0 || relativeTo > 2) return returnValue(CellError::Value);
            }

            // 이후의 모든 집계는 파라미터를 실어 나르는 이 람다를 쓴다(외부
            // aggregate()를 가린다 — MTR 안전을 위해 전역 상태 없이 캡처).
            const auto aggregate = [aggParam](const std::wstring& f,
                                              const std::vector<const ExcelObj*>& vals,
                                              double grandTotal = 0.0)
            {
                return egtools::core::aggregateObjs(f, vals, aggParam, grandTotal);
            };

            // ---- field_headers: 생략=자동 감지 ----
            int fieldHeaders;
            if (!in.fieldHeaders || in.fieldHeaders->isMissing())
            {
                bool firstAllText = Rall > 1;
                for (size_t c = 0; firstAllText && c < K; ++c)
                    if (rk.at(0, (ExcelArray::col_t)c).type() != ExcelType::Str) firstAllText = false;
                for (size_t c = 0; firstAllText && c < V; ++c)
                    if (vv.at(0, (ExcelArray::col_t)c).type() != ExcelType::Str) firstAllText = false;
                for (size_t c = 0; firstAllText && hasCols && c < Kc; ++c)
                    if (ckStore.at(0, (ExcelArray::col_t)c).type() != ExcelType::Str) firstAllText = false;
                bool belowHasNonText = false;
                for (size_t r = 1; !belowHasNonText && r < Rall; ++r)
                    for (size_t c = 0; c < V; ++c)
                        if (vv.at((ExcelArray::row_t)r, (ExcelArray::col_t)c).type() != ExcelType::Str)
                        { belowHasNonText = true; break; }
                // 네이티브 자동 감지: 헤더가 있으면 "있음·표시 안 함"(1)로 동작.
                fieldHeaders = (firstAllText && belowHasNonText) ? 1 : 0;
            }
            else
            {
                fieldHeaders = in.fieldHeaders->get<int>(0);
                if (fieldHeaders < 0 || fieldHeaders > 3) return returnValue(CellError::Value);
            }
            const bool dataHasHeaders = fieldHeaders == 1 || fieldHeaders == 3;
            const bool showHeaders = fieldHeaders >= 2;
            const size_t r0 = dataHasHeaders ? 1 : 0;
            if (r0 >= Rall) return returnValue(CellError::Value);

            // 헤더 이름 — 생성 시(모드 2)는 네이티브식 현지화 이름("행 필드 1", "값 1").
            std::vector<std::wstring> keyNames(K), valNames(V);
            for (size_t c = 0; c < K; ++c)
                keyNames[c] = dataHasHeaders
                    ? rk.at(0, (ExcelArray::col_t)c).toString()
                    : egtools::i18n::t(L"agg.rowField") + L" " + std::to_wstring(c + 1);
            for (size_t c = 0; c < V; ++c)
                valNames[c] = dataHasHeaders
                    ? vv.at(0, (ExcelArray::col_t)c).toString()
                    : egtools::i18n::t(L"agg.value") + L" " + std::to_wstring(c + 1);
            std::vector<std::wstring> colFieldNames(Kc);
            for (size_t c = 0; c < Kc; ++c)
                colFieldNames[c] = dataHasHeaders
                    ? ckStore.at(0, (ExcelArray::col_t)c).toString()
                    : egtools::i18n::t(L"agg.colField") + L" " + std::to_wstring(c + 1);

            // ---- total_depth ----
            int rowDepth = readInt(in.rowDepth, 1);
            int colDepth = readInt(in.colDepth, hasCols ? 1 : 0);
            if (abs(rowDepth) > 2) return returnValue(CellError::Value);
            if (abs(colDepth) > 2) return returnValue(CellError::Value);
            if (abs(rowDepth) == 2 && K < 2) rowDepth = rowDepth > 0 ? 1 : -1;
            if (abs(colDepth) == 2 && Kc < 2) colDepth = colDepth > 0 ? 1 : -1;

            // ---- filter ----
            std::vector<ExcelArray::row_t> dataRows;
            if (in.filter && !in.filter->isMissing())
            {
                if (!in.filter->isType(ExcelType::Multi)) return returnValue(CellError::Value);
                ExcelArray fa(*in.filter);
                const size_t FR = (size_t)fa.nRows() * fa.nCols();
                // 헤더 포함/미포함 길이 둘 다 허용.
                const size_t need = Rall - r0;
                const size_t off = (FR == Rall) ? r0 : 0;
                if (FR != need && FR != Rall) return returnValue(CellError::Value);
                for (size_t r = r0; r < Rall; ++r)
                {
                    const ExcelObj& f = fa.at((ExcelArray::row_t)(r - r0 + off));
                    double d;
                    if (asNum(f, d) && d != 0.0) dataRows.push_back((ExcelArray::row_t)r);
                }
            }
            else
                for (size_t r = r0; r < Rall; ++r) dataRows.push_back((ExcelArray::row_t)r);
            if (dataRows.empty()) return returnValue(CellError::NA);

            // ---- 행/열 그룹 (해시, 등장 순서) + 데이터행→그룹 사상 ----
            auto buildGroups = [&dataRows](const ExcelArray& arr, size_t nCols,
                                           std::vector<Group>& groups, std::vector<size_t>& idxOf)
            {
                std::unordered_map<std::wstring, size_t> lookup;
                lookup.reserve(dataRows.size() * 2);
                idxOf.resize(dataRows.size());
                std::wstring canon;
                for (size_t i = 0; i < dataRows.size(); ++i)
                {
                    const auto r = dataRows[i];
                    canon.clear();
                    for (size_t c = 0; c < nCols; ++c)
                        appendCanonical(canon, arr.at(r, (ExcelArray::col_t)c));
                    const auto it = lookup.find(canon);
                    if (it == lookup.end())
                    {
                        Group g;
                        for (size_t c = 0; c < nCols; ++c)
                            g.key.emplace_back(arr.at(r, (ExcelArray::col_t)c));
                        g.rows.push_back(r);
                        idxOf[i] = groups.size();
                        lookup.emplace(canon, groups.size());
                        groups.push_back(std::move(g));
                    }
                    else
                    {
                        idxOf[i] = it->second;
                        groups[it->second].rows.push_back(r);
                    }
                }
            };
            std::vector<Group> rowGroups, colGroups;
            std::vector<size_t> rgIdx, cgIdx;
            buildGroups(rk, K, rowGroups, rgIdx);
            if (hasCols) buildGroups(ckStore, Kc, colGroups, cgIdx);

            // ---- 집계 사전 계산 ----
            const size_t nG = rowGroups.size();
            const size_t nC = hasCols ? colGroups.size() : V;

            // PERCENTOF 분모.
            std::vector<double> grandPerVal(V, 0.0);
            if (fn == L"PERCENTOF")
                for (auto r : dataRows)
                    for (size_t c = 0; c < V; ++c)
                    {
                        double d;
                        if (asNum(vv.at(r, (ExcelArray::col_t)c), d)) grandPerVal[c] += d;
                    }

            // 그룹 셀 수집 헬퍼.
            auto cellsOf = [&](const std::vector<ExcelArray::row_t>& rows, size_t vcol)
            {
                std::vector<const ExcelObj*> cells;
                cells.reserve(rows.size());
                for (auto r : rows) cells.push_back(&vv.at(r, (ExcelArray::col_t)vcol));
                return cells;
            };
            // 행×열 교차 행 목록 (PIVOTBY).
            std::vector<std::vector<std::vector<ExcelArray::row_t>>> cross;
            if (hasCols)
            {
                cross.assign(nG, std::vector<std::vector<ExcelArray::row_t>>(nC));
                for (size_t i = 0; i < dataRows.size(); ++i)
                    cross[rgIdx[i]][cgIdx[i]].push_back(dataRows[i]);
            }

            // ---- 정렬 ----
            auto cmpObj = [](const ExcelObj& a, const ExcelObj& b) -> int
            {
                return ExcelObj::compare(a, b);
            };
            std::vector<std::pair<int, int>> rowSortSpec, colSortSpec;
            if (!readSortSpec(in.rowSort, rowSortSpec) || !readSortSpec(in.colSort, colSortSpec))
                return returnValue(CellError::Value);

            // 정렬 키로 값 열을 지정할 수 있으므로 그룹별 집계값을 먼저 준비(GROUPBY).
            std::vector<std::vector<ExcelObj>> groupAgg;   // GROUPBY: [g][v]
            if (!hasCols)
            {
                groupAgg.resize(nG);
                for (size_t g = 0; g < nG; ++g)
                {
                    groupAgg[g].reserve(V);
                    for (size_t c = 0; c < V; ++c)
                        groupAgg[g].push_back(aggregate(fn, cellsOf(rowGroups[g].rows, c),
                                                        grandPerVal[c]));
                }
            }

            // 정렬 레벨 — 네이티브 실측 규칙: 키 필드의 계층 순서는 항상 유지되고
            // sort_order는 해당 필드의 '방향'만 바꾼다(P.sort). 값 열 지정은 마지막
            // 키 필드 레벨을 대체해 그 안쪽을 정렬하고, 마지막 키는 동점 보조로
            // 뒤에 남는다(G.sortVal — 동점은 키 오름차순).
            struct SortLevel { bool isValue; size_t idx; int dir; };
            auto makeLevels = [](size_t nKeys, size_t nVals,
                                 const std::vector<std::pair<int, int>>& spec)
            {
                std::vector<int> keyDir(nKeys, 1);
                std::vector<SortLevel> valLv;
                for (const auto& [idx, dir] : spec)
                {
                    if (idx >= 1 && (size_t)idx <= nKeys) keyDir[(size_t)idx - 1] = dir;
                    else if ((size_t)idx > nKeys && (size_t)idx <= nKeys + nVals)
                        valLv.push_back(SortLevel{ true, (size_t)idx - nKeys - 1, dir });
                }
                std::vector<SortLevel> lv;
                for (size_t c = 0; c < nKeys; ++c)
                {
                    if (c == nKeys - 1)
                        for (const auto& v : valLv) lv.push_back(v);
                    lv.push_back(SortLevel{ false, c, keyDir[c] });
                }
                return lv;
            };

            const auto rowLevels = makeLevels(K, hasCols ? 0 : V, rowSortSpec);
            std::vector<size_t> rowOrder(nG);
            std::iota(rowOrder.begin(), rowOrder.end(), (size_t)0);
            std::stable_sort(rowOrder.begin(), rowOrder.end(), [&](size_t x, size_t y)
            {
                for (const auto& l : rowLevels)
                {
                    const int cv = l.isValue
                        ? cmpObj(groupAgg[x][l.idx], groupAgg[y][l.idx])
                        : cmpObj(rowGroups[x].key[l.idx], rowGroups[y].key[l.idx]);
                    if (cv != 0) return l.dir > 0 ? cv < 0 : cv > 0;
                }
                return false;
            });

            std::vector<size_t> colOrder(hasCols ? nC : 0);
            if (hasCols)
            {
                const auto colLevels = makeLevels(Kc, 0, colSortSpec);
                std::iota(colOrder.begin(), colOrder.end(), (size_t)0);
                std::stable_sort(colOrder.begin(), colOrder.end(), [&](size_t x, size_t y)
                {
                    for (const auto& l : colLevels)
                    {
                        const int cv = cmpObj(colGroups[x].key[l.idx], colGroups[y].key[l.idx]);
                        if (cv != 0) return l.dir > 0 ? cv < 0 : cv > 0;
                    }
                    return false;
                });
            }

            // ---- 출력 조립 ----
            std::vector<std::vector<ExcelObj>> out;
            const std::wstring kTotal = egtools::i18n::t(L"agg.total");
            const std::wstring kGrandTotal = egtools::i18n::t(L"agg.grandTotal");

            if (!hasCols)
            {
                // ---- GROUPBY: (K + V)열 ----
                if (showHeaders)
                {
                    std::vector<ExcelObj> h;
                    for (auto& n : keyNames) h.emplace_back(std::wstring_view(n));
                    for (auto& n : valNames) h.emplace_back(std::wstring_view(n));
                    out.push_back(std::move(h));
                }

                const bool wantGrand = rowDepth != 0;
                const bool wantSub = abs(rowDepth) == 2;
                const bool atTop = rowDepth < 0;

                // 총계 라벨: 부분합이 있으면 "총합계", 아니면 "합계" (네이티브 현지화).
                auto emitTotalRow = [&](const ExcelObj& label,
                                        const std::vector<ExcelArray::row_t>& rows)
                {
                    std::vector<ExcelObj> row;
                    row.push_back(ExcelObj(label));
                    for (size_t c = 1; c < K; ++c) row.emplace_back(std::wstring_view(L""));
                    for (size_t c = 0; c < V; ++c)
                        row.push_back(aggregate(fn, cellsOf(rows, c), grandPerVal[c]));
                    out.push_back(std::move(row));
                };
                const ExcelObj grandLabel(std::wstring_view(wantSub ? kGrandTotal : kTotal));

                if (wantGrand && atTop) emitTotalRow(grandLabel, dataRows);

                // 첫 키 블록 단위(정렬 순서 기준)로 순회하며 부분합 삽입.
                // 부분합 라벨 = 첫 키 값 그대로(나머지 키 열은 빈칸) — 네이티브 동일.
                size_t i = 0;
                while (i < nG)
                {
                    size_t j = i;
                    std::vector<ExcelArray::row_t> blockRows;
                    while (j < nG &&
                           cmpObj(rowGroups[rowOrder[i]].key[0], rowGroups[rowOrder[j]].key[0]) == 0)
                    {
                        blockRows.insert(blockRows.end(),
                                         rowGroups[rowOrder[j]].rows.begin(),
                                         rowGroups[rowOrder[j]].rows.end());
                        ++j;
                    }
                    if (wantSub && atTop)
                        emitTotalRow(rowGroups[rowOrder[i]].key[0], blockRows);
                    for (size_t g = i; g < j; ++g)
                    {
                        std::vector<ExcelObj> row;
                        for (const auto& k : rowGroups[rowOrder[g]].key) row.emplace_back(k);
                        for (size_t c = 0; c < V; ++c) row.push_back(ExcelObj(groupAgg[rowOrder[g]][c]));
                        out.push_back(std::move(row));
                    }
                    if (wantSub && !atTop)
                        emitTotalRow(rowGroups[rowOrder[i]].key[0], blockRows);
                    i = j;
                }

                if (wantGrand && !atTop) emitTotalRow(grandLabel, dataRows);
            }
            else
            {
                // ---- PIVOTBY: (K + 열라인×V)열 × (Kc[+값이름] 헤더행 + 행라인)행 ----
                const bool wantRowGrand = rowDepth != 0;
                const bool wantRowSub = abs(rowDepth) == 2;
                const bool rowAtTop = rowDepth < 0;
                const bool wantColGrand = colDepth != 0;
                const bool wantColSub = abs(colDepth) == 2;
                const bool colAtLeft = colDepth < 0;

                // 원시 합(PERCENTOF 분모, 값 열별).
                std::vector<std::vector<double>> rawRow, rawCol;
                std::vector<double> rawGrand(V, 0.0);
                if (fn == L"PERCENTOF")
                {
                    rawRow.assign(nG, std::vector<double>(V, 0.0));
                    rawCol.assign(nC, std::vector<double>(V, 0.0));
                    double d;
                    for (size_t i = 0; i < dataRows.size(); ++i)
                        for (size_t v = 0; v < V; ++v)
                            if (asNum(vv.at(dataRows[i], (ExcelArray::col_t)v), d))
                            {
                                rawRow[rgIdx[i]][v] += d;
                                rawCol[cgIdx[i]][v] += d;
                                rawGrand[v] += d;
                            }
                }

                // 출력 라인: 리프(그룹 1개)·부분합(첫 필드 블록)·총계(전체).
                struct Line { std::vector<size_t> gs; int kind; };   // 0=리프 1=부분합 2=총계
                auto buildLines = [&cmpObj](const std::vector<size_t>& order,
                                            const std::vector<Group>& groups,
                                            bool wantGrand, bool wantSub, bool grandFirst)
                {
                    std::vector<Line> lines;
                    if (wantGrand && grandFirst) lines.push_back(Line{ {}, 2 });
                    const size_t n = order.size();
                    size_t i = 0;
                    while (i < n)
                    {
                        size_t j = i;
                        Line sub{ {}, 1 };
                        while (j < n && cmpObj(groups[order[i]].key[0], groups[order[j]].key[0]) == 0)
                            sub.gs.push_back(order[j++]);
                        if (wantSub && grandFirst) lines.push_back(sub);
                        for (size_t g = i; g < j; ++g) lines.push_back(Line{ {order[g]}, 0 });
                        if (wantSub && !grandFirst) lines.push_back(std::move(sub));
                        i = j;
                    }
                    if (wantGrand && !grandFirst) lines.push_back(Line{ {}, 2 });
                    return lines;
                };
                const auto rowLines = buildLines(rowOrder, rowGroups, wantRowGrand, wantRowSub, rowAtTop);
                const auto colLines = buildLines(colOrder, colGroups, wantColGrand, wantColSub, colAtLeft);

                // 셀 값 계산 — 라인 쌍의 교차 행 집합을 집계.
                // PERCENTOF 분모: relative_to 0=열합(기본), 1=행합, 2=총합.
                // 합계 열의 열합 = 총합, 합계 행의 행합 = 총합 (관측된 네이티브 규칙).
                auto lineCells = [&](const Line& rl, const Line& cl, size_t v)
                {
                    std::vector<const ExcelObj*> cells;
                    auto push = [&](const std::vector<ExcelArray::row_t>& rows)
                    {
                        for (auto r : rows) cells.push_back(&vv.at(r, (ExcelArray::col_t)v));
                    };
                    if (rl.kind == 2 && cl.kind == 2) push(dataRows);
                    else if (rl.kind == 2) for (auto c : cl.gs) push(colGroups[c].rows);
                    else if (cl.kind == 2) for (auto g : rl.gs) push(rowGroups[g].rows);
                    else for (auto g : rl.gs) for (auto c : cl.gs) push(cross[g][c]);
                    return cells;
                };
                auto pivotCell = [&](const Line& rl, const Line& cl, size_t v) -> ExcelObj
                {
                    // 의도적 차이(plan/21 D5): 네이티브는 다중 값(V>1)일 때 합계/
                    // 부분합 '열'의 데이터 셀을 비워 두지만, EG는 값 열별로 집계해
                    // 채운다(사용자 결정 2026-08-07). 빈 교차는 종류 무관 빈칸.
                    auto cells = lineCells(rl, cl, v);
                    if (cells.empty())
                        return ExcelObj(std::wstring_view(L""));   // 빈 교차 = 빈칸
                    if (fn != L"PERCENTOF")
                        return aggregate(fn, cells);
                    double num = 0, d;
                    for (auto* o : cells) if (asNum(*o, d)) num += d;
                    double denom = 0;
                    if (relativeTo == 2) denom = rawGrand[v];
                    else if (relativeTo == 1)
                    {
                        if (rl.kind == 2) denom = rawGrand[v];
                        else for (auto g : rl.gs) denom += rawRow[g][v];
                    }
                    else
                    {
                        if (cl.kind == 2) denom = rawGrand[v];
                        else for (auto c : cl.gs) denom += rawCol[c][v];
                    }
                    if (denom == 0.0) return ExcelObj(CellError::Div0);
                    return ExcelObj(num / denom);
                };

                // 라벨: 부분합 존재 시 총계는 "총합계"(행/열 각각 판정).
                const std::wstring& rowGrandLabel = wantRowSub ? kGrandTotal : kTotal;
                const std::wstring& colGrandLabel = wantColSub ? kGrandTotal : kTotal;
                const ExcelObj blank(std::wstring_view(L""));

                // 헤더(네이티브 실측 배치) —
                //   표시 모드(2·3): [열 필드명 ", " 연결 1행] + [열 키 Kc행] +
                //                   [행 필드명 + 값 이름 1행]
                //   숨김 모드(0·1): [열 키 Kc행]만 (값 이름 행 없음, 다중 값이어도).
                // 부분합 열 헤더 = 블록 키(레벨 0)만, 총계 열 = 라벨(레벨 0)만.
                if (showHeaders)
                {
                    std::wstring joined;
                    for (size_t c = 0; c < Kc; ++c)
                    {
                        if (c) joined += L", ";
                        joined += colFieldNames[c];
                    }
                    std::vector<ExcelObj> h;
                    for (size_t c = 0; c < K; ++c) h.emplace_back(blank);
                    h.emplace_back(std::wstring_view(joined));
                    out.push_back(std::move(h));   // 이후 열은 그리드 패딩으로 빈칸
                }
                for (size_t lvl = 0; lvl < Kc; ++lvl)
                {
                    std::vector<ExcelObj> h;
                    for (size_t c = 0; c < K; ++c) h.emplace_back(blank);
                    for (const auto& cl : colLines)
                        for (size_t v = 0; v < V; ++v)
                        {
                            if (cl.kind == 0) h.emplace_back(colGroups[cl.gs[0]].key[lvl]);
                            else if (cl.kind == 1)
                                h.emplace_back(lvl == 0
                                    ? ExcelObj(colGroups[cl.gs[0]].key[0]) : blank);
                            else
                                h.emplace_back(lvl == 0
                                    ? ExcelObj(std::wstring_view(colGrandLabel)) : blank);
                        }
                    out.push_back(std::move(h));
                }
                if (showHeaders)
                {
                    std::vector<ExcelObj> h;
                    for (size_t c = 0; c < K; ++c)
                        h.emplace_back(std::wstring_view(keyNames[c]));
                    for (const auto& cl : colLines)
                        for (size_t v = 0; v < V; ++v)
                            h.emplace_back(std::wstring_view(valNames[v]));
                    out.push_back(std::move(h));
                }

                // 데이터 행: 행 키 K열 + 라인×값 셀.
                // 부분합 행 라벨 = 첫 키 값(나머지 키 열 빈칸), 총계 행 = 라벨.
                for (const auto& rl : rowLines)
                {
                    std::vector<ExcelObj> row;
                    if (rl.kind == 0)
                        for (const auto& k : rowGroups[rl.gs[0]].key) row.emplace_back(k);
                    else
                    {
                        row.emplace_back(rl.kind == 1 ? ExcelObj(rowGroups[rl.gs[0]].key[0])
                                                      : ExcelObj(std::wstring_view(rowGrandLabel)));
                        for (size_t c = 1; c < K; ++c) row.emplace_back(blank);
                    }
                    for (const auto& cl : colLines)
                        for (size_t v = 0; v < V; ++v)
                            row.push_back(pivotCell(rl, cl, v));
                    out.push_back(std::move(row));
                }
            }

            // 그리드 → 배열.
            size_t cols = 0;
            for (const auto& r : out) cols = (std::max)(cols, r.size());
            std::vector<ExcelObj> vals;
            vals.reserve(out.size() * cols);
            for (auto& r : out)
                for (size_t c = 0; c < cols; ++c)
                    vals.push_back(c < r.size() ? std::move(r[c]) : ExcelObj(std::wstring_view(L"")));
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)out.size(), (ExcelArrayBuilder::col_t)cols, vals));
        }
    }

    void registerGroup()
    {
        // 9번째 relative_to는 EGTools 확장(네이티브 GROUPBY는 8인수): 파라미터형
        // 집계자(LARGE/SMALL/PERCENTILE*/QUARTILE*/TEXTJOIN)의 k/p/quart/구분자.
        egtools::core::registerFn(L"GROUPBY",
            [](const ExcelObj& rowF, const ExcelObj& values, const ExcelObj& fn,
               const ExcelObj& fieldHeaders, const ExcelObj& totalDepth,
               const ExcelObj& sortOrder, const ExcelObj& filterArray,
               const ExcelObj& /*fieldRelationship — 미지원, 무시*/,
               const ExcelObj& relativeTo) -> ExcelObj*
            {
                try
                {
                    PivotArgs a;
                    a.rowF = &rowF; a.values = &values; a.fn = &fn;
                    a.fieldHeaders = &fieldHeaders; a.rowDepth = &totalDepth;
                    a.rowSort = &sortOrder; a.filter = &filterArray;
                    a.relativeTo = &relativeTo;
                    return pivotEngine(a);
                }
                catch (...) { return returnValue(CellError::Value); }
            });

        egtools::core::registerFn(L"PIVOTBY",
            [](const ExcelObj& rowF, const ExcelObj& colF, const ExcelObj& values,
               const ExcelObj& fn, const ExcelObj& fieldHeaders,
               const ExcelObj& rowTotalDepth, const ExcelObj& rowSortOrder,
               const ExcelObj& colTotalDepth, const ExcelObj& colSortOrder,
               const ExcelObj& filterArray, const ExcelObj& relativeTo) -> ExcelObj*
            {
                try
                {
                    PivotArgs a;
                    a.rowF = &rowF; a.colF = &colF; a.values = &values; a.fn = &fn;
                    a.fieldHeaders = &fieldHeaders;
                    a.rowDepth = &rowTotalDepth; a.rowSort = &rowSortOrder;
                    a.colDepth = &colTotalDepth; a.colSort = &colSortOrder;
                    a.filter = &filterArray; a.relativeTo = &relativeTo;
                    return pivotEngine(a);
                }
                catch (...) { return returnValue(CellError::Value); }
            });
    }
}
