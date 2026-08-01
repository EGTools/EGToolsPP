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
//   * PIVOTBY의 col_fields·values는 1열만 지원(다중 열 중첩 미지원).
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

        // 집계. grandTotal은 PERCENTOF 분모.
        ExcelObj aggregate(const std::wstring& fnUpper, const std::vector<const ExcelObj*>& vals,
                           double grandTotal = 0.0)
        {
            const std::wstring& f = fnUpper;
            if (f == L"PERCENTOF")
            {
                double s = 0, d;
                for (auto* o : vals) if (asNum(*o, d)) s += d;
                if (grandTotal == 0.0) return ExcelObj(CellError::Div0);
                return ExcelObj(s / grandTotal);
            }
            if (f == L"COUNT")
            {
                int c = 0; double d;
                for (auto* o : vals) if (asNum(*o, d)) ++c;
                return ExcelObj((double)c);
            }
            if (f == L"COUNTA")
            {
                int c = 0;
                for (auto* o : vals)
                {
                    auto t = o->type();
                    if (!(t == ExcelType::Missing || t == ExcelType::Nil
                          || (t == ExcelType::Str && o->stringLength() == 0))) ++c;
                }
                return ExcelObj((double)c);
            }
            if (f == L"SUM")
            {
                double s = 0, d;
                for (auto* o : vals) if (asNum(*o, d)) s += d;
                return ExcelObj(s);
            }
            if (f == L"AVERAGE")
            {
                double s = 0, d; int c = 0;
                for (auto* o : vals) if (asNum(*o, d)) { s += d; ++c; }
                return c ? ExcelObj(s / c) : ExcelObj(CellError::Div0);
            }
            if (f == L"MAX" || f == L"MIN")
            {
                const bool mx = (f == L"MAX");
                double best = mx ? -std::numeric_limits<double>::infinity()
                                 :  std::numeric_limits<double>::infinity();
                bool any = false; double d;
                for (auto* o : vals) if (asNum(*o, d)) { any = true; best = mx ? (d > best ? d : best) : (d < best ? d : best); }
                return ExcelObj(any ? best : 0.0);
            }
            if (f == L"PRODUCT")
            {
                double p = 1, d; bool any = false;
                for (auto* o : vals) if (asNum(*o, d)) { any = true; p *= d; }
                return ExcelObj(any ? p : 0.0);
            }
            return ExcelObj(CellError::Value);
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
            const ExcelObj* rowF = nullptr;      // 필수
            const ExcelObj* colF = nullptr;      // PIVOTBY만 (1열)
            const ExcelObj* values = nullptr;    // 필수 (GROUPBY 다열 허용)
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
            if (hasCols)
            {
                if (ckStore.nRows() != Rall) return returnValue(CellError::Value);
                // 미지원 조합: 다중 열축·다중 값·다중 행키 (plan/21 D 제한).
                if (ckStore.nCols() != 1 || V != 1 || K != 1)
                    return returnValue(CellError::Value);
            }

            const std::wstring fn = upper(in.fn->toString());
            {
                static const wchar_t* kFns[] = { L"SUM", L"AVERAGE", L"COUNT", L"COUNTA",
                                                 L"MAX", L"MIN", L"PRODUCT", L"PERCENTOF" };
                bool ok = false;
                for (auto* f : kFns) if (fn == f) { ok = true; break; }
                if (!ok) return returnValue(CellError::Value);
            }

            // relative_to (PERCENTOF): 0 총합(기본), 1 행합, 2 열합.
            const int relativeTo = readInt(in.relativeTo, 0);
            if (relativeTo < 0 || relativeTo > 2) return returnValue(CellError::Value);
            if (relativeTo != 0 && fn != L"PERCENTOF") { /* 무시 (네이티브도 PERCENTOF 외 무의미) */ }

            // ---- field_headers: 생략=자동 감지 ----
            int fieldHeaders;
            if (!in.fieldHeaders || in.fieldHeaders->isMissing())
            {
                bool firstAllText = Rall > 1;
                for (size_t c = 0; firstAllText && c < K; ++c)
                    if (rk.at(0, (ExcelArray::col_t)c).type() != ExcelType::Str) firstAllText = false;
                for (size_t c = 0; firstAllText && c < V; ++c)
                    if (vv.at(0, (ExcelArray::col_t)c).type() != ExcelType::Str) firstAllText = false;
                if (firstAllText && hasCols &&
                    ckStore.at(0, 0).type() != ExcelType::Str) firstAllText = false;
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

            // ---- total_depth ----
            int rowDepth = readInt(in.rowDepth, 1);
            const int colDepth = readInt(in.colDepth, hasCols ? 1 : 0);
            if (abs(rowDepth) > 2) return returnValue(CellError::Value);
            if (abs(colDepth) > 1) return returnValue(CellError::Value);   // 열축은 총계만
            if (abs(rowDepth) == 2 && K < 2) rowDepth = rowDepth > 0 ? 1 : -1;

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

            // ---- 행 그룹 (해시, 등장 순서) ----
            std::vector<Group> rowGroups;
            {
                std::unordered_map<std::wstring, size_t> lookup;
                lookup.reserve(dataRows.size() * 2);
                std::wstring canon;
                for (auto r : dataRows)
                {
                    canon.clear();
                    for (size_t c = 0; c < K; ++c) appendCanonical(canon, rk.at(r, (ExcelArray::col_t)c));
                    const auto it = lookup.find(canon);
                    if (it == lookup.end())
                    {
                        Group g;
                        for (size_t c = 0; c < K; ++c) g.key.emplace_back(rk.at(r, (ExcelArray::col_t)c));
                        g.rows.push_back(r);
                        lookup.emplace(canon, rowGroups.size());
                        rowGroups.push_back(std::move(g));
                    }
                    else
                        rowGroups[it->second].rows.push_back(r);
                }
            }

            // ---- 열 그룹 ----
            std::vector<Group> colGroups;
            std::unordered_map<std::wstring, size_t> colLookup;
            if (hasCols)
            {
                std::wstring canon;
                for (auto r : dataRows)
                {
                    canon.clear();
                    appendCanonical(canon, ckStore.at(r, 0));
                    const auto it = colLookup.find(canon);
                    if (it == colLookup.end())
                    {
                        Group g;
                        g.key.emplace_back(ckStore.at(r, 0));
                        g.rows.push_back(r);
                        colLookup.emplace(canon, colGroups.size());
                        colGroups.push_back(std::move(g));
                    }
                    else
                        colGroups[it->second].rows.push_back(r);
                }
            }

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
                std::unordered_map<std::wstring, size_t> rowLookup;
                {
                    std::wstring canon;
                    for (size_t g = 0; g < nG; ++g)
                    {
                        canon.clear();
                        for (const auto& k : rowGroups[g].key) appendCanonical(canon, k);
                        rowLookup.emplace(canon, g);
                    }
                }
                std::wstring canon;
                for (auto r : dataRows)
                {
                    canon.clear();
                    for (size_t c = 0; c < K; ++c) appendCanonical(canon, rk.at(r, (ExcelArray::col_t)c));
                    const size_t g = rowLookup[canon];
                    canon.clear();
                    appendCanonical(canon, ckStore.at(r, 0));
                    cross[g][colLookup[canon]].push_back(r);
                }
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

            std::vector<size_t> rowOrder(nG);
            std::iota(rowOrder.begin(), rowOrder.end(), (size_t)0);
            std::stable_sort(rowOrder.begin(), rowOrder.end(), [&](size_t x, size_t y)
            {
                if (rowSortSpec.empty())
                {
                    for (size_t c = 0; c < K; ++c)
                    {
                        const int cv = cmpObj(rowGroups[x].key[c], rowGroups[y].key[c]);
                        if (cv != 0) return cv < 0;
                    }
                    return false;
                }
                for (const auto& [idx, dir] : rowSortSpec)
                {
                    int cv = 0;
                    if (idx >= 1 && (size_t)idx <= K)
                        cv = cmpObj(rowGroups[x].key[(size_t)idx - 1],
                                    rowGroups[y].key[(size_t)idx - 1]);
                    else if (!hasCols && (size_t)idx <= K + V)
                        cv = cmpObj(groupAgg[x][(size_t)idx - K - 1],
                                    groupAgg[y][(size_t)idx - K - 1]);
                    else
                        continue;
                    if (cv != 0) return dir > 0 ? cv < 0 : cv > 0;
                }
                return false;
            });

            std::vector<size_t> colOrder(hasCols ? nC : 0);
            if (hasCols)
            {
                std::iota(colOrder.begin(), colOrder.end(), (size_t)0);
                std::stable_sort(colOrder.begin(), colOrder.end(), [&](size_t x, size_t y)
                {
                    const int cv = cmpObj(colGroups[x].key[0], colGroups[y].key[0]);
                    const int dir = colSortSpec.empty() ? 1 : colSortSpec[0].second;
                    return dir > 0 ? cv < 0 : cv > 0;
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
                // ---- PIVOTBY: (1 + nC [+합계])열 × (헤더 + nG [+합계])행 ----
                const bool wantRowGrand = rowDepth != 0;
                const bool rowAtTop = rowDepth < 0;
                const bool wantColGrand = colDepth != 0;
                const bool colAtLeft = colDepth < 0;

                // 원시 합(PERCENTOF 분모/분자)과 교차 존재 여부.
                std::vector<std::vector<double>> rawCell(nG, std::vector<double>(nC, 0.0));
                std::vector<double> rawRowTot(nG, 0.0), rawColTot(nC, 0.0);
                double rawGrand = 0.0;
                if (fn == L"PERCENTOF")
                {
                    double d;
                    for (size_t g = 0; g < nG; ++g)
                        for (size_t c = 0; c < nC; ++c)
                            for (auto r : cross[g][c])
                                if (asNum(vv.at(r, 0), d))
                                { rawCell[g][c] += d; rawRowTot[g] += d; rawColTot[c] += d; rawGrand += d; }
                }

                // 셀 값 계산. rowIdx/colIdx == SIZE_MAX 는 합계 행/열.
                // PERCENTOF 분모: relative_to 0=열합(기본), 1=행합, 2=총합.
                // 합계 열의 열합 = 총합, 합계 행의 행합 = 총합 (관측된 네이티브 규칙).
                auto pivotCell = [&](size_t g, size_t c) -> ExcelObj
                {
                    const bool grandRow = g == SIZE_MAX;
                    const bool grandCol = c == SIZE_MAX;
                    // 대상 행 집합.
                    const std::vector<ExcelArray::row_t>* rows;
                    if (grandRow && grandCol) rows = &dataRows;
                    else if (grandRow) rows = &colGroups[c].rows;
                    else if (grandCol) rows = &rowGroups[g].rows;
                    else rows = &cross[g][c];

                    if (!grandRow && !grandCol && rows->empty())
                        return ExcelObj(std::wstring_view(L""));   // 빈 교차 = 빈칸

                    if (fn != L"PERCENTOF")
                        return aggregate(fn, cellsOf(*rows, 0));

                    double num = 0, d;
                    for (auto r : *rows) if (asNum(vv.at(r, 0), d)) num += d;
                    double denom;
                    if (relativeTo == 2) denom = rawGrand;
                    else if (relativeTo == 1) denom = grandRow ? rawGrand : rawRowTot[g];
                    else denom = grandCol ? rawGrand : rawColTot[c];
                    if (denom == 0.0) return ExcelObj(CellError::Div0);
                    return ExcelObj(num / denom);
                };

                // 헤더 행 (열 키는 항상 필요).
                {
                    std::vector<ExcelObj> h;
                    h.emplace_back(std::wstring_view(showHeaders ? keyNames[0] : L""));
                    if (wantColGrand && colAtLeft) h.emplace_back(std::wstring_view(kTotal));
                    for (size_t ci = 0; ci < nC; ++ci)
                        h.emplace_back(colGroups[colOrder[ci]].key[0]);
                    if (wantColGrand && !colAtLeft) h.emplace_back(std::wstring_view(kTotal));
                    out.push_back(std::move(h));
                }

                auto emitDataRow = [&](size_t g /*실제 그룹 인덱스; SIZE_MAX=합계*/)
                {
                    std::vector<ExcelObj> row;
                    if (g == SIZE_MAX) row.emplace_back(std::wstring_view(kTotal));
                    else row.emplace_back(rowGroups[g].key[0]);
                    if (wantColGrand && colAtLeft) row.push_back(pivotCell(g, SIZE_MAX));
                    for (size_t ci = 0; ci < nC; ++ci)
                        row.push_back(pivotCell(g, colOrder[ci]));
                    if (wantColGrand && !colAtLeft) row.push_back(pivotCell(g, SIZE_MAX));
                    out.push_back(std::move(row));
                };

                if (wantRowGrand && rowAtTop) emitDataRow(SIZE_MAX);
                for (size_t gi = 0; gi < nG; ++gi) emitDataRow(rowOrder[gi]);
                if (wantRowGrand && !rowAtTop) emitDataRow(SIZE_MAX);
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
        egtools::core::registerFn(L"GROUPBY",
            [](const ExcelObj& rowF, const ExcelObj& values, const ExcelObj& fn,
               const ExcelObj& fieldHeaders, const ExcelObj& totalDepth,
               const ExcelObj& sortOrder, const ExcelObj& filterArray,
               const ExcelObj& /*fieldRelationship — 미지원, 무시*/) -> ExcelObj*
            {
                try
                {
                    PivotArgs a;
                    a.rowF = &rowF; a.values = &values; a.fn = &fn;
                    a.fieldHeaders = &fieldHeaders; a.rowDepth = &totalDepth;
                    a.rowSort = &sortOrder; a.filter = &filterArray;
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
