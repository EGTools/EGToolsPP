// FxEgData.cpp — EGTools 전용 데이터 함수 (EGToolsVB FX1_Data.vb 포팅, plan/20 E2).
//   UNPIVOT        — 집계표 → 정규 데이터 표(피벗 해제). 값 기반(병합셀 채움은 미지원).
//   BOMTREE        — 모품목/자품목/소요량 → BOM 트리 전개(+WBS). 순환 깊이 가드 추가.
//   EXPLODE        — 지정 열을 구분자로 분해해 행(세로) 또는 열(가로)로 전개.
//   COMPARELIST    — 여러 목록 비교(합집합/전용/단독/교집합/부분교집합). 가변 인수.
//   ADVANCEDFILTER — 고급필터(조건 그리드: 행=OR, 열=AND, 연산자·와일드카드).
//   MVLOOKUP       — 대량 VLOOKUP(다중 반환 열, 정확/유사 일치).
//   SUMBYBOM       — BOM 1단계 전개로 품목별 소요 수량 합산. 가변 인수.
// 전부 순수 계산(비매크로형). VISIBLEAGGR(숨김 셀)·ILOOKUP(이미지)은 매크로형 배치로 보류.

#include "../core/Registry.h"
#include "../core/ArrayUtil.h"
#include "../core/Spill.h"
#include "../core/I18n.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>

#include <algorithm>
#include <cwctype>
#include <map>
#include <numeric>
#include <set>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // ---- 공용 ----------------------------------------------------------

        bool isNumType(const ExcelObj& e)
        {
            return e.type() == ExcelType::Num || e.type() == ExcelType::Int;
        }
        bool isEmptyish(const ExcelObj& e)
        {
            switch (e.type())
            {
            case ExcelType::Missing: case ExcelType::Nil: return true;
            case ExcelType::Str: return e.toString().empty();
            default: return false;
            }
        }

        std::wstring upperCopy(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)std::towupper(c);
            return s;
        }

        // 비교/사전 키: 숫자는 정규화된 수치 표기, 빈칸은 표식, 텍스트는 대문자.
        std::wstring keyOf(const ExcelObj& e)
        {
            if (isEmptyish(e)) return L"__EMPTY__";
            if (isNumType(e))
            {
                wchar_t buf[40];
                swprintf_s(buf, L"'%.15g", e.get<double>(0.0));
                return buf;
            }
            if (e.type() == ExcelType::Bool)
                return e.get<bool>(false) ? L"'TRUE" : L"'FALSE";
            return upperCopy(e.toString());
        }

        // 정렬용 값 비교: 숫자 < 텍스트, 숫자는 수치, 텍스트는 대소문자 무시.
        int cmpVals(const ExcelObj& a, const ExcelObj& b)
        {
            const bool na = isNumType(a), nb = isNumType(b);
            if (na && nb)
            {
                const double x = a.get<double>(0.0), y = b.get<double>(0.0);
                if (x < y) return -1;
                if (x > y) return 1;
                return 0;
            }
            if (na) return -1;
            if (nb) return 1;
            return _wcsicmp(a.toString().c_str(), b.toString().c_str());
        }

        // Excel식 와일드카드(*, ?) 매칭, 대소문자 무시. 반복 O(n*m) DP-포인터 방식.
        bool wildMatch(const std::wstring& text, const std::wstring& pat)
        {
            size_t t = 0, p = 0, star = std::wstring::npos, mark = 0;
            while (t < text.size())
            {
                if (p < pat.size() &&
                    (pat[p] == L'?' || std::towupper(pat[p]) == std::towupper(text[t])))
                { ++t; ++p; }
                else if (p < pat.size() && pat[p] == L'*')
                { star = p++; mark = t; }
                else if (star != std::wstring::npos)
                { p = star + 1; t = ++mark; }
                else return false;
            }
            while (p < pat.size() && pat[p] == L'*') ++p;
            return p == pat.size();
        }

        std::wstring trimWs2(const std::wstring& s)
        {
            const wchar_t* ws = L" \t\r\n";
            const size_t b = s.find_first_not_of(ws);
            if (b == std::wstring::npos) return {};
            return s.substr(b, s.find_last_not_of(ws) - b + 1);
        }

        bool toDouble(const ExcelObj& e, double& out)
        {
            if (isEmptyish(e)) { out = 0; return true; }
            if (isNumType(e)) { out = e.get<double>(0.0); return true; }
            if (e.type() == ExcelType::Bool) { out = e.get<bool>(false) ? 1.0 : 0.0; return true; }
            if (e.type() == ExcelType::Err) return false;
            std::wstring t = trimWs2(e.toString());
            std::wstring u;
            for (wchar_t c : t) if (c != L',') u.push_back(c);
            if (u.empty()) { out = 0; return true; }
            wchar_t* end = nullptr;
            const double v = wcstod(u.c_str(), &end);
            if (!end || *end != L'\0') return false;
            out = v;
            return true;
        }

        // 2D 그리드(vector<vector<ExcelObj>>) → 배열 출력.
        ExcelObj* objGrid(const std::vector<std::vector<ExcelObj>>& grid)
        {
            if (grid.empty()) return returnValue(CellError::NA);
            size_t cols = 0;
            for (const auto& r : grid) cols = (std::max)(cols, r.size());
            if (cols == 0) return returnValue(CellError::NA);
            std::vector<ExcelObj> vals;
            vals.reserve(grid.size() * cols);
            for (const auto& r : grid)
                for (size_t c = 0; c < cols; ++c)
                    vals.push_back(c < r.size() ? ExcelObj(r[c]) : ExcelObj(std::wstring_view(L"")));
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)grid.size(), (ExcelArrayBuilder::col_t)cols, vals));
        }

        // 인수를 1D ExcelObj 목록으로 (스칼라→1개, 배열→전체 행우선).
        std::vector<ExcelObj> to1D(const ExcelObj& v)
        {
            std::vector<ExcelObj> out;
            if (v.isMissing()) return out;
            if (v.isType(ExcelType::Multi))
            {
                ExcelArray a(v);
                const size_t n = (size_t)a.nRows() * a.nCols();
                for (size_t k = 0; k < n; ++k) out.push_back(ExcelObj(a.at(k)));
            }
            else out.push_back(ExcelObj(v));
            return out;
        }

        // ---- UNPIVOT -------------------------------------------------------

        ExcelObj* unpivot(const ExcelObj& dataObj, const ExcelObj& fcObj,
                          const ExcelObj& frObj, const ExcelObj& sortObj,
                          const ExcelObj& namesObj, const ExcelObj& subsetObj)
        {
            if (!dataObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            ExcelArray a(dataObj);
            const int nR = (int)a.nRows(), nC = (int)a.nCols();
            auto cell = [&](int r, int c) -> const ExcelObj& { return a.at((size_t)r * nC + c); };

            int fixedCol = fcObj.isMissing() ? 1 : fcObj.get<int>(1);
            int fixedRow = frObj.isMissing() ? 1 : frObj.get<int>(1);
            const bool hasSubset = subsetObj.isMissing() ? false
                                                         : (subsetObj.get<double>(0.0) != 0.0);
            if (fixedCol < 0 || fixedRow < 0) return returnValue(CellError::Value);
            if (fixedRow >= nR || fixedCol >= nC) return returnValue(CellError::Value);
            if (fixedCol + fixedRow == 0) return returnValue(CellError::Value);

            std::vector<int> sortCol, sortDir;
            for (const auto& s : to1D(sortObj))
            {
                double d;
                if (!toDouble(s, d) || (int)d == 0) return returnValue(CellError::Value);
                sortCol.push_back(abs((int)d));
                sortDir.push_back((int)d > 0 ? 1 : -1);
            }
            std::vector<ExcelObj> colNames = to1D(namesObj);

            // Subset: 마지막 고정행의 서로 다른 라벨 수가 값 열 수를 나누면 집합으로.
            int iSet = 0;
            std::vector<ExcelObj> setLabels;
            if (hasSubset && fixedRow > 0)
            {
                std::vector<std::wstring> seen;
                for (int c = fixedCol; c < nC; ++c)
                {
                    const std::wstring k = keyOf(cell(fixedRow - 1, c));
                    if (std::find(seen.begin(), seen.end(), k) == seen.end())
                    {
                        seen.push_back(k);
                        setLabels.push_back(ExcelObj(cell(fixedRow - 1, c)));
                    }
                }
                const int cnt = (int)seen.size();
                if (cnt > 0 && (nC - fixedCol) >= cnt && (nC - fixedCol) % cnt == 0)
                {
                    iSet = cnt;
                    fixedRow -= 1;
                }
            }
            const int valueCols = iSet == 0 ? 1 : iSet;
            const int outCols = fixedCol + fixedRow + valueCols;

            // 제목줄: 지정값 + 기본값 병합.
            colNames.resize((size_t)outCols);
            const bool ko = egtools::i18n::current() == L"ko";
            for (int c = 0; c < fixedCol; ++c)
                if (isEmptyish(colNames[c])) colNames[c] = ExcelObj(cell(0, c));
            for (int c = 0; c < fixedRow; ++c)
                if (isEmptyish(colNames[(size_t)fixedCol + c]))
                    colNames[(size_t)fixedCol + c] = ExcelObj(std::wstring_view(
                        (ko ? L"열" : L"Column") + std::to_wstring(c + 1)));
            if (iSet == 0)
            {
                if (isEmptyish(colNames[(size_t)fixedCol + fixedRow]))
                    colNames[(size_t)fixedCol + fixedRow] =
                        ExcelObj(std::wstring_view(ko ? L"값" : L"Value"));
            }
            else
                for (int c = 0; c < iSet; ++c)
                    if (isEmptyish(colNames[(size_t)fixedCol + fixedRow + c]))
                        colNames[(size_t)fixedCol + fixedRow + c] = ExcelObj(setLabels[c]);

            // 데이터 전개.
            std::vector<std::vector<ExcelObj>> rows;
            for (int dR = fixedRow; dR < nR; ++dR)
            {
                if (dR == 0) continue;
                if (iSet > 0 && dR == fixedRow) continue;   // subset 라벨 행

                for (int dC = fixedCol; dC < nC; dC += valueCols)
                {
                    bool hasValue = false;
                    for (int i = 0; i < valueCols; ++i)
                        if (!isEmptyish(cell(dR, dC + i))) { hasValue = true; break; }
                    if (!hasValue) continue;

                    std::vector<ExcelObj> row((size_t)outCols);
                    for (int c = 0; c < fixedCol; ++c) row[c] = ExcelObj(cell(dR, c));
                    for (int r = 0; r < fixedRow; ++r)
                        row[(size_t)fixedCol + r] = ExcelObj(cell(r, dC));
                    if (iSet == 0)
                        row[(size_t)fixedCol + fixedRow] = ExcelObj(cell(dR, dC));
                    else
                        for (int i = 0; i < iSet; ++i)
                        {
                            // 라벨 위치에 맞춰 배치.
                            const std::wstring k = keyOf(cell(fixedRow, dC + i));
                            for (int m = 0; m < iSet; ++m)
                                if (keyOf(setLabels[m]) == k)
                                {
                                    row[(size_t)fixedCol + fixedRow + m] =
                                        isEmptyish(cell(dR, dC + i))
                                            ? ExcelObj(std::wstring_view(L""))
                                            : ExcelObj(cell(dR, dC + i));
                                    break;
                                }
                        }
                    rows.push_back(std::move(row));
                }
            }

            if (!sortCol.empty())
            {
                for (int c : sortCol)
                    if (c < 1 || c > outCols) return returnValue(CellError::Value);
                std::stable_sort(rows.begin(), rows.end(),
                    [&](const std::vector<ExcelObj>& x, const std::vector<ExcelObj>& y)
                    {
                        for (size_t k = 0; k < sortCol.size(); ++k)
                        {
                            const int c = cmpVals(x[(size_t)sortCol[k] - 1],
                                                  y[(size_t)sortCol[k] - 1]) * sortDir[k];
                            if (c != 0) return c < 0;
                        }
                        return false;
                    });
            }

            std::vector<std::vector<ExcelObj>> out;
            out.reserve(rows.size() + 1);
            out.push_back(std::move(colNames));
            for (auto& r : rows) out.push_back(std::move(r));
            return objGrid(out);
        }

        // ---- BOMTREE -------------------------------------------------------

        struct BomCtx
        {
            std::unordered_map<std::wstring, std::vector<std::wstring>> assy;
            std::unordered_map<std::wstring, ExcelObj> qty;   // "parent||child" → 값
            std::vector<std::vector<ExcelObj>>* out = nullptr;
            bool wbs = false;
        };

        void buildBom(BomCtx& ctx, const std::wstring& parent, int level,
                      const std::wstring& wbs)
        {
            if (level > 64) return;   // 순환/비정상 깊이 가드 (VB에는 없어 무한 재귀였음)
            const auto it = ctx.assy.find(parent);
            if (it == ctx.assy.end()) return;
            int idx = 0;
            for (const auto& child : it->second)
            {
                ++idx;
                std::wstring num = std::to_wstring(idx);
                while (num.size() < 5) num.insert(num.begin(), (wchar_t)0x200B);
                const std::wstring newWbs = wbs + L"." + num;
                const auto q = ctx.qty.find(parent + L"||" + child);
                std::vector<ExcelObj> row;
                row.push_back(ExcelObj((double)level));
                row.push_back(ExcelObj(std::wstring_view(parent)));
                row.push_back(ExcelObj(std::wstring_view(child)));
                row.push_back(q != ctx.qty.end() ? ExcelObj(q->second)
                                                 : ExcelObj(std::wstring_view(L"")));
                if (ctx.wbs) row.push_back(ExcelObj(std::wstring_view(newWbs)));
                ctx.out->push_back(std::move(row));
                buildBom(ctx, child, level + 1, newWbs);
            }
        }

        ExcelObj* bomTree(const ExcelObj& parentsObj, const ExcelObj& childrenObj,
                          const ExcelObj& valuesObj, const ExcelObj& headerObj,
                          const ExcelObj& wbsObj)
        {
            const auto parents = to1D(parentsObj);
            const auto children = to1D(childrenObj);
            const auto values = to1D(valuesObj);
            if (parents.empty() || children.empty() || values.empty())
                return returnValue(CellError::Value);
            if (parents.size() != children.size() || parents.size() != values.size())
                return returnValue(CellError::Value);
            const bool header = headerObj.isMissing() ? false : (headerObj.get<double>(0.0) != 0.0);
            const bool wbs = wbsObj.isMissing() ? false : (wbsObj.get<double>(0.0) != 0.0);

            std::unordered_set<std::wstring> childSet;
            for (const auto& c : children) childSet.insert(trimWs2(c.toString()));

            BomCtx ctx;
            ctx.wbs = wbs;
            std::vector<std::wstring> tops;
            std::unordered_set<std::wstring> topSeen;

            for (size_t i = 0; i < parents.size(); ++i)
            {
                const std::wstring parent = trimWs2(parents[i].toString());
                const std::wstring child = trimWs2(children[i].toString());
                if (parent.empty()) break;   // VB 동작: 빈 모품목에서 중단

                ExcelObj value = isEmptyish(values[i]) ? ExcelObj(0.0) : ExcelObj(values[i]);
                double d;
                if (!isNumType(value) && toDouble(value, d)) value = ExcelObj(d);

                if (childSet.find(parent) == childSet.end() && topSeen.insert(parent).second)
                    tops.push_back(parent);

                const std::wstring key = parent + L"||" + child;
                const auto it = ctx.qty.find(key);
                if (it != ctx.qty.end())
                {
                    if (cmpVals(it->second, value) != 0)
                        return returnValue(ExcelObj(std::wstring_view(
                            L"Error: " + key + L" has different values")));
                }
                else
                {
                    ctx.qty.emplace(key, ExcelObj(value));
                    ctx.assy[parent].push_back(child);
                }
            }

            std::vector<std::vector<ExcelObj>> out;
            ctx.out = &out;
            if (header)
            {
                std::vector<ExcelObj> h = { ExcelObj(std::wstring_view(L"level")),
                    ExcelObj(std::wstring_view(L"Parent")), ExcelObj(std::wstring_view(L"Child")),
                    ExcelObj(std::wstring_view(L"Value")) };
                if (wbs) h.push_back(ExcelObj(std::wstring_view(L"WBS")));
                out.push_back(std::move(h));
            }
            int cnt = 0;
            for (const auto& top : tops)
            {
                ++cnt;
                std::wstring num = std::to_wstring(cnt);
                while (num.size() < 5) num.insert(num.begin(), (wchar_t)0x200B);
                std::vector<ExcelObj> row = { ExcelObj(0.0),
                    ExcelObj(std::wstring_view(L"Top Level")),
                    ExcelObj(std::wstring_view(top)), ExcelObj(std::wstring_view(L"")) };
                if (wbs) row.push_back(ExcelObj(std::wstring_view(num)));
                out.push_back(std::move(row));
                buildBom(ctx, top, 1, num);
            }
            return objGrid(out);
        }

        // ---- EXPLODE -------------------------------------------------------

        std::vector<std::wstring> splitByDelims(std::wstring s,
                                                const std::vector<std::wstring>& delims,
                                                bool ignoreEmpty)
        {
            // 각 구분자를 \x07로 치환한 뒤 분리 (긴 구분자 우선).
            for (const auto& d : delims)
            {
                if (d.empty()) continue;
                size_t p = 0;
                while ((p = s.find(d, p)) != std::wstring::npos)
                {
                    s.replace(p, d.size(), 1, (wchar_t)7);
                    ++p;
                }
            }
            std::vector<std::wstring> parts;
            std::wstring cur;
            for (wchar_t c : s)
            {
                if (c == (wchar_t)7) { parts.push_back(cur); cur.clear(); }
                else cur.push_back(c);
            }
            parts.push_back(cur);
            if (ignoreEmpty)
            {
                std::vector<std::wstring> f;
                for (auto& p : parts) if (!p.empty()) f.push_back(std::move(p));
                if (f.empty()) f.push_back(L"");
                return f;
            }
            return parts;
        }

        ExcelObj* explode(const ExcelObj& arrObj, const ExcelObj& colObj,
                          const ExcelObj& delimObj, const ExcelObj& byColObj,
                          const ExcelObj& padObj, const ExcelObj& ignoreObj)
        {
            if (!arrObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            ExcelArray a(arrObj);
            const int nR = (int)a.nRows(), nC = (int)a.nCols();
            auto cell = [&](int r, int c) -> const ExcelObj& { return a.at((size_t)r * nC + c); };

            const int colNum = colObj.isMissing() ? 0 : colObj.get<int>(0);
            if (colNum < 1 || colNum > nC) return returnValue(CellError::Num);
            const bool byCol = byColObj.isMissing() ? false : (byColObj.get<double>(0.0) != 0.0);
            const bool ignoreEmpty = ignoreObj.isMissing() ? false
                                                           : (ignoreObj.get<double>(0.0) != 0.0);
            // 배열 채움값은 결과 셀에 그대로 복사되어 조용히 #VALUE!로 오염됨 — 거부.
            if (padObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            const ExcelObj pad = padObj.isMissing() ? ExcelObj(std::wstring_view(L""))
                                                    : ExcelObj(padObj);

            std::vector<std::wstring> delims;
            for (const auto& d : to1D(delimObj)) delims.push_back(d.toString());
            if (delims.empty()) delims.push_back(L" ");
            std::sort(delims.begin(), delims.end(),
                      [](const std::wstring& x, const std::wstring& y)
                      { return x.size() > y.size(); });

            std::vector<std::vector<std::wstring>> parts((size_t)nR);
            size_t maxSplit = 1;
            for (int r = 0; r < nR; ++r)
            {
                parts[r] = splitByDelims(cell(r, colNum - 1).toString(), delims, ignoreEmpty);
                maxSplit = (std::max)(maxSplit, parts[r].size());
            }

            std::vector<std::vector<ExcelObj>> out;
            if (byCol)
            {
                for (int r = 0; r < nR; ++r)
                {
                    std::vector<ExcelObj> row;
                    row.reserve((size_t)nC + maxSplit - 1);
                    for (int c = 0; c < colNum - 1; ++c) row.push_back(ExcelObj(cell(r, c)));
                    for (size_t i = 0; i < maxSplit; ++i)
                        row.push_back(i < parts[r].size()
                            ? ExcelObj(std::wstring_view(parts[r][i])) : ExcelObj(pad));
                    for (int c = colNum; c < nC; ++c) row.push_back(ExcelObj(cell(r, c)));
                    out.push_back(std::move(row));
                }
            }
            else
            {
                for (int r = 0; r < nR; ++r)
                    for (const auto& p : parts[r])
                    {
                        std::vector<ExcelObj> row;
                        row.reserve((size_t)nC);
                        for (int c = 0; c < nC; ++c)
                            row.push_back(c == colNum - 1 ? ExcelObj(std::wstring_view(p))
                                                          : ExcelObj(cell(r, c)));
                        out.push_back(std::move(row));
                    }
            }
            return objGrid(out);
        }

        // ---- COMPARELIST (가변 인수) ---------------------------------------

        ExcelObj* compareList(const FuncInfo& info, const ExcelObj** args)
        {
            const size_t total = (size_t)info.numArgs();
            if (total < 3 || args[0]->isMissing() || args[1]->isMissing())
                return returnValue(CellError::Value);
            const int uniqueType = args[0]->get<int>(0);
            const bool byRow = args[1]->get<double>(0.0) != 0.0;

            // 목록 수집.
            std::vector<std::vector<std::vector<ExcelObj>>> lists;   // list → rows → cells
            for (size_t i = 2; i < total; ++i)
            {
                if (args[i]->isMissing()) continue;
                std::vector<std::vector<ExcelObj>> rows;
                if (args[i]->isType(ExcelType::Multi))
                {
                    ExcelArray a(*args[i]);
                    const size_t R = a.nRows(), C = a.nCols();
                    if (byRow)
                        for (size_t r = 0; r < R; ++r)
                        {
                            std::vector<ExcelObj> row;
                            for (size_t c = 0; c < C; ++c) row.push_back(ExcelObj(a.at(r * C + c)));
                            rows.push_back(std::move(row));
                        }
                    else
                        for (size_t k = 0; k < R * C; ++k)
                            rows.push_back({ ExcelObj(a.at(k)) });
                }
                else
                    rows.push_back({ ExcelObj(*args[i]) });
                lists.push_back(std::move(rows));
            }
            const int n = (int)lists.size();
            if (n == 0) return returnValue(CellError::Value);
            if (uniqueType < -3 || uniqueType > n) return returnValue(CellError::Value);

            // 키 → (등장 목록 집합, 대표 행). 삽입 순서 유지.
            std::vector<std::wstring> order;
            std::unordered_map<std::wstring, std::set<int>> members;
            std::unordered_map<std::wstring, std::vector<ExcelObj>> rep;

            for (int li = 0; li < n; ++li)
                for (auto& row : lists[li])
                {
                    std::wstring key;
                    for (const auto& cell : row)
                    {
                        if (!key.empty()) key += L'|';
                        key += keyOf(cell);
                    }
                    // 우측 빈칸 제거(열 수만 다른 동일 행 구별 방지).
                    static const std::wstring kTail = L"|__EMPTY__";
                    while (key.size() >= kTail.size() &&
                           key.compare(key.size() - kTail.size(), kTail.size(), kTail) == 0)
                        key.erase(key.size() - kTail.size());

                    if (members.find(key) == members.end())
                    {
                        order.push_back(key);
                        rep[key] = row;
                    }
                    members[key].insert(li);
                }

            std::vector<std::wstring> picked;
            for (const auto& key : order)
            {
                const auto& m = members[key];
                bool ok = false;
                if (uniqueType == 0) ok = true;
                else if (uniqueType >= 1)
                    ok = m.size() == 1 && m.count(uniqueType - 1) > 0;
                else if (uniqueType == -1) ok = m.size() == 1;
                else if (uniqueType == -2) ok = (int)m.size() == n;
                else if (uniqueType == -3) ok = m.size() > 1 && (int)m.size() < n;
                if (ok) picked.push_back(key);
            }
            if (picked.empty()) return returnValue(CellError::NA);

            std::vector<std::vector<ExcelObj>> out;
            for (const auto& key : picked)
            {
                auto row = rep[key];
                for (auto& cell : row)
                    if (isEmptyish(cell)) cell = ExcelObj(std::wstring_view(L""));
                out.push_back(std::move(row));
            }
            return objGrid(out);
        }

        // ---- ADVANCEDFILTER ------------------------------------------------

        bool filterPass(const ExcelObj& dataVal, const ExcelObj& filterVal)
        {
            if (isEmptyish(filterVal)) return true;
            std::wstring f = filterVal.toString();
            // 선행 공백 제거 후 연산자 파싱.
            const size_t b = f.find_first_not_of(L" \t");
            if (b == std::wstring::npos) return true;
            f = f.substr(b);
            static const wchar_t* kOps[] = { L">=", L"<=", L"<>", L"=", L">", L"<" };
            std::wstring op;
            for (auto* o : kOps)
                if (f.compare(0, wcslen(o), o) == 0) { op = o; f = f.substr(wcslen(o)); break; }

            // 숫자 비교(데이터가 숫자이고 조건값도 숫자로 해석되면).
            double dv, fv;
            const bool dNum = isNumType(dataVal) || (dataVal.type() == ExcelType::Bool);
            if (dNum && toDouble(dataVal, dv) && !f.empty() && toDouble(ExcelObj(std::wstring_view(f)), fv))
            {
                if (op == L"=" || op.empty()) return dv == fv;
                if (op == L"<>") return dv != fv;
                if (op == L">") return dv > fv;
                if (op == L"<") return dv < fv;
                if (op == L">=") return dv >= fv;
                if (op == L"<=") return dv <= fv;
                return dv == fv;
            }

            // 문자열 비교 (와일드카드는 =/<>만, 그 외는 사전순).
            const std::wstring ds = dataVal.isType(ExcelType::Nil) ? L"" : dataVal.toString();
            const bool hasWild = f.find(L'*') != std::wstring::npos ||
                                 f.find(L'?') != std::wstring::npos;
            auto starts = [&]() {
                if (f.size() > ds.size()) return false;
                return _wcsnicmp(ds.c_str(), f.c_str(), f.size()) == 0;
            };
            if (op == L"<>")
            {
                if (hasWild) return !wildMatch(ds, f);
                if (f.empty()) return !ds.empty();
                return !starts();
            }
            if (op.empty() || op == L"=")
            {
                if (hasWild) return wildMatch(ds, f);
                if (f.empty()) return ds.empty();
                return starts();
            }
            const int c = _wcsicmp(ds.c_str(), f.c_str());
            if (op == L">") return c > 0;
            if (op == L">=") return c >= 0;
            if (op == L"<") return c < 0;
            if (op == L"<=") return c <= 0;
            return false;
        }

        ExcelObj* advancedFilter(const ExcelObj& dataObj, const ExcelObj& filterObj,
                                 const ExcelObj& uniqueObj)
        {
            if (!dataObj.isType(ExcelType::Multi) || !filterObj.isType(ExcelType::Multi))
                return returnValue(CellError::Value);
            ExcelArray d(dataObj), flt(filterObj);
            const int dR = (int)d.nRows(), dC = (int)d.nCols();
            const int fR = (int)flt.nRows(), fC = (int)flt.nCols();
            if (dR < 2 || dC == 0 || fR < 2 || fC == 0) return returnValue(CellError::Value);
            auto dCell = [&](int r, int c) -> const ExcelObj& { return d.at((size_t)r * dC + c); };
            auto fCell = [&](int r, int c) -> const ExcelObj& { return flt.at((size_t)r * fC + c); };
            const bool unique = uniqueObj.isMissing() ? false : (uniqueObj.get<double>(0.0) != 0.0);

            // 필터 헤더 → 데이터 열 매핑.
            std::vector<int> colMap((size_t)fC, -1);
            for (int c = 0; c < fC; ++c)
            {
                if (isEmptyish(fCell(0, c))) continue;
                const std::wstring name = keyOf(fCell(0, c));
                for (int k = 0; k < dC; ++k)
                    if (keyOf(dCell(0, k)) == name) { colMap[c] = k; break; }
                if (colMap[c] < 0) return returnValue(CellError::Value);
            }

            std::vector<std::vector<ExcelObj>> out;
            std::unordered_set<std::wstring> seen;
            auto addRow = [&](int r)
            {
                std::vector<ExcelObj> row;
                std::wstring key;
                for (int c = 0; c < dC; ++c)
                {
                    row.push_back(isEmptyish(dCell(r, c)) ? ExcelObj(std::wstring_view(L""))
                                                          : ExcelObj(dCell(r, c)));
                    if (unique) { key += keyOf(dCell(r, c)); key += L'|'; }
                }
                if (unique && !seen.insert(key).second) return;
                out.push_back(std::move(row));
            };

            addRow(0);   // 헤더 행 포함 (VB 동작)
            for (int r = 1; r < dR; ++r)
            {
                bool rowPass = false;
                for (int fr = 1; fr < fR && !rowPass; ++fr)
                {
                    bool condPass = true;
                    for (int fc = 0; fc < fC; ++fc)
                    {
                        if (colMap[fc] < 0) continue;
                        if (!filterPass(dCell(r, colMap[fc]), fCell(fr, fc)))
                        { condPass = false; break; }
                    }
                    if (condPass) rowPass = true;
                }
                if (rowPass) addRow(r);
            }
            if (out.size() <= 1) return returnValue(CellError::NA);
            return objGrid(out);
        }

        // ---- MVLOOKUP ------------------------------------------------------

        ExcelObj* mvLookup(const ExcelObj& keysObj, const ExcelObj& dataObj,
                           const ExcelObj& colsObj, const ExcelObj& modeObj,
                           const ExcelObj& nfObj)
        {
            if (keysObj.isMissing() || dataObj.isMissing() || colsObj.isMissing())
                return returnValue(CellError::Value);
            if (!dataObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            ExcelArray d(dataObj);
            const size_t dR = d.nRows(), dC = d.nCols();
            if (dR == 0 || dC == 0) return returnValue(CellError::Value);
            auto dCell = [&](size_t r, size_t c) -> const ExcelObj& { return d.at(r * dC + c); };

            std::vector<ExcelObj> keys;
            if (keysObj.isType(ExcelType::Multi))
            {
                // 세로/가로/2D 검색값 모두 허용(행 우선 평탄화) — 결과는 검색값
                // 개수만큼 세로 스택. 기존에는 가로 1×N을 거부했다(plan/22 A-P1).
                ExcelArray k(keysObj);
                const size_t n = (size_t)k.nRows() * k.nCols();
                for (size_t i = 0; i < n; ++i) keys.push_back(ExcelObj(k.at(i)));
            }
            else keys.push_back(ExcelObj(keysObj));

            std::vector<int> cols;
            for (const auto& c : to1D(colsObj))
            {
                double v;
                if (!toDouble(c, v) || (int)v < 1 || (size_t)(int)v > dC)
                    return returnValue(CellError::Value);
                cols.push_back((int)v);
            }
            if (cols.empty()) return returnValue(CellError::Value);
            const bool approx = modeObj.isMissing() ? false : (modeObj.get<double>(0.0) != 0.0);
            // 배열 없을때값은 미일치 셀을 조용히 #VALUE!로 오염시킴 — 거부.
            if (nfObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            const ExcelObj notFound = nfObj.isMissing() ? ExcelObj(std::wstring_view(L""))
                                                        : ExcelObj(nfObj);

            std::vector<std::vector<ExcelObj>> out;
            out.reserve(keys.size());

            if (!approx)
            {
                // 정확 일치: 해시(첫 등장 행).
                std::unordered_map<std::wstring, size_t> map;
                map.reserve(dR * 2);
                for (size_t r = 0; r < dR; ++r)
                    map.emplace(keyOf(dCell(r, 0)), r);
                for (const auto& key : keys)
                {
                    const auto it = map.find(keyOf(key));
                    std::vector<ExcelObj> row;
                    for (int c : cols)
                        row.push_back(it == map.end() ? ExcelObj(notFound)
                                                      : ExcelObj(dCell(it->second, (size_t)c - 1)));
                    out.push_back(std::move(row));
                }
            }
            else
            {
                // 유사 일치: 키 열 기준 안정 정렬 후 이진 탐색(작거나 같은 마지막).
                std::vector<size_t> idx(dR);
                std::iota(idx.begin(), idx.end(), (size_t)0);
                std::stable_sort(idx.begin(), idx.end(), [&](size_t x, size_t y)
                                 { return cmpVals(dCell(x, 0), dCell(y, 0)) < 0; });
                for (const auto& key : keys)
                {
                    size_t lo = 0, hi = dR;   // [lo,hi): key보다 큰 첫 위치 탐색
                    while (lo < hi)
                    {
                        const size_t mid = (lo + hi) / 2;
                        if (cmpVals(dCell(idx[mid], 0), key) <= 0) lo = mid + 1;
                        else hi = mid;
                    }
                    std::vector<ExcelObj> row;
                    for (int c : cols)
                        row.push_back(lo == 0 ? ExcelObj(notFound)
                                              : ExcelObj(dCell(idx[lo - 1], (size_t)c - 1)));
                    out.push_back(std::move(row));
                }
            }
            return objGrid(out);
        }

        // ---- SUMBYBOM (가변 인수) ------------------------------------------

        ExcelObj* sumByBom(const FuncInfo& info, const ExcelObj** args)
        {
            const size_t total = (size_t)info.numArgs();

            // 모든 입력을 열 단위로 평탄화.
            std::vector<std::vector<ExcelObj>> colBuckets;
            long bomRows = -1, qtyRows = -1;
            for (size_t i = 0; i < total; ++i)
            {
                if (args[i]->isMissing()) break;
                std::vector<std::vector<ExcelObj>> colsHere;
                if (args[i]->isType(ExcelType::Multi))
                {
                    ExcelArray a(*args[i]);
                    const size_t R = a.nRows(), C = a.nCols();
                    colsHere.resize(C);
                    for (size_t c = 0; c < C; ++c)
                        for (size_t r = 0; r < R; ++r)
                            colsHere[c].push_back(ExcelObj(a.at(r * C + c)));
                }
                else
                    colsHere.push_back({ ExcelObj(*args[i]) });

                for (auto& col : colsHere)
                {
                    const long len = (long)col.size();
                    if (colBuckets.size() < 3)
                    {
                        if (bomRows == -1) bomRows = len;
                        else if (bomRows != len) return returnValue(CellError::Value);
                    }
                    else
                    {
                        if (qtyRows == -1) qtyRows = len;
                        else if (qtyRows != len) return returnValue(CellError::Value);
                    }
                    colBuckets.push_back(std::move(col));
                }
            }
            if (colBuckets.size() < 5 || bomRows < 0 || qtyRows < 0)
                return returnValue(CellError::Value);

            const auto& parentCol = colBuckets[0];
            const auto& childCol = colBuckets[1];
            const auto& usageCol = colBuckets[2];
            const auto& nameCol = colBuckets[3];
            const size_t qtyCount = colBuckets.size() - 4;
            const long rowCount = (std::min)(bomRows, qtyRows);

            // BOM 관계.
            std::unordered_map<std::wstring, std::vector<std::pair<std::wstring, double>>> bom;
            std::unordered_set<std::wstring> parentItems;
            for (long i = 0; i < bomRows; ++i)
            {
                const std::wstring p = upperCopy(trimWs2(parentCol[i].toString()));
                const std::wstring c = trimWs2(childCol[i].toString());
                if (p.empty() || c.empty()) continue;
                double usage;
                if (!toDouble(usageCol[i], usage)) return returnValue(CellError::Value);
                bom[p].emplace_back(c, usage);
                parentItems.insert(p);
            }

            // 집계: key(대문자) → 수량 벡터, 표시 이름.
            std::vector<std::wstring> orderKeys;
            std::unordered_map<std::wstring, std::vector<double>> totals;
            std::unordered_map<std::wstring, std::wstring> display;
            auto add = [&](const std::wstring& name, size_t q, double amount)
            {
                const std::wstring key = upperCopy(name);
                auto it = totals.find(key);
                if (it == totals.end())
                {
                    orderKeys.push_back(key);
                    it = totals.emplace(key, std::vector<double>(qtyCount, 0.0)).first;
                    display[key] = name;
                }
                it->second[q] += amount;
            };

            for (long i = 0; i < rowCount; ++i)
            {
                const std::wstring name = trimWs2(nameCol[i].toString());
                if (name.empty()) continue;
                const std::wstring nameKey = upperCopy(name);
                if (parentItems.count(nameKey))
                {
                    const auto it = bom.find(nameKey);
                    if (it == bom.end()) continue;
                    for (const auto& [child, usage] : it->second)
                        for (size_t q = 0; q < qtyCount; ++q)
                        {
                            double qty;
                            if (!toDouble(colBuckets[4 + q][i], qty))
                                return returnValue(CellError::Value);
                            if (qty < 0) return returnValue(CellError::Num);
                            add(child, q, qty * usage);
                        }
                }
                else
                    for (size_t q = 0; q < qtyCount; ++q)
                    {
                        double qty;
                        if (!toDouble(colBuckets[4 + q][i], qty))
                            return returnValue(CellError::Value);
                        if (qty < 0) return returnValue(CellError::Num);
                        if (qty > 0) add(name, q, qty);
                    }
            }
            if (orderKeys.empty()) return returnValue(CellError::NA);

            std::sort(orderKeys.begin(), orderKeys.end(),
                      [](const std::wstring& x, const std::wstring& y)
                      { return _wcsicmp(x.c_str(), y.c_str()) < 0; });

            std::vector<std::vector<ExcelObj>> out;
            for (const auto& key : orderKeys)
            {
                std::vector<ExcelObj> row;
                row.push_back(ExcelObj(std::wstring_view(display[key])));
                for (double v : totals[key]) row.push_back(ExcelObj(v));
                out.push_back(std::move(row));
            }
            return objGrid(out);
        }
    }

    void registerEgData()
    {
        egtools::core::registerFn(L"UNPIVOT",
            [](const ExcelObj& d, const ExcelObj& fc, const ExcelObj& fr,
               const ExcelObj& so, const ExcelObj& cn, const ExcelObj& hs) -> ExcelObj*
            { return unpivot(d, fc, fr, so, cn, hs); });

        egtools::core::registerFn(L"BOMTREE",
            [](const ExcelObj& p, const ExcelObj& c, const ExcelObj& v,
               const ExcelObj& h, const ExcelObj& w) -> ExcelObj*
            { return bomTree(p, c, v, h, w); });

        egtools::core::registerFn(L"EXPLODE",
            [](const ExcelObj& a, const ExcelObj& col, const ExcelObj& del,
               const ExcelObj& bc, const ExcelObj& pad, const ExcelObj& ie) -> ExcelObj*
            { return explode(a, col, del, bc, pad, ie); });

        egtools::core::registerRawFn(L"COMPARELIST",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                try { return compareList(info, args); }
                catch (...) { return returnValue(CellError::Value); }
            });

        egtools::core::registerFn(L"ADVANCEDFILTER",
            [](const ExcelObj& d, const ExcelObj& f, const ExcelObj& u) -> ExcelObj*
            { return advancedFilter(d, f, u); });

        egtools::core::registerFn(L"MVLOOKUP",
            [](const ExcelObj& k, const ExcelObj& d, const ExcelObj& c,
               const ExcelObj& m, const ExcelObj& nf) -> ExcelObj*
            { return mvLookup(k, d, c, m, nf); });

        egtools::core::registerRawFn(L"SUMBYBOM",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                try { return sumByBom(info, args); }
                catch (...) { return returnValue(CellError::Value); }
            });
    }
}
