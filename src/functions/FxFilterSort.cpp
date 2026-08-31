// FxFilterSort.cpp — dynamic-array lookup/shape functions
// (XMATCH, FILTER, SORT, SORTBY, UNIQUE). All array-in → array-out.

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "RegexMatchMode.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <algorithm>
#include <numeric>
#include <vector>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        bool truthyCell(const ExcelObj& v)
        {
            auto t = v.type();
            return (t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool)
                && v.get<double>(0.0) != 0.0;
        }

        // --- XMATCH(lookup, lookup_array, [match_mode], [search_mode]) -------
        // match_mode 0/-1/1/3(정규식); search_mode 1/-1.
        ExcelObj* xmatch(const ExcelObj& lookup, const ExcelObj& arr,
                         const ExcelObj& matchMode, const ExcelObj& searchMode)
        {
            try
            {
                ExcelArray la(arr);
                const bool vertical = (la.nCols() == 1);
                const size_t N = vertical ? la.nRows() : la.nCols();
                if (N == 0) return returnValue(CellError::NA);
                const int mm = matchMode.isMissing() ? 0 : matchMode.get<int>(0);
                const int sm = searchMode.isMissing() ? 1 : searchMode.get<int>(1);

                auto elem = [&](size_t i) -> const ExcelObj& {
                    return vertical ? la.at((ExcelArray::row_t)i, 0)
                                    : la.at(0, (ExcelArray::col_t)i);
                };
                if (mm == 3)
                {
                    // match_mode 3: lookup은 정규식 패턴 — 텍스트가 아니거나
                    // 이진 검색(sm=±2)과 조합이면 #VALUE!(네이티브 정합).
                    if (sm == 2 || sm == -2 || lookup.type() != ExcelType::Str)
                        return returnValue(CellError::Value);
                    const auto re = regexForLookup(lookup.toString());
                    if (sm == -1)
                    {
                        for (long long i = (long long)N - 1; i >= 0; --i)
                            if (regexCellMatch(elem((size_t)i), re))
                                return returnValue(ExcelObj((double)(i + 1)));
                    }
                    else
                    {
                        for (size_t i = 0; i < N; ++i)
                            if (regexCellMatch(elem(i), re))
                                return returnValue(ExcelObj((double)(i + 1)));
                    }
                    return returnValue(CellError::NA);
                }
                long long found = -1, approx = -1;
                auto consider = [&](size_t i) -> bool {
                    int c = ExcelObj::compare(elem(i), lookup);
                    if (c == 0) { found = (long long)i; return true; }
                    if (mm == -1 && c < 0) { if (approx < 0 || ExcelObj::compare(elem((size_t)approx), elem(i)) < 0) approx = (long long)i; }
                    else if (mm == 1 && c > 0) { if (approx < 0 || ExcelObj::compare(elem(i), elem((size_t)approx)) < 0) approx = (long long)i; }
                    return false;
                };
                if (sm == -1) for (long long i = (long long)N - 1; i >= 0; --i) { if (consider((size_t)i)) break; }
                else          for (size_t i = 0; i < N; ++i) { if (consider(i)) break; }
                if (found < 0 && (mm == -1 || mm == 1)) found = approx;
                if (found < 0) return returnValue(CellError::NA);
                return returnValue(ExcelObj((double)(found + 1)));
            }
            catch (...) { return returnValue(CellError::Value); }
        }

        // --- FILTER(array, include, [if_empty]) -----------------------------
        ExcelObj* filterFn(const ExcelObj& array, const ExcelObj& include, const ExcelObj& ifEmpty)
        {
            try
            {
                ExcelArray a(array), inc(include);
                const ExcelArray::row_t R = a.nRows();
                const ExcelArray::col_t C = a.nCols();

                if (inc.nCols() == 1 && inc.nRows() == R)          // filter rows
                {
                    std::vector<ExcelArray::row_t> keep;
                    for (ExcelArray::row_t i = 0; i < R; ++i)
                    {
                        // include의 오류는 전체 결과로 전파(네이티브 실측).
                        if (inc.at(i, 0).type() == ExcelType::Err)
                            return returnValue(ExcelObj(inc.at(i, 0)));
                        if (truthyCell(inc.at(i, 0))) keep.push_back(i);
                    }
                    if (keep.empty())
                        return ifEmpty.isMissing() ? returnValue(CellError::NA)
                                                   : returnValue(ExcelObj(ifEmpty));
                    std::vector<ExcelObj> vals;
                    for (auto i : keep) for (ExcelArray::col_t j = 0; j < C; ++j) vals.emplace_back(a.at(i, j));
                    return egtools::core::output(egtools::core::makeArray((ExcelArrayBuilder::row_t)keep.size(), C, vals));
                }
                if (inc.nRows() == 1 && inc.nCols() == C)          // filter columns
                {
                    std::vector<ExcelArray::col_t> keep;
                    for (ExcelArray::col_t j = 0; j < C; ++j)
                    {
                        if (inc.at(0, j).type() == ExcelType::Err)
                            return returnValue(ExcelObj(inc.at(0, j)));
                        if (truthyCell(inc.at(0, j))) keep.push_back(j);
                    }
                    if (keep.empty())
                        return ifEmpty.isMissing() ? returnValue(CellError::NA)
                                                   : returnValue(ExcelObj(ifEmpty));
                    std::vector<ExcelObj> vals;
                    for (ExcelArray::row_t i = 0; i < R; ++i) for (auto j : keep) vals.emplace_back(a.at(i, j));
                    return egtools::core::output(egtools::core::makeArray(R, (ExcelArrayBuilder::col_t)keep.size(), vals));
                }
                return returnValue(CellError::Value);
            }
            catch (...) { return returnValue(CellError::Value); }
        }

        // --- SORT(array, [sort_index], [sort_order], [by_col]) --------------
        ExcelObj* sortFn(const ExcelObj& array, const ExcelObj& idxA,
                         const ExcelObj& orderA, const ExcelObj& byColA)
        {
            try
            {
                ExcelArray a(array);
                const ExcelArray::row_t R = a.nRows();
                const ExcelArray::col_t C = a.nCols();
                const bool byCol = byColA.isMissing() ? false : (byColA.get<double>(0.0) != 0.0);

                // sort_index·sort_order는 배열 허용 = 다중 키 우선순위 정렬
                // (네이티브 정합, plan/22 T3). order만 배열이면 #VALUE!(T9);
                // order가 스칼라면 전 키에 적용.
                auto readIntList = [](const ExcelObj& o, int def, std::vector<int>& out)
                {
                    if (o.isMissing()) { out.push_back(def); return true; }
                    if (o.isType(ExcelType::Multi))
                    {
                        ExcelArray v(o);
                        const size_t n = (size_t)v.nRows() * v.nCols();
                        if (n == 0) return false;
                        for (size_t i = 0; i < n; ++i) out.push_back(v.at(i).get<int>(def));
                        return true;
                    }
                    out.push_back(o.get<int>(def));
                    return true;
                };
                std::vector<int> idxs, orders;
                if (!readIntList(idxA, 1, idxs) || !readIntList(orderA, 1, orders))
                    return returnValue(CellError::Value);
                if (orders.size() > idxs.size()) return returnValue(CellError::Value);
                if (orders.size() < idxs.size()) orders.resize(idxs.size(), orders[0]);

                if (!byCol)
                {
                    for (int ix : idxs)
                        if (ix < 1 || (ExcelArray::col_t)ix > C) return returnValue(CellError::Value);
                    std::vector<ExcelArray::row_t> ord(R);
                    std::iota(ord.begin(), ord.end(), (ExcelArray::row_t)0);
                    std::stable_sort(ord.begin(), ord.end(), [&](auto r1, auto r2) {
                        for (size_t k = 0; k < idxs.size(); ++k)
                        {
                            int c = ExcelObj::compare(a.at(r1, idxs[k] - 1), a.at(r2, idxs[k] - 1));
                            if (c != 0) return orders[k] < 0 ? c > 0 : c < 0;
                        }
                        return false;
                    });
                    std::vector<ExcelObj> vals;
                    for (auto r : ord) for (ExcelArray::col_t j = 0; j < C; ++j) vals.emplace_back(a.at(r, j));
                    return egtools::core::output(egtools::core::makeArray(R, C, vals));
                }
                else
                {
                    for (int ix : idxs)
                        if (ix < 1 || (ExcelArray::row_t)ix > R) return returnValue(CellError::Value);
                    std::vector<ExcelArray::col_t> ord(C);
                    std::iota(ord.begin(), ord.end(), (ExcelArray::col_t)0);
                    std::stable_sort(ord.begin(), ord.end(), [&](auto c1, auto c2) {
                        for (size_t k = 0; k < idxs.size(); ++k)
                        {
                            int c = ExcelObj::compare(a.at(idxs[k] - 1, c1), a.at(idxs[k] - 1, c2));
                            if (c != 0) return orders[k] < 0 ? c > 0 : c < 0;
                        }
                        return false;
                    });
                    std::vector<ExcelObj> vals;
                    for (ExcelArray::row_t i = 0; i < R; ++i) for (auto c : ord) vals.emplace_back(a.at(i, c));
                    return egtools::core::output(egtools::core::makeArray(R, C, vals));
                }
            }
            catch (...) { return returnValue(CellError::Value); }
        }

        // --- SORTBY(array, by1, [order1], [by2], [order2]) ------------------
        // SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], …) —
        // variadic key pairs. `args[0]` is the array; args[1..] are by/order pairs.
        ExcelObj* sortByFn(const ExcelObj** args, size_t n)
        {
            try
            {
                if (n < 2) return returnValue(CellError::Value);
                ExcelArray a(*args[0]);
                const ExcelArray::row_t R = a.nRows();
                const ExcelArray::col_t C = a.nCols();
                struct Key { ExcelArray arr; int order; };
                std::vector<Key> keys;
                for (size_t i = 1; i < n; i += 2)
                {
                    const ExcelObj& b = *args[i];
                    if (b.isMissing()) continue;
                    int order = 1;
                    if (i + 1 < n && !args[i + 1]->isMissing()) order = args[i + 1]->get<int>(1);
                    ExcelArray ba(b);
                    if (ba.nRows() == R) keys.push_back({ ba, order });
                }
                if (keys.empty()) return returnValue(CellError::Value);

                std::vector<ExcelArray::row_t> ord(R);
                std::iota(ord.begin(), ord.end(), (ExcelArray::row_t)0);
                std::stable_sort(ord.begin(), ord.end(), [&](auto r1, auto r2) {
                    for (auto& k : keys)
                    {
                        int c = ExcelObj::compare(k.arr.at(r1, 0), k.arr.at(r2, 0));
                        if (c != 0) return k.order < 0 ? c > 0 : c < 0;
                    }
                    return false;
                });
                std::vector<ExcelObj> vals;
                for (auto r : ord) for (ExcelArray::col_t j = 0; j < C; ++j) vals.emplace_back(a.at(r, j));
                return egtools::core::output(egtools::core::makeArray(R, C, vals));
            }
            catch (...) { return returnValue(CellError::Value); }
        }

        // --- UNIQUE(array, [by_col], [exactly_once]) ------------------------
        ExcelObj* uniqueFn(const ExcelObj& array, const ExcelObj& byColA, const ExcelObj& onceA)
        {
            try
            {
                ExcelArray a(array);
                const ExcelArray::row_t R = a.nRows();
                const ExcelArray::col_t C = a.nCols();
                const bool byCol = byColA.isMissing() ? false : (byColA.get<double>(0.0) != 0.0);
                const bool once = onceA.isMissing() ? false : (onceA.get<double>(0.0) != 0.0);

                auto rowsEqual = [&](ExcelArray::row_t r1, ExcelArray::row_t r2) {
                    for (ExcelArray::col_t j = 0; j < C; ++j)
                        if (ExcelObj::compare(a.at(r1, j), a.at(r2, j)) != 0) return false;
                    return true;
                };
                auto colsEqual = [&](ExcelArray::col_t c1, ExcelArray::col_t c2) {
                    for (ExcelArray::row_t i = 0; i < R; ++i)
                        if (ExcelObj::compare(a.at(i, c1), a.at(i, c2)) != 0) return false;
                    return true;
                };

                if (!byCol)
                {
                    std::vector<ExcelArray::row_t> uniq;
                    for (ExcelArray::row_t i = 0; i < R; ++i)
                    {
                        long count = 0; for (ExcelArray::row_t k = 0; k < R; ++k) if (rowsEqual(i, k)) ++count;
                        bool firstSeen = true; for (ExcelArray::row_t k = 0; k < i; ++k) if (rowsEqual(i, k)) { firstSeen = false; break; }
                        if (firstSeen && (!once || count == 1)) uniq.push_back(i);
                    }
                    if (uniq.empty()) return returnValue(CellError::NA);
                    std::vector<ExcelObj> vals;
                    for (auto i : uniq) for (ExcelArray::col_t j = 0; j < C; ++j) vals.emplace_back(a.at(i, j));
                    return egtools::core::output(egtools::core::makeArray((ExcelArrayBuilder::row_t)uniq.size(), C, vals));
                }
                else
                {
                    std::vector<ExcelArray::col_t> uniq;
                    for (ExcelArray::col_t j = 0; j < C; ++j)
                    {
                        long count = 0; for (ExcelArray::col_t k = 0; k < C; ++k) if (colsEqual(j, k)) ++count;
                        bool firstSeen = true; for (ExcelArray::col_t k = 0; k < j; ++k) if (colsEqual(j, k)) { firstSeen = false; break; }
                        if (firstSeen && (!once || count == 1)) uniq.push_back(j);
                    }
                    if (uniq.empty()) return returnValue(CellError::NA);
                    std::vector<ExcelObj> vals;
                    for (ExcelArray::row_t i = 0; i < R; ++i) for (auto j : uniq) vals.emplace_back(a.at(i, j));
                    return egtools::core::output(egtools::core::makeArray(R, (ExcelArrayBuilder::col_t)uniq.size(), vals));
                }
            }
            catch (...) { return returnValue(CellError::Value); }
        }
    }

    void registerFilterSort()
    {
        egtools::core::registerFn(L"XMATCH",
            [](const ExcelObj& l, const ExcelObj& a, const ExcelObj& mm, const ExcelObj& sm) -> ExcelObj*
            { return xmatch(l, a, mm, sm); });
        egtools::core::registerFn(L"FILTER",
            [](const ExcelObj& a, const ExcelObj& inc, const ExcelObj& ie) -> ExcelObj*
            { return filterFn(a, inc, ie); });
        egtools::core::registerFn(L"SORT",
            [](const ExcelObj& a, const ExcelObj& idx, const ExcelObj& ord, const ExcelObj& bc) -> ExcelObj*
            { return sortFn(a, idx, ord, bc); });
        egtools::core::registerRawFn(L"SORTBY",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            { return sortByFn(args, info.numArgs()); });
        egtools::core::registerFn(L"UNIQUE",
            [](const ExcelObj& a, const ExcelObj& bc, const ExcelObj& once) -> ExcelObj*
            { return uniqueFn(a, bc, once); });
    }
}
