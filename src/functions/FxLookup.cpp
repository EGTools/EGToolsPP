// FxLookup.cpp — lookup/reference functions (XLOOKUP, …).
//
// Principle (revised D5): array in → array out. XLOOKUP returns the matched
// row/column as an array, so it spills via core::output(). An array
// lookup_value is looked up element-wise (native behaviour) and the results
// spill as one array.

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "../core/Apply.h"
#include "RegexMatchMode.h"

#include <xlOil/xlOil.h>
#include <xlOil/ArrayBuilder.h>
#include <xlOil/ExcelArray.h>

#include <stdexcept>
#include <vector>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // XLOOKUP(lookup_value, lookup_array, return_array,
        //         [if_not_found], [match_mode], [search_mode])
        // match_mode 0/-1/1/3(정규식); search_mode 1/-1.
        ExcelObj* xlookup(
            const ExcelObj& lookup, const ExcelObj& lookupArr, const ExcelObj& returnArr,
            const ExcelObj& ifNotFound, const ExcelObj& matchMode, const ExcelObj& searchMode)
        {
            try
            {
                if (lookupArr.isMissing() || returnArr.isMissing())
                    return returnValue(CellError::NA);

                ExcelArray la(lookupArr);
                ExcelArray ra(returnArr);

                const bool vertical = (la.nCols() == 1);   // N×1 → match rows
                const size_t N = vertical ? la.nRows() : la.nCols();
                if (N == 0)
                    return returnValue(CellError::NA);

                const int mm = matchMode.isMissing()  ? 0 : matchMode.get<int>(0);
                const int sm = searchMode.isMissing() ? 1 : searchMode.get<int>(1);
                if (mm == 3 && (sm == 2 || sm == -2))   // 정규식과 이진 검색은 배타(네이티브)
                    return returnValue(CellError::Value);

                auto elem = [&](size_t i) -> const ExcelObj& {
                    return vertical ? la.at((ExcelArray::row_t)i, 0)
                                    : la.at(0, (ExcelArray::col_t)i);
                };

                // Index of `key` in the lookup vector, or -1 (honours mm/sm).
                auto find = [&](const ExcelObj& key) -> long long {
                    if (mm == 3)
                    {
                        // match_mode 3: key는 정규식 패턴 — 텍스트가 아니면 #VALUE!.
                        if (key.type() != ExcelType::Str)
                            throw std::invalid_argument("regex pattern must be text");
                        const auto re = regexForLookup(key.toString());
                        if (sm == -1)
                        {
                            for (long long i = (long long)N - 1; i >= 0; --i)
                                if (regexCellMatch(elem((size_t)i), re)) return i;
                        }
                        else
                        {
                            for (size_t i = 0; i < N; ++i)
                                if (regexCellMatch(elem(i), re)) return (long long)i;
                        }
                        return -1;
                    }
                    long long found = -1;
                    long long approx = -1;   // best approximate index for mm = -1/1
                    auto consider = [&](size_t i) -> bool {
                        const int c = ExcelObj::compare(elem(i), key);  // -1 / 0 / 1
                        if (c == 0) { found = (long long)i; return true; }
                        if (mm == -1 && c < 0) {  // largest value <= lookup
                            if (approx < 0 ||
                                ExcelObj::compare(elem((size_t)approx), elem(i)) < 0)
                                approx = (long long)i;
                        }
                        else if (mm == 1 && c > 0) {  // smallest value >= lookup
                            if (approx < 0 ||
                                ExcelObj::compare(elem(i), elem((size_t)approx)) < 0)
                                approx = (long long)i;
                        }
                        return false;
                    };
                    if (sm == -1)
                        for (long long i = (long long)N - 1; i >= 0; --i) { if (consider((size_t)i)) break; }
                    else
                        for (size_t i = 0; i < N; ++i) { if (consider(i)) break; }
                    if (found < 0 && (mm == -1 || mm == 1))
                        found = approx;
                    return found;
                };

                auto notFoundCell = [&]() -> ExcelObj {
                    return ifNotFound.isMissing() ? ExcelObj(CellError::NA)
                                                  : ExcelObj(ifNotFound);
                };

                // Scalar lookup_value: slice return_array at the matched index.
                auto scalarResult = [&](const ExcelObj& key) -> ExcelObj* {
                    // 오류 키는 그 오류를 전파(네이티브: #DIV/0! 키 → #DIV/0!).
                    if (key.type() == ExcelType::Err)
                        return returnValue(ExcelObj(key));
                    const long long found = find(key);
                    if (found < 0)
                    {
                        if (!ifNotFound.isMissing())
                            return returnValue(ExcelObj(ifNotFound));
                        return returnValue(CellError::NA);
                    }
                    if (vertical)
                    {
                        if ((ExcelArray::row_t)found >= ra.nRows())
                            return returnValue(CellError::Ref);
                        const auto cols = ra.nCols();
                        std::vector<ExcelObj> vals;
                        vals.reserve(cols);
                        for (ExcelArray::col_t j = 0; j < cols; ++j)
                            vals.emplace_back(ra.at((ExcelArray::row_t)found, j));
                        return egtools::core::output(egtools::core::makeArray(1, cols, vals));
                    }
                    else
                    {
                        if ((ExcelArray::col_t)found >= ra.nCols())
                            return returnValue(CellError::Ref);
                        const auto rows = ra.nRows();
                        std::vector<ExcelObj> vals;
                        vals.reserve(rows);
                        for (ExcelArray::row_t i = 0; i < rows; ++i)
                            vals.emplace_back(ra.at(i, (ExcelArray::col_t)found));
                        return egtools::core::output(egtools::core::makeArray(rows, 1, vals));
                    }
                };

                if (!lookup.isType(ExcelType::Multi))
                    return scalarResult(lookup);

                ExcelArray lv(lookup);
                if (lv.size() <= 1)
                    return scalarResult(lv.size() == 1 ? lv.at(0, 0) : lookup);

                // Array lookup_value: element-wise lookup, one array out.
                // Native rule (plan/22 실측 T7·U15): the result has the SHAPE of
                // lookup_value and each element DEMOTES to the first value of
                // its matched return slice, regardless of return_array width.
                std::vector<ExcelObj> vals;
                vals.reserve(lv.size());
                for (ExcelArray::row_t i = 0; i < lv.nRows(); ++i)
                    for (ExcelArray::col_t j = 0; j < lv.nCols(); ++j)
                    {
                        if (lv.at(i, j).type() == ExcelType::Err)
                        {
                            vals.emplace_back(lv.at(i, j));   // 오류 키 전파
                            continue;
                        }
                        long long idx;
                        try { idx = find(lv.at(i, j)); }
                        catch (...)   // mm=3: 원소가 텍스트 아님/패턴 오류 → 원소별 #VALUE!
                        {
                            vals.emplace_back(CellError::Value);
                            continue;
                        }
                        if (idx < 0)
                            vals.emplace_back(egtools::core::demote(notFoundCell()));
                        else if (vertical
                                     ? (ExcelArray::row_t)idx >= ra.nRows()
                                     : (ExcelArray::col_t)idx >= ra.nCols())
                            vals.emplace_back(CellError::Ref);
                        else
                            vals.emplace_back(vertical
                                ? ra.at((ExcelArray::row_t)idx, 0)
                                : ra.at(0, (ExcelArray::col_t)idx));
                    }
                return egtools::core::output(
                    egtools::core::makeArray(lv.nRows(), lv.nCols(), vals));
            }
            catch (...)
            {
                return returnValue(CellError::Value);
            }
        }
    }

    void registerLookup()
    {
        egtools::core::registerFn(L"XLOOKUP",
            [](const ExcelObj& lookup, const ExcelObj& lookupArr, const ExcelObj& returnArr,
               const ExcelObj& ifNotFound, const ExcelObj& matchMode,
               const ExcelObj& searchMode) -> ExcelObj*
            {
                return xlookup(lookup, lookupArr, returnArr, ifNotFound, matchMode, searchMode);
            });
    }
}
