// FxLookup.cpp — lookup/reference functions (XLOOKUP, …).
//
// Principle (revised D5): array in → array out. XLOOKUP returns the matched
// row/column as an array, so it spills via core::output().

#include "../core/Registry.h"
#include "../core/Spill.h"

#include <xlOil/xlOil.h>
#include <xlOil/ArrayBuilder.h>
#include <xlOil/ExcelArray.h>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // XLOOKUP(lookup_value, lookup_array, return_array,
        //         [if_not_found], [match_mode], [search_mode])
        // v1: scalar lookup_value; match_mode 0/-1/1; search_mode 1/-1.
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

                auto elem = [&](size_t i) -> const ExcelObj& {
                    return vertical ? la.at((ExcelArray::row_t)i, 0)
                                    : la.at(0, (ExcelArray::col_t)i);
                };

                long long found = -1;
                long long approx = -1;   // best approximate index for mm = -1/1

                auto consider = [&](size_t i) -> bool {
                    const int c = ExcelObj::compare(elem(i), lookup);  // -1 / 0 / 1
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

                if (found < 0)
                {
                    if (!ifNotFound.isMissing())
                        return returnValue(ExcelObj(ifNotFound));
                    return returnValue(CellError::NA);
                }

                // Slice return_array at the matched index → array out (spills).
                if (vertical)
                {
                    if ((ExcelArray::row_t)found >= ra.nRows())
                        return returnValue(CellError::Ref);
                    const auto cols = ra.nCols();
                    size_t strLen = 0;
                    for (ExcelArray::col_t j = 0; j < cols; ++j)
                        strLen += ra.at((ExcelArray::row_t)found, j).stringLength();
                    ExcelArrayBuilder b(1, cols, strLen);
                    for (ExcelArray::col_t j = 0; j < cols; ++j)
                        b(0, j) = ra.at((ExcelArray::row_t)found, j);
                    return egtools::core::output(b.toExcelObj());
                }
                else
                {
                    if ((ExcelArray::col_t)found >= ra.nCols())
                        return returnValue(CellError::Ref);
                    const auto rows = ra.nRows();
                    size_t strLen = 0;
                    for (ExcelArray::row_t i = 0; i < rows; ++i)
                        strLen += ra.at(i, (ExcelArray::col_t)found).stringLength();
                    ExcelArrayBuilder b(rows, 1, strLen);
                    for (ExcelArray::row_t i = 0; i < rows; ++i)
                        b(i, 0) = ra.at(i, (ExcelArray::col_t)found);
                    return egtools::core::output(b.toExcelObj());
                }
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
