// FxInfo.cpp — information / reference functions.
//
// All take a cell REFERENCE (xlOil RangeArg) and use XLM information helpers
// (ISFORMULA / GET.CELL / GET.WORKBOOK / xlSheetNm), so they are registered
// macro-type (grants macro-sheet API access — the way Excel-DNA does it).

#include "../core/Registry.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelRef.h>      // RangeArg
#include <xlOil/ExcelCall.h>     // tryCallExcel
#include <xlOil/ExcelArray.h>
#include <string>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // Full sheet name ("[Book]Sheet") of a reference, via xlSheetNm — matches
        // the entries returned by GET.WORKBOOK(1).
        std::wstring sheetNameOf(const ExcelObj& ref)
        {
            auto r = tryCallExcel(msxll::xlSheetNm, ref);
            return r.second == 0 ? r.first.toString() : std::wstring();
        }

        // 1-based position of a sheet within the active workbook.
        ExcelObj sheetPosition(const std::wstring& fullName)
        {
            if (fullName.empty()) return ExcelObj(CellError::NA);
            auto names = tryCallExcel(msxll::xlfGetWorkbook, 1);   // ordered sheet names
            if (names.second != 0) return ExcelObj(CellError::Value);
            ExcelArray a(names.first);
            const size_t total = (size_t)a.nRows() * a.nCols();
            for (size_t k = 0; k < total; ++k)
                if (a.at(k).toString() == fullName)
                    return ExcelObj((double)(k + 1));
            return ExcelObj(CellError::NA);
        }
    }

    void registerInfo()
    {
        // FORMULATEXT(reference): the cell's formula as text, else #N/A.
        egtools::core::registerFn(L"FORMULATEXT",
            [](const RangeArg& ref) -> ExcelObj*
            {
                if (ref.isMissing()) return returnValue(CellError::NA);
                auto isF = tryCallExcel(msxll::xlfIsformula, ref);
                if (isF.second != 0 || !isF.first.get<bool>(false))
                    return returnValue(CellError::NA);
                auto f = tryCallExcel(msxll::xlfGetCell, 6, ref);   // 6 = formula text
                if (f.second != 0) return returnValue(CellError::NA);
                return returnValue(f.first);
            }, /*macro*/ true);

        // ISFORMULA(reference): TRUE if the cell contains a formula.
        egtools::core::registerFn(L"ISFORMULA",
            [](const RangeArg& ref) -> ExcelObj*
            {
                if (ref.isMissing()) return returnValue(CellError::Value);
                auto isF = tryCallExcel(msxll::xlfIsformula, ref);
                if (isF.second != 0) return returnValue(CellError::Value);
                return returnValue(ExcelObj(isF.first.get<bool>(false)));
            }, /*macro*/ true);

        // SHEET([reference]): 1-based sheet position (of the reference, or caller).
        egtools::core::registerFn(L"SHEET",
            [](const RangeArg& ref) -> ExcelObj*
            {
                std::wstring name;
                if (ref.isMissing())
                {
                    auto caller = tryCallExcel(msxll::xlfCaller);
                    if (caller.second == 0) name = sheetNameOf(caller.first);
                }
                else
                    name = sheetNameOf(ref);
                return returnValue(sheetPosition(name));
            }, /*macro*/ true);

        // SHEETS([reference]): worksheet count (or 1 for a single-area reference).
        egtools::core::registerFn(L"SHEETS",
            [](const RangeArg& ref) -> ExcelObj*
            {
                if (!ref.isMissing()) return returnValue(ExcelObj(1));
                auto n = tryCallExcel(msxll::xlfGetWorkbook, 4);   // sheet count
                if (n.second != 0) return returnValue(CellError::Value);
                return returnValue(n.first);
            }, /*macro*/ true);
    }
}
