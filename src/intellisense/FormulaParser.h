// FormulaParser.h — parse an in-edit formula prefix into (functionName, argIndex).
//
// C++ port of Excel-DNA IntelliSense FormulaParser.TryGetFormulaInfo, INCLUDING
// the user's fix for structured table references (Table1[[#All],[Field]]), whose
// inner separators must not be counted as argument separators.
// Reference: plan/ExcelDna.IntelliSense.UIMonitor.FormulaParser_수정사항.txt

#pragma once
#include <string>

namespace egtools::intellisense
{
    // Given the formula text up to the caret (e.g. "=EG.HELLO(1,"), determine the
    // innermost function name being edited and the current 0-based argument index.
    // Returns false if the caret is not inside a function argument list.
    bool tryGetFormulaInfo(const std::wstring& formulaPrefix,
                           std::wstring& functionName,
                           int& currentArgIndex,
                           wchar_t listSeparator = L',');
}
