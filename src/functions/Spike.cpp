// Spike.cpp — Phase 0 / S5 proof-of-concept functions.
//
// Validates the building blocks of the dynamic-array (spill) story:
//   * EG.SEQ2D    — return a 2D array, to confirm array marshalling and native
//     spill on Excel 2021/365. On legacy Excel (<=2019) this would only show
//     the top-left cell without the forced-resize engine (see plan/04 §4).
//
// The legacy forced-resize engine itself (isMaster / Caller / runExcelThread /
// xlcFormulaArray) is designed but not built here — it cannot be validated
// without a <=2019 Excel. See plan/08_Phase0_검증보고.md.

#include <xlOil/xlOil.h>
#include <xlOil/ArrayBuilder.h>
#include <xlOil/State.h>
#include <algorithm>

#include "../core/I18n.h"
#include "../core/Version.h"

#include <string>

using namespace xloil;

// S6/Phase1 diagnostic: report detected Excel edition + native-function decisions.
XLO_FUNC_START( EG_VERSIONINFO() )
{
    std::wstring s = L"year=" + std::to_wstring(egtools::core::excelYear())
        + L"; DA=" + (egtools::core::supportsDynamicArrays() ? L"1" : L"0")
        + L"; XLOOKUP_native=" + (egtools::core::hasNativeFunction(L"XLOOKUP") ? L"1" : L"0")
        + L"; IFS_native=" + (egtools::core::hasNativeFunction(L"IFS") ? L"1" : L"0")
        + L"; FOO_native=" + (egtools::core::hasNativeFunction(L"FOOBAR") ? L"1" : L"0");
    return returnValue(s);
}
XLO_FUNC_END(EG_VERSIONINFO)
    .threadsafe()
    .name(L"EG.VERSIONINFO")
    .category(L"EGTools")
    .help(L"Diagnostic: detected Excel edition and native-function decisions.");

// S6 diagnostics: inspect / switch the active UI language at runtime.
XLO_FUNC_START( EG_LANG() )
{
    return returnValue(egtools::i18n::current());
}
XLO_FUNC_END(EG_LANG)
    .threadsafe()
    .name(L"EG.LANG")
    .category(L"EGTools")
    .help(L"Returns the active EGTools++ UI language code.");

XLO_FUNC_START( EG_SETLANG(const ExcelObj* code) )
{
    egtools::i18n::setCurrent(code->toString());
    return returnValue(egtools::i18n::current());
}
XLO_FUNC_END(EG_SETLANG)
    .name(L"EG.SETLANG")
    .category(L"EGTools")
    .help(L"Sets the EGTools++ UI language (ko/en/zh-CN/zh-TW/ja/es).")
    .arg(L"code", L"language code");

XLO_FUNC_START( EG_SEQ2D(const ExcelObj* rows, const ExcelObj* cols) )
{
    const int nR = std::max(1, rows->get<int>(1));
    const int nC = std::max(1, cols->get<int>(1));

    ExcelArrayBuilder builder(nR, nC);
    int v = 1;
    for (int i = 0; i < nR; ++i)
        for (int j = 0; j < nC; ++j)
            builder(i, j) = static_cast<double>(v++);

    return returnValue(builder.toExcelObj());
}
XLO_FUNC_END(EG_SEQ2D)
    .threadsafe()
    .name(L"EG.SEQ2D")
    .category(L"EGTools")
    .help(L"S5 spike: returns a rows×cols array of 1..n (spills on modern Excel).")
    .arg(L"rows", L"number of rows")
    .arg(L"cols", L"number of columns");
