// AddinMain.cpp — EGTools++ XLL entry / registration orchestration.
//
// Phase 0 / Step B (S1): minimal static-linked xlOil add-in exposing EG.HELLO,
// to prove the self-contained static .xll path (build + load in Excel).
//
// xlOil registers the XLL entry points (xlAutoOpen/Close, xlAddInManagerInfo)
// via XLO_DECLARE_ADDIN. Functions are registered statically via XLO_FUNC_*.
// See plan/04_아키텍처설계.md §3.

#include <xlOil/xlOil.h>
#include <xlOil/XllEntryPoint.h>

#include "intellisense/IntelliSense.h"
#include "core/I18n.h"
#include "core/Registry.h"
#include "ribbon/Ribbon.h"

using namespace xloil;

// The add-in object: constructed on xlAutoOpen, destroyed on xlAutoClose.
struct EGToolsAddin
{
    EGToolsAddin()
    {
        // Load embedded i18n catalogs and pick the UI language.
        egtools::i18n::load();
        egtools::i18n::detectAndSet();
        // Register worksheet functions (version-adaptive names + i18n help/args).
        egtools::core::registerFunctions();
        // Start in-cell IntelliSense. Its WinEvent hook runs on a background thread
        // but the Excel C API read (LPenHelper) is marshalled to Excel's main thread
        // (see IntelliSense.cpp / getFormulaPrefix), which is required on Excel 2016
        // where a cross-thread call crashes. plan/13_x86_2016_런타임수정.md.
        egtools::intellisense::install();
        // Ribbon: "EGTools++" tab with the compatibility-conversion commands.
        // Deferred + retried on the WINDOW queue: at startup auto-load there is
        // no workbook window yet, so the COM connection it needs cannot be made
        // during xlAutoOpen. plan/18 §3.
        egtools::ribbon::install();
    }

    ~EGToolsAddin()
    {
        egtools::ribbon::uninstall();
        egtools::intellisense::uninstall();
        egtools::core::unregisterFunctions();
        // NB: the message-queue/COM teardown a static XLL needs happens at the
        // very end of xlAutoClose (xlOil patch: reload-safety), not here — it
        // must run after xlOil unregisters its own functions, which still need
        // the message queue. plan/18.
    }

    static std::wstring addInManagerInfo()
    {
        // Shown in Excel's Add-in Manager / options window.
        return std::wstring(L"EGTools++");
    }
};
XLO_DECLARE_ADDIN(EGToolsAddin);

// EG.HELLO() — connectivity smoke test for Phase 0.
XLO_FUNC_START( EG_HELLO() )
{
    return returnValue(L"EGTools++ OK");
}
XLO_FUNC_END(EG_HELLO)
    .threadsafe()
    .name(L"EG.HELLO")
    .category(L"EGTools")
    .help(L"Returns a greeting — EGTools++ connectivity test.");
