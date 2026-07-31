// Ribbon.cpp — EGTools++ Ribbon (COM add-in) via xlOil.
//
// xlOil's makeComAddin()/connect() registers a COM add-in and serves the given
// customUI XML through IRibbonExtensibility — this works from a static XLL and
// runs its callbacks on Excel's main thread. We expose one tab ("EGTools++")
// with one group ("EGTools") and two buttons wired to the conversion commands
// in Convert.cpp. Button captions/tips are localised from the current i18n
// language (self-contained table — no catalog edits).

#include "Ribbon.h"
#include "Convert.h"
#include "../core/I18n.h"

#include <xlOil/ExcelUI.h>
#include <xlOil/ExcelThread.h>

#include <atomic>
#include <memory>
#include <string>

using namespace xloil;

// Static XLL does not auto-connect COM (see FxLet.cpp).
namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::ribbon
{
    namespace
    {
        std::shared_ptr<IComAddin> g_addin;
        std::atomic<bool> g_shutdown{ false };
        int g_attempts = 0;                     // main-thread only
        constexpr int kMaxAttempts = 600;       // ~10 min at 1s between retries

        // ── localisation ─────────────────────────────────────────────────────
        struct Strings
        {
            const wchar_t* compatLabel;
            const wchar_t* compatTip;
            const wchar_t* nativeLabel;
            const wchar_t* nativeTip;
        };

        const Strings& strings()
        {
            // Keyed by egtools::i18n::current() (ko / en / ja / zh-CN / zh-TW / es).
            static const Strings ko = {
                L"구버전 호환 적용",
                L"모던/네이티브(_xlfn.*) 함수를 EGTools 호환 함수명으로 일괄 변환합니다. 변환 전 백업 파일이 저장됩니다.",
                L"내장 함수로 복원",
                L"EGTools 호환 함수를 네이티브 내장 함수로 일괄 복원합니다(IMAGE는 EG.IMAGE 유지). 변환 전 백업 파일이 저장됩니다." };
            static const Strings en = {
                L"Apply Legacy Compatibility",
                L"Batch-convert modern/native (_xlfn.*) function tokens to EGTools-compatible names. A backup file is saved first.",
                L"Restore Native Functions",
                L"Batch-restore EGTools functions to native built-ins (IMAGE stays EG.IMAGE). A backup file is saved first." };
            static const Strings ja = {
                L"旧バージョン互換に変換",
                L"モダン/ネイティブ(_xlfn.*)関数を EGTools 互換関数名に一括変換します。変換前にバックアップを保存します。",
                L"組み込み関数に復元",
                L"EGTools 互換関数をネイティブ組み込み関数に一括復元します(IMAGE は EG.IMAGE を維持)。変換前にバックアップを保存します。" };
            static const Strings zhCN = {
                L"应用旧版兼容",
                L"将现代/原生(_xlfn.*)函数批量转换为 EGTools 兼容函数名。转换前会保存备份文件。",
                L"还原为内置函数",
                L"将 EGTools 函数批量还原为原生内置函数(IMAGE 保持 EG.IMAGE)。转换前会保存备份文件。" };
            static const Strings zhTW = {
                L"套用舊版相容",
                L"將現代/原生(_xlfn.*)函數批次轉換為 EGTools 相容函數名稱。轉換前會儲存備份檔案。",
                L"還原為內建函數",
                L"將 EGTools 函數批次還原為原生內建函數(IMAGE 維持 EG.IMAGE)。轉換前會儲存備份檔案。" };
            static const Strings es = {
                L"Aplicar compatibilidad",
                L"Convierte funciones modernas/nativas (_xlfn.*) a nombres compatibles con EGTools. Se guarda una copia de seguridad antes.",
                L"Restaurar funciones nativas",
                L"Restaura las funciones de EGTools a funciones nativas (IMAGE se mantiene como EG.IMAGE). Se guarda una copia de seguridad antes." };

            const std::wstring lang = egtools::i18n::current();
            if (lang == L"ko") return ko;
            if (lang == L"ja") return ja;
            if (lang == L"zh-CN") return zhCN;
            if (lang == L"zh-TW") return zhTW;
            if (lang == L"es") return es;
            return en;
        }

        std::wstring ribbonXml()
        {
            const Strings& s = strings();
            std::wstring xml =
                L"<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">"
                L"<ribbon><tabs>"
                L"<tab id=\"egpp.tab\" label=\"EGTools++\">"
                L"<group id=\"egpp.grp\" label=\"EGTools\">"
                L"<button id=\"egpp.toCompat\" size=\"large\" imageMso=\"FunctionWizard\""
                L" onAction=\"onToCompat\" label=\"";
            xml += s.compatLabel;
            xml += L"\" screentip=\"";
            xml += s.compatLabel;
            xml += L"\" supertip=\"";
            xml += s.compatTip;
            xml += L"\"/>"
                L"<button id=\"egpp.toNative\" size=\"large\" imageMso=\"CalculateNow\""
                L" onAction=\"onToNative\" label=\"";
            xml += s.nativeLabel;
            xml += L"\" screentip=\"";
            xml += s.nativeLabel;
            xml += L"\" supertip=\"";
            xml += s.nativeTip;
            xml += L"\"/>"
                L"</group></tab></tabs></ribbon></customUI>";
            return xml;
        }

        // Maps customUI onAction names to C++ callbacks. Office invokes onAction
        // as onAction(control); we ignore the args and run the command.
        IComAddin::RibbonCallback callbackFor(const wchar_t* name)
        {
            const std::wstring n = name ? name : L"";
            if (n == L"onToCompat")
                return [](const RibbonControl&, VARIANT*, int, VARIANT**) { convertToCompat(); };
            if (n == L"onToNative")
                return [](const RibbonControl&, VARIANT*, int, VARIANT**) { convertToNative(); };
            return IComAddin::RibbonCallback();
        }
    }

    void install()
    {
        // When Excel auto-loads the .xll at startup, xlAutoOpen runs before the
        // COM/ribbon machinery is ready — a direct connect fails silently and
        // the tab never appears (plan/14 §6). Defer to the WINDOW queue (the
        // only queue that runs in this static XLL) and retry until Excel
        // accepts the COM add-in connection: return false → xlOil reschedules
        // after waitBetweenRetries.
        g_shutdown = false;
        g_attempts = 0;
        xloil::detail::runExcelThreadImpl(
            []() -> bool
            {
                if (g_shutdown || g_addin)
                    return true;            // done (or cancelled) — stop retrying
                ++g_attempts;
                try
                {
                    if (COM::connectCom())
                    {
                        auto addin = makeComAddin(L"EGTools++", L"EGTools++ 호환 변환 도구");
                        const std::wstring xml = ribbonXml();
                        addin->connect(xml.c_str(), callbackFor);
                        g_addin = std::move(addin);
                        return true;
                    }
                }
                catch (...) {}
                // connectCom() needs a workbook window (it binds via the EXCEL7
                // child of XLMAIN), which does NOT exist while Excel sits on the
                // Start screen — the whole reason a direct install at xlAutoOpen
                // never worked. Keep retrying so the tab appears as soon as the
                // user opens/creates a workbook; give up after ~10 min.
                if (g_attempts >= kMaxAttempts)
                    return true;            // give up quietly
                return false;               // false → xlOil requeues us
            },
            ExcelRunQueue::WINDOW | ExcelRunQueue::ENQUEUE,
            /*waitBeforeCall*/ 200, /*waitBetweenRetries*/ 1000);
    }

    void uninstall()
    {
        g_shutdown = true;   // cancel any pending deferred install
        try
        {
            if (g_addin)
            {
                g_addin->disconnect();
                g_addin->close();
            }
        }
        catch (...) {}
        g_addin.reset();
    }
}
