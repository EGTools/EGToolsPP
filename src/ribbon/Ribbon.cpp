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
#include <xlOil/State.h>

#include <windows.h>
#include <commctrl.h>
#include <shellapi.h>
#include <winhttp.h>

#include <atomic>
#include <memory>
#include <string>

#pragma comment(lib, "winhttp.lib")

// EG_VERSION_STR("major.minor.build")는 빌드마다 생성되는 BuildNum.h가 제공.
#include <BuildNum.h>

// Widen the narrow EG_VERSION_STR macro to a wide literal.
#define EG_WIDEN2(x) L##x
#define EG_WIDEN(x)  EG_WIDEN2(x)

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
            const wchar_t* aboutGroup;      // ribbon group caption
            const wchar_t* aboutLabel;      // About button caption
            const wchar_t* aboutTip;        // About button supertip
            const wchar_t* aboutSubtitle;   // one-line product description
            const wchar_t* aboutDist;       // "Distribution" field label
            const wchar_t* aboutBase;       // "Based on" field label
            const wchar_t* aboutNewVer;     // "new version %s available" (link text)
        };

        const Strings& strings()
        {
            // Keyed by egtools::i18n::current() (ko / en / ja / zh-CN / zh-TW / es).
            static const Strings ko = {
                L"구버전 호환 적용",
                L"모던/네이티브(_xlfn.*) 함수를 EGTools 호환 함수명으로 일괄 변환합니다. 변환 전 백업 파일이 저장됩니다.",
                L"내장 함수로 복원",
                L"EGTools 호환 함수를 네이티브 내장 함수로 일괄 복원합니다(IMAGE는 EG.IMAGE 유지). 변환 전 백업 파일이 저장됩니다.",
                L"정보",
                L"EGTools++ 정보",
                L"버전·배포처 등 추가기능 기본 정보를 표시합니다.",
                L"구형 Excel용 최신 함수 호환 추가기능",
                L"배포처",
                L"기반",
                L"새 버전 %s 등록됨 — 배포처에서 받기" };
            static const Strings en = {
                L"Apply Legacy Compatibility",
                L"Batch-convert modern/native (_xlfn.*) function tokens to EGTools-compatible names. A backup file is saved first.",
                L"Restore Native Functions",
                L"Batch-restore EGTools functions to native built-ins (IMAGE stays EG.IMAGE). A backup file is saved first.",
                L"Info",
                L"About EGTools++",
                L"Shows add-in information: version, distribution page and more.",
                L"Modern Excel functions for legacy Excel",
                L"Distribution",
                L"Based on",
                L"New version %s available — get it from the distribution page" };
            static const Strings ja = {
                L"旧バージョン互換に変換",
                L"モダン/ネイティブ(_xlfn.*)関数を EGTools 互換関数名に一括変換します。変換前にバックアップを保存します。",
                L"組み込み関数に復元",
                L"EGTools 互換関数をネイティブ組み込み関数に一括復元します(IMAGE は EG.IMAGE を維持)。変換前にバックアップを保存します。",
                L"情報",
                L"EGTools++ について",
                L"バージョン・配布元などアドインの基本情報を表示します。",
                L"旧バージョン Excel 向けモダン関数互換アドイン",
                L"配布元",
                L"ベース",
                L"新バージョン %s 公開 — 配布元から入手" };
            static const Strings zhCN = {
                L"应用旧版兼容",
                L"将现代/原生(_xlfn.*)函数批量转换为 EGTools 兼容函数名。转换前会保存备份文件。",
                L"还原为内置函数",
                L"将 EGTools 函数批量还原为原生内置函数(IMAGE 保持 EG.IMAGE)。转换前会保存备份文件。",
                L"信息",
                L"关于 EGTools++",
                L"显示加载项基本信息:版本、发布页等。",
                L"面向旧版 Excel 的现代函数兼容加载项",
                L"发布页",
                L"基于",
                L"新版本 %s 已发布 — 前往发布页获取" };
            static const Strings zhTW = {
                L"套用舊版相容",
                L"將現代/原生(_xlfn.*)函數批次轉換為 EGTools 相容函數名稱。轉換前會儲存備份檔案。",
                L"還原為內建函數",
                L"將 EGTools 函數批次還原為原生內建函數(IMAGE 維持 EG.IMAGE)。轉換前會儲存備份檔案。",
                L"資訊",
                L"關於 EGTools++",
                L"顯示增益集基本資訊:版本、發佈頁等。",
                L"適用於舊版 Excel 的現代函數相容增益集",
                L"發佈頁",
                L"基於",
                L"新版本 %s 已發佈 — 前往發佈頁取得" };
            static const Strings es = {
                L"Aplicar compatibilidad",
                L"Convierte funciones modernas/nativas (_xlfn.*) a nombres compatibles con EGTools. Se guarda una copia de seguridad antes.",
                L"Restaurar funciones nativas",
                L"Restaura las funciones de EGTools a funciones nativas (IMAGE se mantiene como EG.IMAGE). Se guarda una copia de seguridad antes.",
                L"Información",
                L"Acerca de EGTools++",
                L"Muestra la información del complemento: versión, página de distribución, etc.",
                L"Funciones modernas de Excel para versiones antiguas",
                L"Distribución",
                L"Basado en",
                L"Nueva versión %s disponible — descargar de la página de distribución" };

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
                L"</group>"
                L"<group id=\"egpp.grpInfo\" label=\"";
            xml += s.aboutGroup;
            xml += L"\">"
                L"<button id=\"egpp.about\" size=\"large\" imageMso=\"Info\""
                L" onAction=\"onAbout\" label=\"";
            xml += s.aboutLabel;
            xml += L"\" screentip=\"";
            xml += s.aboutLabel;
            xml += L"\" supertip=\"";
            xml += s.aboutTip;
            xml += L"\"/>"
                L"</group></tab></tabs></ribbon></customUI>";
            return xml;
        }

        constexpr const wchar_t* kRepoUrl = L"https://github.com/EGTools/EGToolsPP";
        constexpr const wchar_t* kReleasesUrl = L"https://github.com/EGTools/EGToolsPP/releases/latest";

        // GitHub 최신 릴리즈 태그("v0.2.40") 조회. 짧은 타임아웃, 실패 시 false.
        bool latestReleaseTag(std::wstring& tag)
        {
            HINTERNET hSession = WinHttpOpen(L"EGTools++",
                WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);
            if (!hSession) return false;
            WinHttpSetTimeouts(hSession, 2000, 2000, 3000, 4000);

            bool ok = false;
            std::string body;
            HINTERNET hConnect = WinHttpConnect(hSession, L"api.github.com",
                                                INTERNET_DEFAULT_HTTPS_PORT, 0);
            if (hConnect)
            {
                HINTERNET hReq = WinHttpOpenRequest(hConnect, L"GET",
                    L"/repos/EGTools/EGToolsPP/releases/latest", nullptr,
                    WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
                if (hReq)
                {
                    if (WinHttpSendRequest(hReq, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
                        WinHttpReceiveResponse(hReq, nullptr))
                    {
                        DWORD avail = 0;
                        do
                        {
                            avail = 0;
                            if (!WinHttpQueryDataAvailable(hReq, &avail) || avail == 0) break;
                            std::string chunk((size_t)avail, '\0');
                            DWORD read = 0;
                            if (!WinHttpReadData(hReq, chunk.data(), avail, &read)) break;
                            body.append(chunk.data(), read);
                        } while (avail > 0);
                    }
                    WinHttpCloseHandle(hReq);
                }
                WinHttpCloseHandle(hConnect);
            }
            WinHttpCloseHandle(hSession);

            const size_t k = body.find("\"tag_name\"");
            if (k == std::string::npos) return false;
            const size_t q1 = body.find('"', body.find(':', k));
            if (q1 == std::string::npos) return false;
            const size_t q2 = body.find('"', q1 + 1);
            if (q2 == std::string::npos) return false;
            const std::string t = body.substr(q1 + 1, q2 - q1 - 1);
            tag.assign(t.begin(), t.end());   // 태그는 ASCII
            return !tag.empty();
        }

        // 태그("v0.2.40" / "0.2.40") → (major, minor, build). 부족한 자리는 0.
        void parseTag(const std::wstring& tag, int& M, int& m, int& b)
        {
            M = m = b = 0;
            const wchar_t* p = tag.c_str();
            if (*p == L'v' || *p == L'V') ++p;
            swscanf_s(p, L"%d.%d.%d", &M, &m, &b);
        }

        // 새 릴리즈가 있으면 태그를 돌려준다.
        bool newVersionAvailable(std::wstring& tag)
        {
            if (!latestReleaseTag(tag)) return false;
            int M, m, b;
            parseTag(tag, M, m, b);
            const int cur[3] = { EG_VER_MAJOR, EG_VER_MINOR, EG_BUILD_NUM };
            const int rel[3] = { M, m, b };
            for (int i = 0; i < 3; ++i)
            {
                if (rel[i] > cur[i]) return true;
                if (rel[i] < cur[i]) return false;
            }
            return false;
        }

        // TaskDialog 하이퍼링크 클릭 → 기본 브라우저로 이동.
        HRESULT CALLBACK aboutDlgProc(HWND, UINT msg, WPARAM, LPARAM lParam, LONG_PTR)
        {
            if (msg == TDN_HYPERLINK_CLICKED && lParam)
                ShellExecuteW(nullptr, L"open", (LPCWSTR)lParam, nullptr, nullptr, SW_SHOWNORMAL);
            return S_OK;
        }

        // About: 모달 TaskDialog — 버전/비트수, 클릭 가능한 배포처 링크, 새 릴리즈 알림.
        // (comctl32 v6의 TaskDialogIndirect를 동적 로드 — 없으면 MessageBox 폴백.)
        void showAbout()
        {
            const Strings& s = strings();
            const wchar_t* bits = (sizeof(void*) == 8) ? L"64" : L"32";
            const HWND owner = (HWND)xloil::Environment::excelProcess().hWnd;

            std::wstring title = L"EGTools++ " EG_WIDEN(EG_VERSION_STR) L" (";
            title += bits;
            title += L"-bit)";

            std::wstring newVerTag;
            const bool hasNew = newVersionAvailable(newVerTag);

            using TDI = HRESULT(WINAPI*)(const TASKDIALOGCONFIG*, int*, int*, BOOL*);
            HMODULE hCom = GetModuleHandleW(L"comctl32.dll");
            if (!hCom) hCom = LoadLibraryW(L"comctl32.dll");
            const TDI taskDialog = hCom ? (TDI)GetProcAddress(hCom, "TaskDialogIndirect")
                                        : nullptr;
            if (taskDialog)
            {
                std::wstring content = s.aboutSubtitle;
                content += L"\n\n";
                content += s.aboutDist;
                content += L": <a href=\"";
                content += kRepoUrl;
                content += L"\">github.com/EGTools/EGToolsPP</a>\n";
                content += s.aboutBase;
                content += L": xlOil 0.22.1 (Apache-2.0)";
                if (hasNew)
                {
                    wchar_t line[256];
                    swprintf_s(line, s.aboutNewVer, newVerTag.c_str());
                    content += L"\n\n<a href=\"";
                    content += kReleasesUrl;
                    content += L"\">";
                    content += line;
                    content += L"</a>";
                }

                TASKDIALOGCONFIG cfg{};
                cfg.cbSize = sizeof(cfg);
                cfg.hwndParent = owner;
                cfg.dwFlags = TDF_ENABLE_HYPERLINKS | TDF_ALLOW_DIALOG_CANCELLATION |
                              TDF_SIZE_TO_CONTENT;
                cfg.dwCommonButtons = TDCBF_OK_BUTTON;
                cfg.pszWindowTitle = s.aboutLabel;
                cfg.pszMainIcon = TD_INFORMATION_ICON;
                cfg.pszMainInstruction = title.c_str();
                cfg.pszContent = content.c_str();
                cfg.pfCallback = aboutDlgProc;
                taskDialog(&cfg, nullptr, nullptr, nullptr);
                return;
            }

            // 폴백: 링크 없는 일반 메시지 상자.
            std::wstring text = title + L"\n";
            text += s.aboutSubtitle;
            text += L"\n\n";
            text += s.aboutDist;
            text += L": ";
            text += kRepoUrl;
            text += L"\n";
            text += s.aboutBase;
            text += L": xlOil 0.22.1 (Apache-2.0)";
            if (hasNew)
            {
                wchar_t line[256];
                swprintf_s(line, s.aboutNewVer, newVerTag.c_str());
                text += L"\n\n";
                text += line;
            }
            MessageBoxW(owner, text.c_str(), s.aboutLabel, MB_OK | MB_ICONINFORMATION);
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
            if (n == L"onAbout")
                return [](const RibbonControl&, VARIANT*, int, VARIANT**) { showAbout(); };
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
