// IntelliSense.cpp — self-implemented in-cell IntelliSense (D2).
//
// Key technique (learned from Excel-DNA IntelliSense): the live formula text is
// NOT read via UI Automation — standard UIA patterns expose nothing for Excel's
// formula bar / in-cell editor. Instead Excel's C API LPenHelper(xlGetFmlaInfo)
// returns the formula being edited + caret offset directly. UI Automation /
// WinEvents are only used to know *when* editing changes (and later, where to
// place the tooltip).
//
// Stage A (this step): capture the live formula prefix via xlGetFmlaInfo and log
// it. Stages B+ add the formula parser (Excel-DNA algorithm + structured-ref
// fix), the i18n catalog lookup, and the tooltip overlay window.

#include "IntelliSense.h"
#include "FormulaParser.h"
#include "Catalog.h"
#include "ToolTipWindow.h"

#include <windows.h>
#include <commctrl.h>   // LVM_* — read the autocomplete popup's ListView in-process
#include <string>
#include <algorithm>
#include <atomic>
#include <chrono>

#include <xloil/ExcelThread.h>   // runExcelThread: marshal Excel C API to main thread

// Excel C API (exported from xlcall32.lib, already linked via cmake/xlOil.cmake).
extern "C" long __stdcall LPenHelper(int wCode, void* lpv);

namespace egtools::intellisense
{
    namespace
    {
        // From the Excel SDK (xlcall.h).
        constexpr int kXlSpecial     = 0x4000;
        constexpr int kXlGetFmlaInfo = 14 | kXlSpecial;  // 0x400E
        constexpr int kXlModeReady   = 0;                // not editing

        struct FMLAINFO
        {
            int      wPointMode;  // edit mode; 0 (Ready) => rest undefined
            int      cch;         // count of characters in formula
            wchar_t* lpch;        // formula characters (Unicode). READ ONLY.
            int      ichFirst;    // start of selection
            int      ichLast;     // end of selection (may be > cch)
            int      ichCaret;    // caret offset
        };

        HANDLE        g_thread = nullptr;
        DWORD         g_threadId = 0;
        std::atomic<bool> g_shutdown{ false };  // cancels a still-pending deferred install
        HWINEVENTHOOK g_hook = nullptr;
        HWINEVENTHOOK g_hookPopup = nullptr;   // CREATE..SELECTION (popup list)
        bool          g_excelFg = true;        // Excel(우리 프로세스)이 포그라운드인가
        UINT_PTR      g_fgPoll = 0;            // 편집 세션 동안 포그라운드 폴링 타이머
        CRITICAL_SECTION g_logLock;
        bool          g_logInit = false;
        std::wstring  g_lastPrefix;
        DWORD         g_lastEventThread = 0;
        HWND          g_lastEventWnd = nullptr;
        UINT_PTR      g_timer = 0;   // pending "settled re-read" timer

        // Autocomplete popup (Excel-DNA technique): Excel's formula-autocomplete
        // dropdown is a real window (class __XLACOOUTER) whose first child is a
        // standard ListView. Being in-process, the selected item's text/rect can
        // be read directly with LVM_* messages — no UIA needed. While an EGTools
        // entry is selected, its description is shown in an anchored tooltip
        // beside the popup (Excel only renders that yellow box for built-ins).
        HWND g_popupOuter = nullptr;   // __XLACOOUTER
        HWND g_popupList  = nullptr;   // first child (the ListView)
        bool g_descActive = false;     // description overlay currently shown

        void logLine(const std::wstring& s)
        {
            std::wstring line = s + L"\r\n";
            int n = WideCharToMultiByte(CP_UTF8, 0, line.c_str(), (int)line.size(),
                                        nullptr, 0, nullptr, nullptr);
            std::string utf8(n, '\0');
            WideCharToMultiByte(CP_UTF8, 0, line.c_str(), (int)line.size(),
                                utf8.data(), n, nullptr, nullptr);
            if (g_logInit) EnterCriticalSection(&g_logLock);
            {
                wchar_t tmp[MAX_PATH]{}; GetTempPathW(MAX_PATH, tmp);
                std::wstring path = std::wstring(tmp) + L"egtools_intellisense.log";
                HANDLE h = CreateFileW(path.c_str(), FILE_APPEND_DATA,
                    FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr,
                    OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
                if (h != INVALID_HANDLE_VALUE)
                {
                    DWORD w = 0; WriteFile(h, utf8.data(), (DWORD)utf8.size(), &w, nullptr);
                    CloseHandle(h);
                }
            }
            if (g_logInit) LeaveCriticalSection(&g_logLock);
        }

        // Returns the formula prefix (text up to the caret) currently being
        // edited, or empty if not editing.
        //
        // The Excel C API (LPenHelper) is MAIN-THREAD ONLY. This function runs on
        // the background WinEvent thread, and calling LPenHelper here crashes Excel
        // 2016 with an access violation (it tolerated it on 365/x64 but that is not
        // supported). So the actual call is marshalled to Excel's main thread via
        // xloil::runExcelThread; we copy the prefix out there (info.lpch is only
        // valid transiently / in-context) and hand back a plain std::wstring.
        // WINDOW|ENQUEUE = dispatch via the hidden main-thread window without COM
        // context switching (xlGetFmlaInfo is an xlSpecial call, not Excel12).
        bool getFormulaPrefix(std::wstring& out, int& caret)
        {
            struct Fml { bool ok; std::wstring prefix; int caret; };

            try
            {
                auto future = xloil::runExcelThread([]() -> Fml
                {
                    FMLAINFO info{};
                    if (LPenHelper(kXlGetFmlaInfo, &info) != 0)
                        return { false, {}, 0 };
                    if (info.wPointMode == kXlModeReady || info.cch <= 0 || !info.lpch)
                        return { false, {}, 0 };
                    // Excel-DNA: prefixLen = min(max(ichCaret, ichLast), cch)
                    int prefixLen = std::min(std::max(info.ichCaret, info.ichLast), info.cch);
                    if (prefixLen < 0) prefixLen = 0;
                    return { true, std::wstring(info.lpch, info.lpch + prefixLen), info.ichCaret };
                }, xloil::ExcelRunQueue::WINDOW | xloil::ExcelRunQueue::ENQUEUE);

                // Don't block this bg thread indefinitely if the main thread isn't
                // servicing the queue (e.g. busy in a modal sub-loop).
                if (future.wait_for(std::chrono::milliseconds(250)) != std::future_status::ready)
                    return false;
                Fml r = future.get();
                if (!r.ok)
                    return false;
                out = std::move(r.prefix);
                caret = r.caret;
                return true;
            }
            catch (...)
            {
                return false;   // marshalling failed — treat as "not editing"
            }
        }

        // Screen position just below the formula caret (fallback: below the
        // editing window).
        POINT caretScreenPos(DWORD threadId, HWND fallback)
        {
            GUITHREADINFO gui{ sizeof(gui) };
            if (GetGUIThreadInfo(threadId, &gui) && gui.hwndCaret)
            {
                POINT p{ gui.rcCaret.left, gui.rcCaret.bottom };
                ClientToScreen(gui.hwndCaret, &p);
                return p;
            }
            RECT r{};
            if (fallback) GetWindowRect(fallback, &r);
            return POINT{ r.left, r.bottom };
        }

        void updateIntelliSense(const std::wstring& prefix, DWORD threadId, HWND wnd)
        {
            std::wstring fn;
            int argIndex = -1;
            FuncInfo fi;
            if (tryGetFormulaInfo(prefix, fn, argIndex) && lookupRegistered(fn, fi))
            {
                POINT p = caretScreenPos(threadId, wnd);
                toolTipShow(p.x, p.y + 2, fi, argIndex);
                return;
            }
            toolTipHide();
        }

        // ── autocomplete-popup description overlay ───────────────────────────
        bool hasClass(HWND h, const wchar_t* cls)
        {
            wchar_t buf[64]{};
            return h && GetClassNameW(h, buf, 64) && wcscmp(buf, cls) == 0;
        }

        void descHide()
        {
            if (!g_descActive) return;
            g_descActive = false;
            toolTipHide();
            // The popup typically closes at the exact moment the user commits
            // the name (types '(' or Tab-completes, which inserts the paren):
            // the VALUECHANGE for that keystroke often arrives BEFORE this HIDE
            // and was deliberately not consumed (doUpdate skips while the desc
            // overlay is active). Clear the last-prefix memory and schedule the
            // settled re-read so the ARGUMENT tooltip takes over immediately.
            g_lastPrefix.clear();
            if (g_timer) KillTimer(nullptr, g_timer);
            g_timer = SetTimer(nullptr, 0, 25, nullptr);
        }

        // In-process read of the popup ListView's selected item; shows/hides the
        // description tooltip accordingly. SendMessageTimeout guards against a
        // wedged main thread (the list lives on Excel's UI thread).
        void updateListSelection()
        {
            if (!g_popupList) { descHide(); return; }

            auto send = [&](UINT msg, WPARAM w, LPARAM l, LRESULT& out) -> bool
            {
                DWORD_PTR r = 0;
                if (!SendMessageTimeoutW(g_popupList, msg, w, l,
                        SMTO_ABORTIFHUNG | SMTO_BLOCK, 200, &r))
                    return false;
                out = (LRESULT)r;
                return true;
            };

            LRESULT idx = -1;
            if (!send(LVM_GETNEXTITEM, (WPARAM)-1, LVNI_SELECTED, idx) || idx < 0)
            { descHide(); return; }

            wchar_t buf[256]{};
            LVITEMW it{};
            it.iSubItem = 0;
            it.pszText = buf;
            it.cchTextMax = 256;
            LRESULT len = 0;
            if (!send(LVM_GETITEMTEXTW, (WPARAM)idx, (LPARAM)&it, len) || len <= 0)
            { descHide(); return; }

            FuncInfo fi;
            if (!lookupRegistered(std::wstring(buf, (size_t)len), fi))
            { descHide(); return; }   // built-in / not ours — Excel's own tooltip

            // Anchor: right edge of the popup, at the selected item's height.
            RECT item{ LVIR_BOUNDS, 0, 0, 0 };
            LRESULT ok = 0;
            send(LVM_GETITEMRECT, (WPARAM)idx, (LPARAM)&item, ok);
            POINT tl{ item.left, item.top };
            ClientToScreen(g_popupList, &tl);
            RECT outer{};
            GetWindowRect(g_popupOuter ? g_popupOuter : g_popupList, &outer);

            g_descActive = true;
            toolTipShowAnchored(outer.right + 4, ok ? tl.y : outer.top, fi);
        }

        // Routes popup-related WinEvents; returns true when consumed.
        bool handlePopupEvent(DWORD event, HWND hwnd)
        {
            switch (event)
            {
            case EVENT_OBJECT_CREATE:
            case EVENT_OBJECT_SHOW:
                if (hasClass(hwnd, L"__XLACOOUTER"))
                {
                    g_popupOuter = hwnd;
                    g_popupList = GetWindow(hwnd, GW_CHILD);
                    updateListSelection();   // first match is pre-selected
                    return true;
                }
                return false;

            case EVENT_OBJECT_HIDE:
            case EVENT_OBJECT_DESTROY:
                if (hwnd && hwnd == g_popupOuter)   // class is gone on DESTROY
                {
                    if (event == EVENT_OBJECT_DESTROY)
                    { g_popupOuter = nullptr; g_popupList = nullptr; }
                    descHide();
                    return true;
                }
                return false;

            case EVENT_OBJECT_SELECTION:
                // Keyboard/mouse moves through the list. Adopt the list window
                // lazily too (popup may predate our hook and get reused).
                if ((g_popupList && hwnd == g_popupList) ||
                    (hwnd && hasClass(GetAncestor(hwnd, GA_PARENT), L"__XLACOOUTER")))
                {
                    g_popupOuter = GetAncestor(hwnd, GA_PARENT);
                    g_popupList = hwnd;
                    updateListSelection();
                    return true;
                }
                return false;

            case EVENT_OBJECT_LOCATIONCHANGE:
                // Popup moved/resized (list filtering) — re-anchor if showing.
                if (g_descActive && hwnd &&
                    (hwnd == g_popupOuter || hwnd == g_popupList))
                {
                    updateListSelection();
                    return true;
                }
                return false;
            }
            return false;
        }

        // 수식 편집 상태는 앱 전환 후에도 유지되므로(셀 편집 모드) 포그라운드가
        // 다른 프로세스면 툴팁을 숨겨야 한다. EVENT_SYSTEM_FOREGROUND 전역 훅은
        // 이 환경에서 발화하지 않는 것으로 실측돼(2026-08-06), 편집 세션 동안만
        // 250ms 폴링으로 전환을 감지한다(비편집 시 타이머 없음 = 상시 비용 0).
        bool excelIsForeground()
        {
            DWORD pid = 0;
            HWND fg = GetForegroundWindow();
            if (fg) GetWindowThreadProcessId(fg, &pid);
            return pid == GetCurrentProcessId();
        }

        void ensureFgPoll()
        {
            if (!g_fgPoll) g_fgPoll = SetTimer(nullptr, 0, 250, nullptr);
        }

        void stopFgPoll()
        {
            if (g_fgPoll) { KillTimer(nullptr, g_fgPoll); g_fgPoll = 0; }
        }

        void doUpdate()
        {
            const bool fg = excelIsForeground();
            if (!fg)
            {
                if (g_excelFg)
                {
                    g_excelFg = false;
                    toolTipHide();      // 다른 앱이 앞 — 즉시 숨김
                }
                ensureFgPoll();         // 복귀 감지를 위해 폴링 유지
                return;
            }
            if (!g_excelFg)
            {
                g_excelFg = true;
                g_lastPrefix.clear();   // 복귀 — 강제 재평가로 툴팁 재표시
            }

            std::wstring prefix;
            int caret = 0;
            if (getFormulaPrefix(prefix, caret))
            {
                ensureFgPoll();         // 편집 중 — 전환 감시 계속
                // While the desc overlay owns the tooltip window, do NOT
                // consume prefix changes — g_lastPrefix must stay stale so the
                // re-read scheduled by descHide() still processes the keystroke
                // that closed the popup (the '(' that starts the arguments).
                if (g_descActive) return;
                if (prefix != g_lastPrefix)
                {
                    g_lastPrefix = prefix;
                    updateIntelliSense(prefix, g_lastEventThread, g_lastEventWnd);
                }
            }
            else
            {
                stopFgPoll();           // 편집 종료 — 폴링 중단
                if (!g_lastPrefix.empty())
                {
                    g_lastPrefix.clear();
                    if (!g_descActive) toolTipHide();   // desc overlay owns the window
                }
            }
        }

        void CALLBACK winEventProc(
            HWINEVENTHOOK, DWORD event, HWND hwnd,
            LONG idObject, LONG, DWORD idEventThread, DWORD)
        {
            if (!hwnd) return;
            // Popup lifecycle/selection first (window-level events only —
            // idObject OBJID_WINDOW=0 / OBJID_CLIENT=-4 both appear here).
            if (idObject == OBJID_WINDOW || idObject == OBJID_CLIENT)
                if (handlePopupEvent(event, hwnd))
                    return;

            if (event != EVENT_OBJECT_VALUECHANGE &&
                event != EVENT_OBJECT_FOCUS &&
                event != EVENT_OBJECT_LOCATIONCHANGE) return;

            g_lastEventThread = idEventThread;
            g_lastEventWnd = hwnd;

            // Read immediately, AND schedule a short settled re-read: Excel updates
            // the formula-edit state (xlGetFmlaInfo) a beat after the key event, so
            // a just-typed separator isn't reflected yet on the immediate read.
            doUpdate();
            if (g_timer) KillTimer(nullptr, g_timer);
            g_timer = SetTimer(nullptr, 0, 25, nullptr);
        }

        DWORD WINAPI threadProc(LPVOID)
        {
            g_hook = SetWinEventHook(
                EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_VALUECHANGE,
                nullptr, winEventProc,
                GetCurrentProcessId(), 0,
                WINEVENT_OUTOFCONTEXT);
            // Popup dropdown lifecycle + selection (CREATE 0x8000 .. SELECTION
            // 0x8006) — separate hook so the main range stays narrow.
            g_hookPopup = SetWinEventHook(
                EVENT_OBJECT_CREATE, EVENT_OBJECT_SELECTION,
                nullptr, winEventProc,
                GetCurrentProcessId(), 0,
                WINEVENT_OUTOFCONTEXT);
            g_excelFg = excelIsForeground();
            logLine(L"--- IntelliSense thread ready (LPenHelper) ---");

            MSG msg;
            while (GetMessageW(&msg, nullptr, 0, 0) > 0)
            {
                if (msg.message == WM_TIMER && g_timer && msg.wParam == g_timer)
                {
                    KillTimer(nullptr, g_timer);
                    g_timer = 0;
                    doUpdate();   // settled re-read
                    continue;
                }
                if (msg.message == WM_TIMER && g_fgPoll && msg.wParam == g_fgPoll)
                {
                    doUpdate();   // 주기 폴링(편집 세션 동안): 포그라운드 전환 감지
                    continue;
                }
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
            stopFgPoll();
            if (g_hookPopup) { UnhookWinEvent(g_hookPopup); g_hookPopup = nullptr; }
            if (g_hook) { UnhookWinEvent(g_hook); g_hook = nullptr; }
            toolTipDestroy();   // window/class cleanup on the owning thread
            return 0;
        }
    }

    void install()
    {
        if (g_thread) return;
        if (!g_logInit) { InitializeCriticalSection(&g_logLock); g_logInit = true; }
        g_shutdown = false;
        // [F8] hardening — do NOT set the process-wide WinEvent hooks in xlAutoOpen.
        // install() runs during Excel's boot add-in-loading loop; a hook thread live
        // there fires winEventProc → runExcelThread (main-thread marshal) while a
        // co-loaded ribbon XLL runs its own fragile boot xlfRegister, and THIS add-in
        // is the potential aggressor in that race (plan/06 [F8]). IntelliSense has no
        // job before a workbook window exists, so defer the hook thread to the WINDOW
        // queue: it only runs once a workbook is up, i.e. after the boot loop. The
        // guards make a load→unload before the queue fires a clean no-op.
        xloil::runExcelThread([]()
        {
            if (g_shutdown || g_thread) return;   // uninstalled first, or already up
            g_thread = CreateThread(nullptr, 0, threadProc, nullptr, 0, &g_threadId);
        }, xloil::ExcelRunQueue::WINDOW | xloil::ExcelRunQueue::ENQUEUE);
    }

    void uninstall()
    {
        g_shutdown = true;   // cancel a deferred install that has not fired yet
        if (!g_thread) return;
        PostThreadMessageW(g_threadId, WM_QUIT, 0, 0);
        WaitForSingleObject(g_thread, 2000);
        CloseHandle(g_thread);
        g_thread = nullptr;
        g_threadId = 0;
    }
}
