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
#include <string>
#include <algorithm>
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
        HWINEVENTHOOK g_hook = nullptr;
        CRITICAL_SECTION g_logLock;
        bool          g_logInit = false;
        std::wstring  g_lastPrefix;
        DWORD         g_lastEventThread = 0;
        HWND          g_lastEventWnd = nullptr;
        UINT_PTR      g_timer = 0;   // pending "settled re-read" timer

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
            if (tryGetFormulaInfo(prefix, fn, argIndex) && lookupFunction(fn, fi))
            {
                POINT p = caretScreenPos(threadId, wnd);
                toolTipShow(p.x, p.y + 2, fi, argIndex);
                return;
            }
            toolTipHide();
        }

        void doUpdate()
        {
            std::wstring prefix;
            int caret = 0;
            if (getFormulaPrefix(prefix, caret))
            {
                if (prefix != g_lastPrefix)
                {
                    g_lastPrefix = prefix;
                    updateIntelliSense(prefix, g_lastEventThread, g_lastEventWnd);
                }
            }
            else if (!g_lastPrefix.empty())
            {
                g_lastPrefix.clear();
                toolTipHide();
            }
        }

        void CALLBACK winEventProc(
            HWINEVENTHOOK, DWORD event, HWND hwnd,
            LONG, LONG, DWORD idEventThread, DWORD)
        {
            if (!hwnd) return;
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
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
            if (g_hook) { UnhookWinEvent(g_hook); g_hook = nullptr; }
            toolTipDestroy();   // window/class cleanup on the owning thread
            return 0;
        }
    }

    void install()
    {
        if (g_thread) return;
        if (!g_logInit) { InitializeCriticalSection(&g_logLock); g_logInit = true; }
        g_thread = CreateThread(nullptr, 0, threadProc, nullptr, 0, &g_threadId);
    }

    void uninstall()
    {
        if (!g_thread) return;
        PostThreadMessageW(g_threadId, WM_QUIT, 0, 0);
        WaitForSingleObject(g_thread, 2000);
        CloseHandle(g_thread);
        g_thread = nullptr;
        g_threadId = 0;
    }
}
