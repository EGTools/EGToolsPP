#include "ToolTipWindow.h"

#include <windows.h>
#include <string>
#include <vector>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

namespace egtools::intellisense
{
    namespace
    {
        const wchar_t* kClass = L"EGToolsPP_IntelliSenseTip";
        HWND  g_wnd = nullptr;
        HFONT g_font = nullptr;
        HFONT g_fontBold = nullptr;

        // Current content (owned by the IntelliSense thread).
        FuncInfo g_func;
        int      g_argIndex = -1;

        // User-chosen position: once the user drags the box, remember where they
        // put it and keep showing it there (until they drag again). Before any
        // drag, the box anchors near the formula caret.
        POINT g_userPos{};
        bool  g_hasUserPos = false;

        const int kPadX = 8, kPadY = 6, kGap = 4;

        HFONT makeFont(bool bold)
        {
            return CreateFontW(-15, 0, 0, 0, bold ? FW_SEMIBOLD : FW_NORMAL,
                FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH,
                L"Segoe UI");
        }

        // Build the signature pieces: name, "(", then arg names separated by ", ",
        // then ")". Each piece carries whether it is the active argument.
        struct Piece { std::wstring text; bool active; };

        std::wstring argLabel(const std::wstring& name, bool optional)
        {
            // Optional arguments (the user may omit them) show as "[name]".
            return optional ? (L"[" + name + L"]") : name;
        }

        // Flat (non-repeating) signature: list every argument 1:1.
        void buildFlat(std::vector<Piece>& tokens, const FuncInfo& f, int argIndex)
        {
            for (size_t i = 0; i < f.args.size(); ++i)
                tokens.push_back({ argLabel(f.args[i].name, f.args[i].optional),
                                   (int)i == argIndex });
        }

        // R2: variadic signature with a repeating block. The catalog holds ONE
        // representative repetition (head + period + tail). We render the head,
        // repetition #1, then — once the caret moves past it — the CURRENT
        // repetition with its real iteration number, with ellipses marking the
        // repetitions we fold away. The active slot maps onto the current rep by
        // (i-head) mod period. The tail (e.g. LET's `calculation`) is always shown
        // but never highlighted: mid-edit we cannot tell a new repetition from the
        // final tail argument (an inherent ambiguity Excel shares).
        void buildRepeat(std::vector<Piece>& tokens, const FuncInfo& f, int argIndex)
        {
            const auto& sp = f.repeat;
            const int H = sp.head, P = sp.period, T = sp.tail;
            const int N = (int)f.args.size();
            if (P <= 0 || N != H + P + T) { buildFlat(tokens, f, argIndex); return; }

            const int maxIter = (sp.max - H - T) / P;   // total repetitions allowed
            const int i = argIndex;

            // Head args.
            for (int h = 0; h < H; ++h)
                tokens.push_back({ argLabel(f.args[h].name, f.args[h].optional), i == h });

            // Which repetition does the caret sit in, and which slot within it?
            int currentRep = 1, activeSlot = -1;
            if (i >= H)
            {
                currentRep = (i - H) / P + 1;
                if (currentRep > maxIter) currentRep = maxIter;
                activeSlot = (i - H) % P;
            }

            auto pushRep = [&](int k, int slot)
            {
                for (int p = 0; p < P; ++p)
                {
                    const ArgInfo& u = f.args[H + p];
                    const bool optional = (k == 1) ? u.optional : true;
                    tokens.push_back({
                        argLabel(egtools::i18n::repeatArgName(u.name, k), optional),
                        p == slot });
                }
            };

            // Repetition #1 (active only if the caret is actually in it).
            pushRep(1, (i >= H && currentRep == 1) ? activeSlot : -1);

            int lastShown = 1;
            if (currentRep >= 2)
            {
                if (currentRep > 2) tokens.push_back({ L"\x2026", false }); // folded reps
                pushRep(currentRep, activeSlot);
                lastShown = currentRep;
            }
            if (lastShown < maxIter) tokens.push_back({ L"\x2026", false }); // more allowed

            // Tail args — shown, never highlighted.
            for (int t = 0; t < T; ++t)
                tokens.push_back({ argLabel(f.args[H + P + t].name,
                                            f.args[H + P + t].optional), false });
        }

        std::vector<Piece> buildSignature(const FuncInfo& f, int argIndex)
        {
            std::vector<Piece> tokens;
            if (f.repeat.has) buildRepeat(tokens, f, argIndex);
            else              buildFlat(tokens, f, argIndex);

            std::vector<Piece> v;
            v.push_back({ f.name + L"(", false });
            for (size_t k = 0; k < tokens.size(); ++k)
            {
                if (k > 0) v.push_back({ L", ", false });
                v.push_back(std::move(tokens[k]));
            }
            v.push_back({ L")", false });
            return v;
        }

        void paint(HDC dc, RECT rc)
        {
            HBRUSH bg = CreateSolidBrush(GetSysColor(COLOR_INFOBK));
            FillRect(dc, &rc, bg);
            DeleteObject(bg);
            FrameRect(dc, &rc, (HBRUSH)GetStockObject(GRAY_BRUSH));

            SetBkMode(dc, TRANSPARENT);
            const COLORREF normal = GetSysColor(COLOR_INFOTEXT);
            const COLORREF active = RGB(0x10, 0x50, 0xC0);

            // Line 1: signature, piece by piece.
            int x = kPadX, y = kPadY;
            auto pieces = buildSignature(g_func, g_argIndex);
            int lineH = 0;
            for (auto& p : pieces)
            {
                HFONT use = p.active ? g_fontBold : g_font;
                HGDIOBJ old = SelectObject(dc, use);
                SetTextColor(dc, p.active ? active : normal);
                SIZE sz{};
                GetTextExtentPoint32W(dc, p.text.c_str(), (int)p.text.size(), &sz);
                TextOutW(dc, x, y, p.text.c_str(), (int)p.text.size());
                x += sz.cx;
                lineH = (sz.cy > lineH) ? sz.cy : lineH;
                SelectObject(dc, old);
            }

            // Line 2: function help.
            y += lineH + kGap;
            SelectObject(dc, g_font);
            SetTextColor(dc, normal);
            if (!g_func.help.empty())
            {
                RECT t{ kPadX, y, rc.right - kPadX, rc.bottom - kPadY };
                DrawTextW(dc, g_func.help.c_str(), -1, &t, DT_LEFT | DT_TOP | DT_NOPREFIX);
            }
        }

        SIZE measure(const FuncInfo& f, int argIndex)
        {
            HDC dc = GetDC(nullptr);
            HGDIOBJ old = SelectObject(dc, g_font);

            int sigW = 0, lineH = 0;
            auto pieces = buildSignature(f, argIndex);
            for (auto& p : pieces)
            {
                SelectObject(dc, p.active ? g_fontBold : g_font);
                SIZE sz{};
                GetTextExtentPoint32W(dc, p.text.c_str(), (int)p.text.size(), &sz);
                sigW += sz.cx;
                lineH = (sz.cy > lineH) ? sz.cy : lineH;
            }
            SelectObject(dc, g_font);
            SIZE hs{};
            if (!f.help.empty())
                GetTextExtentPoint32W(dc, f.help.c_str(), (int)f.help.size(), &hs);

            SelectObject(dc, old);
            ReleaseDC(nullptr, dc);

            int w = (sigW > hs.cx ? sigW : hs.cx) + kPadX * 2;
            int h = lineH + (f.help.empty() ? 0 : (kGap + hs.cy)) + kPadY * 2;
            return SIZE{ w, h };
        }

        LRESULT CALLBACK wndProc(HWND h, UINT m, WPARAM w, LPARAM l)
        {
            // Make the whole window a drag handle so the user can reposition it.
            // (WS_EX_NOACTIVATE keeps Excel's cell edit active while dragging.)
            if (m == WM_NCHITTEST)
                return HTCAPTION;

            // After a user drag, remember the position.
            if (m == WM_EXITSIZEMOVE)
            {
                RECT r{}; GetWindowRect(h, &r);
                g_userPos = POINT{ r.left, r.top };
                g_hasUserPos = true;
                return 0;
            }

            if (m == WM_PAINT)
            {
                PAINTSTRUCT ps;
                HDC dc = BeginPaint(h, &ps);
                RECT rc; GetClientRect(h, &rc);
                paint(dc, rc);
                EndPaint(h, &ps);
                return 0;
            }
            return DefWindowProcW(h, m, w, l);
        }

        void ensureWindow()
        {
            if (g_wnd) return;
            g_font = makeFont(false);
            g_fontBold = makeFont(true);

            WNDCLASSEXW wc{ sizeof(wc) };
            wc.lpfnWndProc = wndProc;
            wc.hInstance = GetModuleHandleW(nullptr);
            wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
            wc.lpszClassName = kClass;
            // A previous load of the add-in in this Excel session may have left
            // the class registered with a wndproc pointing into the old module
            // image — clear it so registration below is always fresh.
            UnregisterClassW(kClass, wc.hInstance);
            RegisterClassExW(&wc);

            g_wnd = CreateWindowExW(
                WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
                kClass, L"", WS_POPUP,
                0, 0, 10, 10, nullptr, nullptr, wc.hInstance, nullptr);
        }
    }

    void toolTipShow(int x, int y, const FuncInfo& f, int argIndex)
    {
        ensureWindow();
        if (!g_wnd) return;
        g_func = f;
        g_argIndex = argIndex;
        SIZE sz = measure(f, argIndex);

        // Position priority: user-chosen position (if ever dragged) > current
        // position while visible (don't jump per keystroke) > caret anchor.
        int px = x, py = y;
        if (g_hasUserPos)
        {
            px = g_userPos.x; py = g_userPos.y;
        }
        else if (IsWindowVisible(g_wnd))
        {
            RECT r{}; GetWindowRect(g_wnd, &r);
            px = r.left; py = r.top;
        }
        SetWindowPos(g_wnd, HWND_TOPMOST, px, py, sz.cx, sz.cy,
                     SWP_NOACTIVATE | SWP_SHOWWINDOW);
        InvalidateRect(g_wnd, nullptr, TRUE);
        UpdateWindow(g_wnd);
    }

    void toolTipHide()
    {
        if (g_wnd) ShowWindow(g_wnd, SW_HIDE);
    }

    // Tear down the tooltip window + class so an add-in unload→reload cycle in
    // the same Excel session starts clean. Must run on the thread that created
    // the window (the IntelliSense background thread, after its message loop).
    void toolTipDestroy()
    {
        if (g_wnd) { DestroyWindow(g_wnd); g_wnd = nullptr; }
        if (g_font) { DeleteObject(g_font); g_font = nullptr; }
        if (g_fontBold) { DeleteObject(g_fontBold); g_fontBold = nullptr; }
        UnregisterClassW(kClass, GetModuleHandleW(nullptr));
    }
}
