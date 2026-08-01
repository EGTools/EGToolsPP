#include "ToolTipWindow.h"

#include <windows.h>
#include <windowsx.h>
#include <shellapi.h>
#include <string>
#include <vector>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "shell32.lib")

namespace egtools::intellisense
{
    namespace
    {
        const wchar_t* kClass = L"EGToolsPP_IntelliSenseTip";
        HWND  g_wnd = nullptr;
        HFONT g_font = nullptr;
        HFONT g_fontBold = nullptr;
        HFONT g_fontLink = nullptr;   // 함수명 하이퍼링크(밑줄)

        // Current content (owned by the IntelliSense thread).
        FuncInfo g_func;
        int      g_argIndex = -1;
        RECT     g_linkRect{};        // 함수명 영역(클라이언트 좌표) — 클릭=매뉴얼 열기

        // User-chosen position: once the user drags the box, remember where they
        // put it and keep showing it there (until they drag again). Before any
        // drag, the box anchors near the formula caret.
        POINT g_userPos{};
        bool  g_hasUserPos = false;

        const int kPadX = 8, kPadY = 6, kGap = 4;

        HFONT makeFont(bool bold, bool underline = false)
        {
            return CreateFontW(-15, 0, 0, 0, bold ? FW_SEMIBOLD : FW_NORMAL,
                FALSE, underline ? TRUE : FALSE, FALSE, DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                DEFAULT_PITCH, L"Segoe UI");
        }

        // UI 언어 → 준비된 매뉴얼 언어 폴더. 준비된 번역이 없으면 영어로 폴백.
        // 번역이 추가되면 여기에 매핑 한 줄만 더한다 (예: ja → L"ja").
        std::wstring manualLang()
        {
            const std::wstring ui = egtools::i18n::current();
            if (ui == L"ko") return L"kr";
            if (ui == L"en") return L"en";
            return L"en";   // ja/es/zh-CN/zh-TW 등 미준비 언어 → 영어 폴백
        }

        // 함수명 → GitHub 매뉴얼 URL. EG./x 접두를 벗겨 bare 이름으로 연결한다.
        std::wstring manualUrl(const std::wstring& funcName)
        {
            std::wstring bare = funcName;
            if (bare.rfind(L"EG.", 0) == 0) bare = bare.substr(3);
            else if (bare.size() > 1 && bare[0] == L'x' &&
                     iswupper(bare[1]))          // xSORT/xFILTER/xLET (구버전 충돌 회피명)
                bare = bare.substr(1);
            return L"https://github.com/EGTools/EGToolsPP/blob/main/manual/" +
                   manualLang() + L"/" + bare + L".md";
        }

        // Build the signature pieces: name, "(", then arg names separated by ", ",
        // then ")". Each piece carries whether it is the active argument; `link`
        // marks the clickable function name.
        struct Piece { std::wstring text; bool active; bool link = false; };

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

        // 2행 텍스트: 활성 인수가 있으면 그 인수의 설명("이름: 설명"),
        // 없거나 설명이 비어 있으면 함수 설명. (반복 인수는 강조와 같은 규칙으로
        // 대표 슬롯에 매핑하고 표시 이름에 반복 회차를 반영한다.)
        std::wstring line2Text(const FuncInfo& f, int argIndex)
        {
            const ArgInfo* arg = nullptr;
            std::wstring shownName;

            if (argIndex >= 0)
            {
                if (f.repeat.has)
                {
                    const auto& sp = f.repeat;
                    const int H = sp.head, P = sp.period, T = sp.tail;
                    const int N = (int)f.args.size();
                    if (P > 0 && N == H + P + T)
                    {
                        if (argIndex < H)
                        {
                            arg = &f.args[argIndex];
                            shownName = arg->name;
                        }
                        else
                        {
                            const int maxIter = (sp.max - H - T) / P;
                            int rep = (argIndex - H) / P + 1;
                            if (rep > maxIter) rep = maxIter;
                            const int slot = (argIndex - H) % P;
                            arg = &f.args[H + slot];
                            shownName = egtools::i18n::repeatArgName(arg->name, rep);
                        }
                    }
                }
                else if (argIndex < (int)f.args.size())
                {
                    arg = &f.args[argIndex];
                    shownName = arg->name;
                }
            }

            if (arg && !arg->help.empty())
                return shownName + L": " + arg->help;
            return f.help;
        }

        std::vector<Piece> buildSignature(const FuncInfo& f, int argIndex)
        {
            std::vector<Piece> tokens;
            if (f.repeat.has) buildRepeat(tokens, f, argIndex);
            else              buildFlat(tokens, f, argIndex);

            std::vector<Piece> v;
            v.push_back({ f.name, false, true });   // 함수명 = 매뉴얼 링크
            v.push_back({ L"(", false });
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
            const COLORREF linkCol = RGB(0x00, 0x66, 0xCC);
            int x = kPadX, y = kPadY;
            auto pieces = buildSignature(g_func, g_argIndex);
            int lineH = 0;
            SetRectEmpty(&g_linkRect);
            for (auto& p : pieces)
            {
                HFONT use = p.link ? g_fontLink : (p.active ? g_fontBold : g_font);
                HGDIOBJ old = SelectObject(dc, use);
                SetTextColor(dc, p.link ? linkCol : (p.active ? active : normal));
                SIZE sz{};
                GetTextExtentPoint32W(dc, p.text.c_str(), (int)p.text.size(), &sz);
                TextOutW(dc, x, y, p.text.c_str(), (int)p.text.size());
                if (p.link)
                    g_linkRect = RECT{ x, y, x + sz.cx, y + sz.cy };
                x += sz.cx;
                lineH = (sz.cy > lineH) ? sz.cy : lineH;
                SelectObject(dc, old);
            }

            // Line 2: 활성 인수의 설명(없으면 함수 설명).
            y += lineH + kGap;
            SelectObject(dc, g_font);
            SetTextColor(dc, normal);
            const std::wstring help = line2Text(g_func, g_argIndex);
            if (!help.empty())
            {
                RECT t{ kPadX, y, rc.right - kPadX, rc.bottom - kPadY };
                DrawTextW(dc, help.c_str(), -1, &t, DT_LEFT | DT_TOP | DT_NOPREFIX);
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
                SelectObject(dc, p.link ? g_fontLink : (p.active ? g_fontBold : g_font));
                SIZE sz{};
                GetTextExtentPoint32W(dc, p.text.c_str(), (int)p.text.size(), &sz);
                sigW += sz.cx;
                lineH = (sz.cy > lineH) ? sz.cy : lineH;
            }
            SelectObject(dc, g_font);
            SIZE hs{};
            const std::wstring help = line2Text(f, argIndex);
            if (!help.empty())
                GetTextExtentPoint32W(dc, help.c_str(), (int)help.size(), &hs);

            SelectObject(dc, old);
            ReleaseDC(nullptr, dc);

            int w = (sigW > hs.cx ? sigW : hs.cx) + kPadX * 2;
            int h = lineH + (help.empty() ? 0 : (kGap + hs.cy)) + kPadY * 2;
            return SIZE{ w, h };
        }

        LRESULT CALLBACK wndProc(HWND h, UINT m, WPARAM w, LPARAM l)
        {
            // Make the whole window a drag handle so the user can reposition it —
            // except the function-name link, which must receive clicks.
            // (WS_EX_NOACTIVATE keeps Excel's cell edit active while dragging.)
            if (m == WM_NCHITTEST)
            {
                POINT pt{ GET_X_LPARAM(l), GET_Y_LPARAM(l) };
                ScreenToClient(h, &pt);
                if (PtInRect(&g_linkRect, pt))
                    return HTCLIENT;
                return HTCAPTION;
            }

            if (m == WM_SETCURSOR)
            {
                POINT pt{};
                GetCursorPos(&pt);
                ScreenToClient(h, &pt);
                if (PtInRect(&g_linkRect, pt))
                {
                    SetCursor(LoadCursor(nullptr, IDC_HAND));
                    return TRUE;
                }
            }

            if (m == WM_LBUTTONUP)
            {
                POINT pt{ GET_X_LPARAM(l), GET_Y_LPARAM(l) };
                if (PtInRect(&g_linkRect, pt) && !g_func.name.empty())
                    ShellExecuteW(nullptr, L"open", manualUrl(g_func.name).c_str(),
                                  nullptr, nullptr, SW_SHOWNORMAL);
                return 0;
            }

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
            g_fontLink = makeFont(true, /*underline*/ true);

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
        if (g_fontLink) { DeleteObject(g_fontLink); g_fontLink = nullptr; }
        UnregisterClassW(kClass, GetModuleHandleW(nullptr));
    }
}
