// DatePicker.cpp — Win32 날짜 선택 달력 다이얼로그 (EGToolsVB Fm_DatePicker.vb 포팅).
//
// WinForms 폼을 리소스 없는 오너드로우 Win32 팝업으로 재구현했다.
//  - 상단 네비게이션(◀◀ 전년 / ◀ 전월 / "2026년 8월" / ▶ 다음달 / ▶▶ 내년),
//    화살표는 VB DrawNavArrow와 동일하게 GDI 폴리곤으로 직접 그린다.
//  - 요일 헤더 7칸(일=빨강, 토=파랑) + 날짜 42칸(6주): 이번 달 진하게,
//    전/다음 달 회색, 오늘 테두리, 선택일 노란 배경, 일요일·공휴일 빨강.
//  - 공휴일은 KoreanHolidays 공용 헬퍼(builtinHolidays + mergeApiHolidays).
//    마우스 오버 시 공휴일 이름을 하단 상태 줄에 표시(툴팁 창 대신 —
//    구현 단순성 우선, TrackMouseEvent 불필요).
//  - 모달은 DialogBox가 아닌 자체 루프(CreateWindow + EnableWindow(owner,FALSE)
//    + GetMessage): 날짜 클릭/Esc/X로 종료. 취소 시 아무 것도 하지 않는다.
//  - 스마트 위치: ActivePane.PointsToScreenPixelsX/Y로 활성 셀 오른쪽에 배치,
//    화면 경계를 넘으면 반전/클램프. COM 실패 시 Excel 창 중앙(VB SetSmartPosition).
//  - 창 클래스는 XLL 모듈 핸들 기준 1회 등록(정적 XLL 수명주기 규약, plan/18).
//    재로드로 RegisterClass가 실패(이미 존재)해도 정상 동작. UnregisterClass 불필요.

#include "Commands.h"
#include "CmdCommon.h"
#include "../../functions/KoreanHolidays.h"

#include <windowsx.h>

#include <map>
#include <string>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

namespace egtools::commands
{
    namespace
    {
        // ── 날짜 직렬(OADate, 1900 기준) 로컬 헬퍼 — CmdCalendar.cpp와 동일 규약 ──
        int serialFromYMD(int y, int m, int d)
        {
            y -= m <= 2;
            const int era = (y >= 0 ? y : y - 399) / 400;
            const unsigned yoe = (unsigned)(y - era * 400);
            const unsigned doy = (153u * (unsigned)(m + (m > 2 ? -3 : 9)) + 2) / 5
                                 + (unsigned)d - 1;
            const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
            return era * 146097 + (int)doe - 719468 + 25569;
        }

        void ymdFromSerial(int serial, int& y, int& m, int& d)
        {
            int z = serial - 25569 + 719468;
            const int era = (z >= 0 ? z : z - 146096) / 146097;
            const unsigned doe = (unsigned)(z - era * 146097);
            const unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
            const int yy = (int)yoe + era * 400;
            const unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
            const unsigned mp = (5 * doy + 2) / 153;
            d = (int)(doy - (153 * mp + 2) / 5 + 1);
            m = (int)(mp + (mp < 10 ? 3 : -9));
            y = yy + (m <= 2);
        }

        int dowOf(int serial)       // 0=일 … 6=토 (serial 1 = 1900-01-01 = 일요일)
        {
            return (serial % 7 + 13) % 7;
        }

        int todaySerial()
        {
            SYSTEMTIME t; GetLocalTime(&t);
            return serialFromYMD(t.wYear, t.wMonth, t.wDay);
        }

        const wchar_t* kMonthEn[12] = { L"January", L"February", L"March", L"April",
                                        L"May", L"June", L"July", L"August",
                                        L"September", L"October", L"November", L"December" };

        // ── 색상(VB Fm_DatePicker와 동일 계열) ──
        constexpr COLORREF kColText     = RGB(0, 0, 0);
        constexpr COLORREF kColGray     = RGB(176, 176, 176);   // 전/다음 달
        constexpr COLORREF kColRed      = RGB(220, 0, 0);       // 일요일·공휴일
        constexpr COLORREF kColBlue     = RGB(0, 0, 220);       // 토요일
        constexpr COLORREF kColSelBack  = RGB(255, 255, 0);     // 선택일(노랑)
        constexpr COLORREF kColTodayBox = RGB(220, 60, 60);     // 오늘 테두리
        constexpr COLORREF kColArrow    = RGB(70, 70, 70);      // 네비 화살표
        constexpr COLORREF kColStatus   = RGB(96, 96, 96);      // 상태 줄
        constexpr COLORREF kColLine     = RGB(200, 200, 200);   // 구분선

        // ── 다이얼로그 상태(창 GWLP_USERDATA로 연결, 수명은 showCalendar 스택) ──
        struct PickerState
        {
            bool ko = false;
            int  viewYear = 2000, viewMonth = 1;    // 표시 중인 연/월
            int  selSerial = 0;                     // 선택(하이라이트) 날짜
            int  today = 0;
            bool done = false, ok = false;
            int  dpi = 96;

            HFONT fontReg = nullptr, fontBold = nullptr, fontTitle = nullptr;

            int daySerials[42] = {};
            std::map<int, std::map<int, std::wstring>> holidayCache;   // 연도 → serial→이름
            std::wstring status;                    // 하단 상태 줄(공휴일 이름)

            // 레이아웃(클라이언트 픽셀)
            int clientW = 0, clientH = 0;
            RECT rcPrevYear{}, rcPrevMon{}, rcNextMon{}, rcNextYear{}, rcTitle{};
            RECT rcDow{}, rcDays[42] = {}, rcToday{}, rcStatus{};
        };

        int scaleDpi(const PickerState& st, int v) { return MulDiv(v, st.dpi, 96); }

        // GetDpiForWindow는 Win10 1607+ — 동적 로드, 없으면 화면 DC 기준(96 폴백).
        UINT dpiForWindow(HWND hwnd)
        {
            typedef UINT(WINAPI* FnGetDpiForWindow)(HWND);
            HMODULE u32 = GetModuleHandleW(L"user32.dll");
            auto fn = u32 ? (FnGetDpiForWindow)GetProcAddress(u32, "GetDpiForWindow")
                          : nullptr;
            if (fn && hwnd)
            {
                const UINT d = fn(hwnd);
                if (d) return d;
            }
            HDC dc = GetDC(nullptr);
            const UINT d = dc ? (UINT)GetDeviceCaps(dc, LOGPIXELSX) : 96;
            if (dc) ReleaseDC(nullptr, dc);
            return d ? d : 96;
        }

        HMODULE thisModule()
        {
            HMODULE mod = nullptr;
            GetModuleHandleExW(
                GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                reinterpret_cast<LPCWSTR>(&thisModule), &mod);
            return mod;
        }

        // ── 레이아웃(96dpi 기준 논리 좌표 → dpi 스케일) ──
        void computeLayout(PickerState& st)
        {
            const int margin  = scaleDpi(st, 8);
            const int cellW   = scaleDpi(st, 34);
            const int cellH   = scaleDpi(st, 24);
            const int nav     = scaleDpi(st, 24);
            const int headerH = scaleDpi(st, 28);
            const int dowH    = scaleDpi(st, 20);
            const int todayH  = scaleDpi(st, 20);
            const int statusH = scaleDpi(st, 16);
            const int gap     = scaleDpi(st, 4);
            const int g2      = scaleDpi(st, 2);

            st.clientW = margin * 2 + cellW * 7;
            int y = margin;

            const int navY = y + (headerH - nav) / 2;
            st.rcPrevYear = { margin, navY, margin + nav, navY + nav };
            st.rcPrevMon  = { margin + nav + g2, navY, margin + nav * 2 + g2, navY + nav };
            st.rcNextYear = { st.clientW - margin - nav, navY, st.clientW - margin, navY + nav };
            st.rcNextMon  = { st.clientW - margin - nav * 2 - g2, navY,
                              st.clientW - margin - nav - g2, navY + nav };
            st.rcTitle    = { st.rcPrevMon.right + g2, y, st.rcNextMon.left - g2, y + headerH };
            y += headerH + gap;

            st.rcDow = { margin, y, margin + cellW * 7, y + dowH };
            y += dowH + g2;

            for (int i = 0; i < 42; ++i)
            {
                const int r = i / 7, c = i % 7;
                st.rcDays[i] = { margin + c * cellW, y + r * cellH,
                                 margin + (c + 1) * cellW, y + (r + 1) * cellH };
            }
            y += cellH * 6 + gap;

            st.rcToday = { margin, y, margin + cellW * 7, y + todayH };
            y += todayH;
            st.rcStatus = { margin, y, margin + cellW * 7, y + statusH };
            y += statusH + margin;
            st.clientH = y;
        }

        // 42칸(6주) 그리드 재계산: 1일이 속한 주의 일요일부터 연속 42일.
        void rebuildGrid(PickerState& st)
        {
            const int first = serialFromYMD(st.viewYear, st.viewMonth, 1);
            const int start = first - dowOf(first);
            for (int i = 0; i < 42; ++i) st.daySerials[i] = start + i;
        }

        // 월 이동(dm=±1 월, ±12 연). 1900~9999 클램프.
        void moveMonth(PickerState& st, int dm)
        {
            int y = st.viewYear, m = st.viewMonth + dm;
            while (m < 1)  { m += 12; --y; }
            while (m > 12) { m -= 12; ++y; }
            if (y < 1900) { y = 1900; m = 1; }
            if (y > 9999) { y = 9999; m = 12; }
            st.viewYear = y; st.viewMonth = m;
            rebuildGrid(st);
        }

        // 연도별 공휴일(내장 계산 + API 병합). VB UpdateHolidays 대응 —
        // mergeApiHolidays는 자체 세션 캐시가 있어 네트워크 재조회는 없다.
        const std::map<int, std::wstring>& holidaysOfYear(PickerState& st, int year)
        {
            auto it = st.holidayCache.find(year);
            if (it == st.holidayCache.end())
            {
                std::map<int, std::wstring> list;
                try
                {
                    egtools::dates::builtinHolidays(year, /*mayDay*/true, st.ko, list);
                    egtools::dates::mergeApiHolidays(year, list);   // 키 있으면 조용히 병합
                }
                catch (...) {}
                it = st.holidayCache.emplace(year, std::move(list)).first;
            }
            return it->second;
        }

        // serial의 공휴일 이름(없으면 빈 문자열). 그리드가 연 경계를 걸치는
        // 12월/1월 뷰를 위해 셀의 실제 연도로 조회한다.
        std::wstring holidayNameOf(PickerState& st, int serial)
        {
            int y, m, d; ymdFromSerial(serial, y, m, d);
            const auto& list = holidaysOfYear(st, y);
            const auto it = list.find(serial);
            return it == list.end() ? L"" : it->second;
        }

        // ── 그리기 ──
        // VB DrawNavArrow/DrawTriangle 대응: direction -1=◀ 1=▶, dbl=이중(연도).
        void drawArrow(HDC dc, const PickerState& st, const RECT& rc, int direction, bool dbl)
        {
            const int aw = scaleDpi(st, 5), half = scaleDpi(st, 5);
            const int cx = (rc.left + rc.right) / 2, cy = (rc.top + rc.bottom) / 2;
            HBRUSH br = CreateSolidBrush(kColArrow);
            HPEN pen = CreatePen(PS_SOLID, 1, kColArrow);
            HGDIOBJ ob = SelectObject(dc, br), op = SelectObject(dc, pen);
            auto tri = [&](int bx)
            {
                POINT pts[3];
                if (direction < 0)  // ◀: 뾰족한 끝이 왼쪽
                {
                    pts[0] = { bx, cy }; pts[1] = { bx + aw, cy - half };
                    pts[2] = { bx + aw, cy + half };
                }
                else                // ▶: 뾰족한 끝이 오른쪽
                {
                    pts[0] = { bx + aw, cy }; pts[1] = { bx, cy - half };
                    pts[2] = { bx, cy + half };
                }
                Polygon(dc, pts, 3);
            };
            if (dbl) { tri(cx - aw); tri(cx); }
            else     tri(cx - aw / 2);
            SelectObject(dc, ob); SelectObject(dc, op);
            DeleteObject(br); DeleteObject(pen);
        }

        void drawTextIn(HDC dc, const RECT& rc, const std::wstring& s, COLORREF col,
                        HFONT font, UINT fmt = DT_CENTER | DT_VCENTER | DT_SINGLELINE)
        {
            HGDIOBJ of = SelectObject(dc, font);
            SetTextColor(dc, col);
            RECT r = rc;
            DrawTextW(dc, s.c_str(), (int)s.size(), &r, fmt | DT_NOPREFIX);
            SelectObject(dc, of);
        }

        void paintAll(HDC dc, PickerState& st, const RECT& client)
        {
            FillRect(dc, &client, (HBRUSH)(COLOR_WINDOW + 1));
            SetBkMode(dc, TRANSPARENT);

            // 네비게이션 + 제목
            drawArrow(dc, st, st.rcPrevYear, -1, true);
            drawArrow(dc, st, st.rcPrevMon,  -1, false);
            drawArrow(dc, st, st.rcNextMon,   1, false);
            drawArrow(dc, st, st.rcNextYear,  1, true);
            wchar_t title[48];
            if (st.ko)
                swprintf_s(title, L"%d년 %d월", st.viewYear, st.viewMonth);
            else
                swprintf_s(title, L"%s %d", kMonthEn[st.viewMonth - 1], st.viewYear);
            drawTextIn(dc, st.rcTitle, title, kColText, st.fontTitle);

            // 요일 헤더(일=빨강, 토=파랑)
            static const wchar_t* kDowKo[7] = { L"일", L"월", L"화", L"수", L"목", L"금", L"토" };
            static const wchar_t* kDowEn[7] = { L"Su", L"Mo", L"Tu", L"We", L"Th", L"Fr", L"Sa" };
            const int cellW = (st.rcDow.right - st.rcDow.left) / 7;
            for (int i = 0; i < 7; ++i)
            {
                RECT r = { st.rcDow.left + i * cellW, st.rcDow.top,
                           st.rcDow.left + (i + 1) * cellW, st.rcDow.bottom };
                const COLORREF col = i == 0 ? kColRed : (i == 6 ? kColBlue : kColText);
                drawTextIn(dc, r, st.ko ? kDowKo[i] : kDowEn[i], col, st.fontBold);
            }

            // 날짜 42칸
            for (int i = 0; i < 42; ++i)
            {
                const int serial = st.daySerials[i];
                int y, m, d; ymdFromSerial(serial, y, m, d);
                const bool inMonth = (y == st.viewYear && m == st.viewMonth);
                const int dow = dowOf(serial);
                const std::wstring holiday = inMonth ? holidayNameOf(st, serial) : L"";

                // 선택일: 노란 배경(VB selectedBackColor)
                if (serial == st.selSerial)
                {
                    HBRUSH br = CreateSolidBrush(kColSelBack);
                    FillRect(dc, &st.rcDays[i], br);
                    DeleteObject(br);
                }
                // 오늘: 테두리 강조(2px)
                if (serial == st.today)
                {
                    HBRUSH br = CreateSolidBrush(kColTodayBox);
                    RECT r = st.rcDays[i];
                    FrameRect(dc, &r, br);
                    InflateRect(&r, -1, -1);
                    FrameRect(dc, &r, br);
                    DeleteObject(br);
                }

                COLORREF col;
                if (!inMonth)                       col = kColGray;
                else if (!holiday.empty() || dow == 0) col = kColRed;
                else if (dow == 6)                  col = kColBlue;
                else                                col = kColText;
                drawTextIn(dc, st.rcDays[i], std::to_wstring(d), col,
                           !holiday.empty() ? st.fontBold : st.fontReg);
            }

            // 구분선 + "오늘: yyyy-mm-dd" (클릭 시 오늘로 이동)
            {
                HPEN pen = CreatePen(PS_SOLID, 1, kColLine);
                HGDIOBJ op = SelectObject(dc, pen);
                MoveToEx(dc, st.rcToday.left, st.rcToday.top - scaleDpi(st, 2), nullptr);
                LineTo(dc, st.rcToday.right, st.rcToday.top - scaleDpi(st, 2));
                SelectObject(dc, op);
                DeleteObject(pen);
            }
            int ty, tm, td; ymdFromSerial(st.today, ty, tm, td);
            wchar_t todayText[48];
            swprintf_s(todayText, L"%s: %04d-%02d-%02d",
                       st.ko ? L"오늘" : L"Today", ty, tm, td);
            drawTextIn(dc, st.rcToday, todayText, kColText, st.fontBold);

            // 상태 줄(마우스 오버한 공휴일 이름)
            if (!st.status.empty())
                drawTextIn(dc, st.rcStatus, st.status, kColStatus, st.fontReg,
                           DT_LEFT | DT_VCENTER | DT_SINGLELINE);
        }

        // ── 히트 테스트 ──
        int dayIndexAt(const PickerState& st, POINT pt)
        {
            for (int i = 0; i < 42; ++i)
                if (PtInRect(&st.rcDays[i], pt)) return i;
            return -1;
        }

        bool clickableAt(const PickerState& st, POINT pt)
        {
            return PtInRect(&st.rcPrevYear, pt) || PtInRect(&st.rcPrevMon, pt) ||
                   PtInRect(&st.rcNextMon, pt) || PtInRect(&st.rcNextYear, pt) ||
                   PtInRect(&st.rcToday, pt) || dayIndexAt(st, pt) >= 0;
        }

        // ── 창 프로시저 ──
        LRESULT CALLBACK pickerProc(HWND wnd, UINT msg, WPARAM wp, LPARAM lp)
        {
            auto* st = (PickerState*)GetWindowLongPtrW(wnd, GWLP_USERDATA);
            switch (msg)
            {
            case WM_NCCREATE:
                SetWindowLongPtrW(wnd, GWLP_USERDATA,
                                  (LONG_PTR)((CREATESTRUCTW*)lp)->lpCreateParams);
                return DefWindowProcW(wnd, msg, wp, lp);

            case WM_ERASEBKGND:
                return 1;   // WM_PAINT에서 더블 버퍼로 전체를 그림(깜빡임 방지)

            case WM_PAINT:
            {
                if (!st) break;
                PAINTSTRUCT ps;
                HDC hdc = BeginPaint(wnd, &ps);
                RECT rc; GetClientRect(wnd, &rc);
                HDC mem = CreateCompatibleDC(hdc);
                HBITMAP bmp = CreateCompatibleBitmap(hdc, rc.right, rc.bottom);
                HGDIOBJ ob = SelectObject(mem, bmp);
                paintAll(mem, *st, rc);
                BitBlt(hdc, 0, 0, rc.right, rc.bottom, mem, 0, 0, SRCCOPY);
                SelectObject(mem, ob);
                DeleteObject(bmp);
                DeleteDC(mem);
                EndPaint(wnd, &ps);
                return 0;
            }

            case WM_MOUSEMOVE:
            {
                if (!st) break;
                const POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
                std::wstring status;
                const int i = dayIndexAt(*st, pt);
                if (i >= 0) status = holidayNameOf(*st, st->daySerials[i]);
                if (status != st->status)
                {
                    st->status = std::move(status);
                    InvalidateRect(wnd, &st->rcStatus, FALSE);
                }
                return 0;
            }

            case WM_SETCURSOR:
                if (st && LOWORD(lp) == HTCLIENT)
                {
                    POINT pt; GetCursorPos(&pt); ScreenToClient(wnd, &pt);
                    if (clickableAt(*st, pt))
                    {
                        SetCursor(LoadCursorW(nullptr, IDC_HAND));
                        return TRUE;
                    }
                }
                break;

            case WM_LBUTTONDOWN:
            {
                if (!st) break;
                const POINT pt = { GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
                if (PtInRect(&st->rcPrevYear, pt))      moveMonth(*st, -12);
                else if (PtInRect(&st->rcPrevMon, pt))  moveMonth(*st, -1);
                else if (PtInRect(&st->rcNextMon, pt))  moveMonth(*st, +1);
                else if (PtInRect(&st->rcNextYear, pt)) moveMonth(*st, +12);
                else if (PtInRect(&st->rcToday, pt))    // VB YearMonthLabel_Click: 오늘로 이동
                {
                    st->selSerial = st->today;
                    int y, m, d; ymdFromSerial(st->today, y, m, d);
                    st->viewYear = y; st->viewMonth = m;
                    rebuildGrid(*st);
                }
                else
                {
                    // 날짜 클릭 → 확정 종료(전/다음 달 날짜도 선택 가능 — VB 동일)
                    const int i = dayIndexAt(*st, pt);
                    if (i < 0) return 0;
                    st->selSerial = st->daySerials[i];
                    st->ok = true;
                    st->done = true;
                    return 0;
                }
                InvalidateRect(wnd, nullptr, FALSE);
                return 0;
            }

            case WM_KEYDOWN:
                if (wp == VK_ESCAPE && st)      // VB DtPicker_KeyDown: Esc=취소
                {
                    st->done = true;
                    return 0;
                }
                break;

            case WM_CLOSE:                      // X 버튼=취소
                if (st) st->done = true;
                return 0;
            }
            return DefWindowProcW(wnd, msg, wp, lp);
        }

        constexpr const wchar_t* kClassName = L"EGToolsPPDatePicker";

        // 창 클래스 1회 등록. 재로드 등으로 이미 존재하면 성공으로 간주.
        bool ensureWindowClass()
        {
            static bool registered = false;
            if (registered) return true;
            WNDCLASSW wc{};
            wc.style = CS_HREDRAW | CS_VREDRAW;
            wc.lpfnWndProc = pickerProc;
            wc.hInstance = thisModule();        // XLL 모듈 기준(plan/18 수명주기 규약)
            wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
            wc.lpszClassName = kClassName;
            if (!RegisterClassW(&wc) && GetLastError() != ERROR_CLASS_ALREADY_EXISTS)
                return false;
            registered = true;
            return true;
        }

        // ── 스마트 위치(VB SetSmartPosition + GetExcelCellScreenRect) ──
        // ActivePane.PointsToScreenPixelsX/Y — 줌·틀 고정이 반영된 화면 픽셀 좌표.
        bool paneToPixels(IDispatch* pane, const wchar_t* name, double points, long& out)
        {
            VARIANT a = varDouble(points);
            VARIANT r; VariantInit(&r);
            const bool ok = invokeRaw(pane, name,
                                      DISPATCH_METHOD | DISPATCH_PROPERTYGET, &r, &a, 1);
            bool got = false;
            if (ok)
            {
                VARIANT o; VariantInit(&o);
                if (SUCCEEDED(VariantChangeType(&o, &r, 0, VT_I4)))
                {
                    out = o.lVal;
                    got = true;
                }
                VariantClear(&o);
            }
            VariantClear(&r);
            return got;
        }

        POINT smartPosition(IDispatch* app, HWND owner, int winW, int winH)
        {
            long cellRight = 0, cellLeft = 0, cellTop = 0;
            bool got = false;
            if (app)
            {
                IDispatch* cell = getObject(app, L"ActiveCell");
                if (cell)
                {
                    Releaser rc{ cell };
                    const double left  = getDouble(cell, L"Left", 0);
                    const double top   = getDouble(cell, L"Top", 0);
                    const double width = getDouble(cell, L"Width", 0);
                    IDispatch* actWin = getObject(app, L"ActiveWindow");
                    if (actWin)
                    {
                        Releaser rw{ actWin };
                        IDispatch* pane = getObject(actWin, L"ActivePane");
                        if (pane)
                        {
                            Releaser rp{ pane };
                            got = paneToPixels(pane, L"PointsToScreenPixelsX",
                                               left + width, cellRight)
                               && paneToPixels(pane, L"PointsToScreenPixelsX",
                                               left, cellLeft)
                               && paneToPixels(pane, L"PointsToScreenPixelsY",
                                               top, cellTop);
                        }
                    }
                }
            }

            POINT p{};
            if (!got)
            {
                // 실패 시 Excel 창 중앙(VB CenterScreen 대응)
                RECT ro{};
                if (!owner || !GetWindowRect(owner, &ro))
                    SystemParametersInfoW(SPI_GETWORKAREA, 0, &ro, 0);
                p.x = ro.left + ((ro.right - ro.left) - winW) / 2;
                p.y = ro.top + ((ro.bottom - ro.top) - winH) / 2;
                return p;
            }

            p.x = cellRight + 4;    // 기본: 셀 오른쪽
            p.y = cellTop;

            // 화면(작업 영역) 경계 검사: 오른쪽 초과 → 셀 왼쪽으로 반전, 나머지는 클램프
            MONITORINFO mi{ sizeof(mi) };
            const POINT anchor = { cellRight, cellTop };
            GetMonitorInfoW(MonitorFromPoint(anchor, MONITOR_DEFAULTTONEAREST), &mi);
            const RECT wa = mi.rcWork;
            if (p.x + winW > wa.right)  p.x = cellLeft - winW - 4;
            if (p.y + winH > wa.bottom) p.y = wa.bottom - winH - 8;
            if (p.x < wa.left)          p.x = wa.left + 8;
            if (p.y < wa.top)           p.y = wa.top + 8;
            return p;
        }

        // ── 모달 표시(자체 모달 루프). 확정 시 직렬값, 취소 시 0 반환 ──
        int showCalendar(IDispatch* app, int initialSerial)
        {
            if (!ensureWindowClass()) return 0;

            HWND owner = excelHwnd();

            PickerState st;
            st.ko = egtools::i18n::current() == L"ko";
            st.today = todaySerial();
            st.selSerial = initialSerial > 0 ? initialSerial : st.today;
            int y, m, d; ymdFromSerial(st.selSerial, y, m, d);
            st.viewYear = y; st.viewMonth = m;
            st.dpi = (int)dpiForWindow(owner);
            computeLayout(st);
            rebuildGrid(st);

            const int pt9 = -MulDiv(9, st.dpi, 72), pt11 = -MulDiv(11, st.dpi, 72);
            st.fontReg = CreateFontW(pt9, 0, 0, 0, FW_NORMAL, 0, 0, 0, DEFAULT_CHARSET,
                                     0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");
            st.fontBold = CreateFontW(pt9, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET,
                                      0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");
            st.fontTitle = CreateFontW(pt11, 0, 0, 0, FW_BOLD, 0, 0, 0, DEFAULT_CHARSET,
                                       0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");

            const DWORD style = WS_POPUP | WS_CAPTION | WS_SYSMENU;
            const DWORD exStyle = WS_EX_TOOLWINDOW | WS_EX_DLGMODALFRAME;
            RECT wr = { 0, 0, st.clientW, st.clientH };
            AdjustWindowRectEx(&wr, style, FALSE, exStyle);
            const int winW = wr.right - wr.left, winH = wr.bottom - wr.top;
            const POINT pos = smartPosition(app, owner, winW, winH);

            HWND wnd = CreateWindowExW(exStyle, kClassName,
                                       st.ko ? L"날짜 선택" : L"Date Picker",
                                       style, pos.x, pos.y, winW, winH,
                                       owner, nullptr, thisModule(), &st);
            if (!wnd)
            {
                DeleteObject(st.fontReg); DeleteObject(st.fontBold);
                DeleteObject(st.fontTitle);
                return 0;
            }

            // 자체 모달: 오너 비활성화 → 메시지 루프 → 오너 복원(파괴 전 필수 —
            // 파괴 후 복원하면 포커스가 다른 앱으로 넘어간다)
            EnableWindow(owner, FALSE);
            ShowWindow(wnd, SW_SHOW);
            SetForegroundWindow(wnd);
            SetFocus(wnd);

            MSG msg;
            while (!st.done)
            {
                const BOOL r = GetMessageW(&msg, nullptr, 0, 0);
                if (r <= 0)
                {
                    if (r == 0) PostQuitMessage((int)msg.wParam);   // WM_QUIT 재전달
                    break;
                }
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }

            EnableWindow(owner, TRUE);
            SetForegroundWindow(owner);
            DestroyWindow(wnd);
            DeleteObject(st.fontReg); DeleteObject(st.fontBold);
            DeleteObject(st.fontTitle);

            return st.ok ? st.selSerial : 0;
        }

        // 셀 값(VT_DATE/숫자)을 날짜 직렬로. 실패 시 false (CmdCalendar 규약).
        bool dateSerialOf(const VARIANT& v, int& serial)
        {
            VARIANT out; VariantInit(&out);
            if (SUCCEEDED(VariantChangeType(&out, const_cast<VARIANT*>(&v), 0, VT_DATE)))
            {
                serial = (int)out.date;
                VariantClear(&out);
                return serial > 0 && serial <= 2958465;     // 1900-01-01 ~ 9999-12-31
            }
            VariantClear(&out);
            return false;
        }
    }

    // 달력에서 날짜를 골라 현재 셀에 입력(컨텍스트 메뉴). VB EGDP2.DatePicker 대응:
    // 활성 셀 값이 날짜면 그 날짜로, 아니면 오늘로 초기화. 취소 시 아무 것도 안 함.
    void pickDateIntoCell()
    {
        try
        {
            IDispatch* app = appDisp();
            if (!app) return;
            Releaser ra{ app };

            // 초기 날짜: 활성 셀 값이 날짜(직렬값)면 사용, 아니면 오늘
            int initial = 0;
            {
                IDispatch* cell = getObject(app, L"ActiveCell");
                if (cell)
                {
                    Releaser rc{ cell };
                    VARIANT v; VariantInit(&v);
                    if (getVar(cell, L"Value", &v))
                    {
                        int s = 0;
                        if (!isEmptyVal(v) && dateSerialOf(v, s)) initial = s;
                        VariantClear(&v);
                    }
                }
            }

            const int picked = showCalendar(app, initial);
            if (picked <= 0) return;    // 취소(Esc/X)

            IDispatch* cell = getObject(app, L"ActiveCell");
            if (!cell) return;          // 셀이 아닌 상태(차트 시트 등)면 무시
            Releaser rc{ cell };
            VARIANT v; VariantInit(&v);
            v.vt = VT_DATE;
            v.date = (DATE)picked;
            putVar(cell, L"Value", v);
        }
        catch (...) {}
    }
}
