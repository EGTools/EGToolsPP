// CmdCalendar.cpp — 달력/일정표 명령 (EGToolsVB C04_Scheduler.vb 포팅, plan/23).
// MakeMonthlyCalendar / MakeAnnualCalendarA·B / MakeWeeklyCalendar /
// MakeDailyCalendarA·B(+CopyTaskFromYesterday). 연간 A/B와 일일 A/B는 레이아웃
// 파라미터/분기만 다른 공통 구현이며 산출 시트 레이아웃은 VB와 동일하다.
// VB의 명언(GetDailyWord) 셀은 제거(사용자 결정). 공휴일은 KoreanHolidays 공용
// 헬퍼(builtinHolidays + mergeApiHolidays)를 사용하고 API 직접 호출은 없다.

#include "Commands.h"
#include "CmdCommon.h"
#include "Dialogs.h"
#include "../../functions/KoreanHolidays.h"

#include <climits>
#include <map>
#include <set>
#include <string>
#include <vector>

namespace egtools::commands
{
    namespace
    {
        // ── Excel 상수 ──
        constexpr long xlHAlignCenter   = -4108;
        constexpr long xlHAlignLeft     = -4131;
        constexpr long xlVAlignCenter   = -4108;
        constexpr long xlContinuous     = 1;
        constexpr long xlHairline       = 1;
        constexpr long xlThin           = 2;
        constexpr long xlLineStyleNone  = -4142;
        constexpr long xlDiagonalDown   = 5;
        constexpr long xlDiagonalUp     = 6;
        constexpr long xlInsideHorizontal = 12;
        constexpr long xlPortrait       = 1;
        constexpr long xlLandscape      = 2;

        // ── 날짜 직렬(OADate, 1900 기준) 로컬 헬퍼 ──
        // days_from_civil(그레고리력) 기반: serial = 1970-01-01 기준 일수 + 25569.
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

        int dowOf(int serial)       // 0=일 … 6=토
        {
            return (serial % 7 + 13) % 7;   // serial 1(1900-01-01)=일요일 규약
        }
        int monthOf(int serial)
        {
            int y, m, d; ymdFromSerial(serial, y, m, d); return m;
        }
        int dayOf(int serial)
        {
            int y, m, d; ymdFromSerial(serial, y, m, d); return d;
        }
        int daysInMonth(int y, int m)
        {
            static const int k[12] = { 31,28,31,30,31,30,31,31,30,31,30,31 };
            if (m == 2 && ((y % 4 == 0 && y % 100 != 0) || y % 400 == 0)) return 29;
            return k[m - 1];
        }
        // 주차번호(수요일 기준): (d+3일).DayOfYear \ 7 + 1 (VB 동일)
        int weekNumOf(int serial)
        {
            const int s3 = serial + 3;
            int y, m, d; ymdFromSerial(s3, y, m, d);
            const int doy = s3 - serialFromYMD(y, 1, 1) + 1;
            return doy / 7 + 1;
        }

        // "YYYY-MM" / "YYYY-MM-DD" 파싱( '/'와 '.' 구분자 허용, VB TryParse 관례).
        bool parseYearMonth(std::wstring s, int& y, int& m)
        {
            for (auto& ch : s) if (ch == L'/' || ch == L'.') ch = L'-';
            int yy = 0, mm = 0;
            if (swscanf_s(s.c_str(), L"%d-%d", &yy, &mm) != 2) return false;
            if (yy < 1900 || yy > 9999 || mm < 1 || mm > 12) return false;
            y = yy; m = mm;
            return true;
        }
        bool parseDate(std::wstring s, int& y, int& m, int& d)
        {
            for (auto& ch : s) if (ch == L'/' || ch == L'.') ch = L'-';
            int yy = 0, mm = 0, dd = 0;
            if (swscanf_s(s.c_str(), L"%d-%d-%d", &yy, &mm, &dd) != 3) return false;
            if (yy < 1900 || yy > 9999 || mm < 1 || mm > 12) return false;
            if (dd < 1 || dd > daysInMonth(yy, mm)) return false;
            y = yy; m = mm; d = dd;
            return true;
        }

        // ── 표시 문자열 ──
        const wchar_t* kWeekKo[7] = { L"일", L"월", L"화", L"수", L"목", L"금", L"토" };
        const wchar_t* kWeekEn[7] = { L"Sun", L"Mon", L"Tue", L"Wed", L"Thu", L"Fri", L"Sat" };
        const wchar_t* kDayFullKo[7] = { L"일요일", L"월요일", L"화요일", L"수요일",
                                         L"목요일", L"금요일", L"토요일" };
        const wchar_t* kDayFullEn[7] = { L"Sunday", L"Monday", L"Tuesday", L"Wednesday",
                                         L"Thursday", L"Friday", L"Saturday" };
        const wchar_t* kMonthEn[12] = { L"January", L"February", L"March", L"April",
                                        L"May", L"June", L"July", L"August",
                                        L"September", L"October", L"November", L"December" };

        std::wstring monthName(int m, bool ko)  // VB ToString("MMMM", 문화권)
        {
            return ko ? std::to_wstring(m) + L"월" : kMonthEn[m - 1];
        }
        // 음력 "MM/dd" (VB SingleToLunar().Substring(5).Replace("-","/") 대응).
        // 범위(1918~2050) 밖이면 빈 문자열.
        std::wstring lunarShort(int serial)
        {
            int ly, lm, ld;
            if (!egtools::dates::lunarOf(serial, ly, lm, ld)) return L"";
            wchar_t buf[8];
            swprintf_s(buf, L"%02d/%02d", lm, ld);
            return buf;
        }
        std::wstring lunarFull(int serial)      // "y-MM-dd"
        {
            int ly, lm, ld;
            if (!egtools::dates::lunarOf(serial, ly, lm, ld)) return L"";
            wchar_t buf[16];
            swprintf_s(buf, L"%d-%02d-%02d", ly, lm, ld);
            return buf;
        }

        // ── 공휴일(공용 헤더) ──
        std::map<int, std::wstring> holidayList(int year, bool ko)
        {
            std::map<int, std::wstring> list;
            egtools::dates::builtinHolidays(year, /*mayDay*/true, ko, list);
            egtools::dates::mergeApiHolidays(year, list);   // 키 있으면 조용히 병합
            return list;
        }

        // ── COM 소품 ──
        VARIANT varBool(bool v)
        {
            VARIANT a; VariantInit(&a); a.vt = VT_BOOL;
            a.boolVal = v ? VARIANT_TRUE : VARIANT_FALSE;
            return a;
        }

        IDispatch* rangeOf(IDispatch* disp, const std::wstring& addr)
        {
            VARIANT a = varBStr(addr);
            IDispatch* r = getObject(disp, L"Range", &a, 1);
            VariantClear(&a);
            return r;
        }
        // Range(Cells(r1,c1), Cells(r2,c2))
        IDispatch* rangeRC(IDispatch* sh, long r1, long c1, long r2, long c2)
        {
            IDispatch* a = getObjectIdx2(sh, L"Cells", r1, c1);
            IDispatch* b = getObjectIdx2(sh, L"Cells", r2, c2);
            IDispatch* rng = nullptr;
            if (a && b)
            {
                VARIANT v[2];
                VariantInit(&v[0]); v[0].vt = VT_DISPATCH; v[0].pdispVal = a;
                VariantInit(&v[1]); v[1].vt = VT_DISPATCH; v[1].pdispVal = b;
                rng = getObject(sh, L"Range", v, 2);
            }
            if (a) a->Release();
            if (b) b->Release();
            return rng;
        }

        void putDate(IDispatch* rng, int serial)
        {
            VARIANT v; VariantInit(&v); v.vt = VT_DATE; v.date = (DATE)serial;
            putVar(rng, L"Value", v);
        }
        void putText(IDispatch* rng, const std::wstring& s)
        {
            VARIANT v = varBStr(s);
            putVar(rng, L"Value", v);
            VariantClear(&v);
        }
        void putNum(IDispatch* rng, double n)
        {
            VARIANT v; VariantInit(&v); v.vt = VT_R8; v.dblVal = n;
            putVar(rng, L"Value", v);
        }

        void setNF(IDispatch* rng, const std::wstring& f)
        {
            putBStr(rng, L"NumberFormat", f);
        }
        void fontSize(IDispatch* rng, long size)
        {
            if (IDispatch* f = getObject(rng, L"Font"))
            { Releaser r{ f }; putLong(f, L"Size", size); }
        }
        void fontColor(IDispatch* rng, long color)
        {
            if (IDispatch* f = getObject(rng, L"Font"))
            { Releaser r{ f }; putLong(f, L"Color", color); }
        }
        void fontBold(IDispatch* rng)
        {
            if (IDispatch* f = getObject(rng, L"Font"))
            { Releaser r{ f }; putBool(f, L"Bold", true); }
        }
        void interiorColor(IDispatch* rng, long color)
        {
            if (IDispatch* i = getObject(rng, L"Interior"))
            { Releaser r{ i }; putLong(i, L"Color", color); }
        }
        void mergeCells(IDispatch* rng, bool across = false)
        {
            if (across)
            {
                VARIANT a = varBool(true);
                callMethod(rng, L"Merge", &a, 1);
            }
            else
                callMethod(rng, L"Merge");
        }
        // BorderAround(LineStyle:=xlContinuous, Weight:=xlThin) — 위치 인수 4개
        void borderAround(IDispatch* rng)
        {
            VARIANT a[4];
            a[0] = varLong(xlContinuous);
            a[1] = varLong(xlThin);
            VariantInit(&a[2]); a[2].vt = VT_ERROR; a[2].scode = DISP_E_PARAMNOTFOUND;
            VariantInit(&a[3]); a[3].vt = VT_ERROR; a[3].scode = DISP_E_PARAMNOTFOUND;
            callMethod(rng, L"BorderAround", a, 4);
        }
        // Borders 전체 LineStyle/Weight + 대각선 제거(VB 관례 시퀀스)
        void bordersHairline(IDispatch* rng)
        {
            if (IDispatch* bs = getObject(rng, L"Borders"))
            {
                Releaser r{ bs };
                putLong(bs, L"LineStyle", xlContinuous);
                putLong(bs, L"Weight", xlHairline);
            }
            for (long idx : { xlDiagonalDown, xlDiagonalUp })
            {
                if (IDispatch* b = getObjectIdx(rng, L"Borders", idx))
                { Releaser r{ b }; putLong(b, L"LineStyle", xlLineStyleNone); }
            }
        }
        void borderInsideHorizontalHairline(IDispatch* rng)
        {
            if (IDispatch* b = getObjectIdx(rng, L"Borders", xlInsideHorizontal))
            {
                Releaser r{ b };
                putLong(b, L"LineStyle", xlContinuous);
                putLong(b, L"Weight", xlHairline);
            }
        }
        void hAlign(IDispatch* rng, long v) { putLong(rng, L"HorizontalAlignment", v); }
        void vAlign(IDispatch* rng, long v) { putLong(rng, L"VerticalAlignment", v); }
        void shrinkToFit(IDispatch* rng)    { putBool(rng, L"ShrinkToFit", true); }

        // 셀 메모 추가(+선택적 자동 크기/숨김)
        void addComment(IDispatch* cell, const std::wstring& text,
                        bool autoSize, bool hide)
        {
            VARIANT a = varBStr(text);
            VARIANT r; VariantInit(&r);
            const bool ok = callMethod(cell, L"AddComment", &a, 1, &r);
            VariantClear(&a);
            IDispatch* cm = (ok && r.vt == VT_DISPATCH) ? r.pdispVal : nullptr;
            if (!cm) { VariantClear(&r); return; }
            Releaser rc{ cm };
            if (hide) putBool(cm, L"Visible", false);
            if (autoSize)
            {
                if (IDispatch* shp = getObject(cm, L"Shape"))
                {
                    Releaser rs{ shp };
                    if (IDispatch* tf = getObject(shp, L"TextFrame"))
                    { Releaser rt{ tf }; putBool(tf, L"AutoSize", true); }
                }
            }
        }

        // Sheets.Add(After:=ActiveSheet) — Before는 VT_ERROR/PARAMNOTFOUND
        IDispatch* addSheetAfterActive(IDispatch* ad, IDispatch* wb)
        {
            IDispatch* cur = getObject(ad, L"ActiveSheet");
            if (!cur) return nullptr;
            Releaser rCur{ cur };
            IDispatch* sheets = getObject(wb, L"Sheets");
            if (!sheets) return nullptr;
            Releaser rs{ sheets };
            VARIANT a[2];
            VariantInit(&a[0]); a[0].vt = VT_ERROR; a[0].scode = DISP_E_PARAMNOTFOUND;
            VariantInit(&a[1]); a[1].vt = VT_DISPATCH; a[1].pdispVal = cur;
            VARIANT r; VariantInit(&r);
            if (!callMethod(sheets, L"Add", a, 2, &r)) return nullptr;
            if (r.vt == VT_DISPATCH && r.pdispVal) return r.pdispVal;   // 소유권 이전
            VariantClear(&r);
            return nullptr;
        }

        std::wstring upperCase(std::wstring s)
        {
            for (auto& ch : s) ch = towupper(ch);
            return s;
        }

        // VB GetSheetName(NewName:=True) 대응: base가 비어 있으면 그대로,
        // 이미 있으면 base_1, base_2 … 첫 빈 이름.
        std::wstring newSheetName(IDispatch* wb, const std::wstring& base)
        {
            std::set<std::wstring> names;
            if (IDispatch* sheets = getObject(wb, L"Sheets"))
            {
                Releaser rs{ sheets };
                const long n = getLong(sheets, L"Count", 0);
                for (long i = 1; i <= n; ++i)
                {
                    IDispatch* sh = getObjectIdx(sheets, L"Item", i);
                    if (!sh) continue;
                    Releaser r{ sh };
                    names.insert(upperCase(getBStr(sh, L"Name")));
                }
            }
            if (!names.count(upperCase(base))) return base;
            for (int i = 1; ; ++i)
            {
                std::wstring cand = base + L"_" + std::to_wstring(i);
                if (!names.count(upperCase(cand))) return cand;
            }
        }

        void pageSetup(IDispatch* sh, long orientation)
        {
            IDispatch* ps = getObject(sh, L"PageSetup");
            if (!ps) return;
            Releaser r{ ps };
            putLong(ps, L"Orientation", orientation);
            putBool(ps, L"Zoom", false);
            putLong(ps, L"FitToPagesWide", 1);
            putLong(ps, L"FitToPagesTall", 1);
        }

        // 시트 활성화 + 셀 활성화 + 틀고정/눈금선(VB 공통 마무리)
        void finishView(IDispatch* ad, IDispatch* sh, const std::wstring& activateAddr,
                        bool unfreezeFirst)
        {
            callMethod(sh, L"Activate");
            if (IDispatch* r = rangeOf(sh, activateAddr))
            { Releaser rr{ r }; callMethod(r, L"Activate"); }
            IDispatch* win = getObject(ad, L"ActiveWindow");
            if (!win) return;
            Releaser rw{ win };
            if (unfreezeFirst) putBool(win, L"FreezePanes", false);
            putBool(win, L"FreezePanes", true);
            putBool(win, L"DisplayGridlines", false);
        }

        // 쓰기용 2차원 SAFEARRAY(VT_ARRAY|VT_VARIANT, 1-base, 열 우선 저장).
        // CmdCommon의 SafeArr2D는 읽기 전용이라 별도로 둔다.
        struct Arr2DOut
        {
            SAFEARRAY* sa = nullptr;
            VARIANT* data = nullptr;
            long rows = 0;

            bool create(long r, long c)
            {
                SAFEARRAYBOUND b[2] = { { (ULONG)r, 1 }, { (ULONG)c, 1 } };
                sa = SafeArrayCreate(VT_VARIANT, 2, b);
                if (!sa) return false;
                if (FAILED(SafeArrayAccessData(sa, (void**)&data)))
                { data = nullptr; return false; }
                rows = r;
                return true;    // 요소는 VT_EMPTY로 초기화됨
            }
            VARIANT& at(long r, long c) { return data[(size_t)(c - 1) * rows + (r - 1)]; }
            void setDate(long r, long c, int serial)
            { VARIANT& v = at(r, c); v.vt = VT_DATE; v.date = (DATE)serial; }
            void setStr(long r, long c, const std::wstring& s)
            { VARIANT& v = at(r, c); v.vt = VT_BSTR; v.bstrVal = SysAllocString(s.c_str()); }

            bool writeValue(IDispatch* rng)
            {
                if (data) { SafeArrayUnaccessData(sa); data = nullptr; }
                VARIANT v; VariantInit(&v);
                v.vt = VT_ARRAY | VT_VARIANT; v.parray = sa;
                return putVar(rng, L"Value", v);
            }
            ~Arr2DOut()
            {
                if (data) SafeArrayUnaccessData(sa);
                if (sa) SafeArrayDestroy(sa);
            }
        };

        // C1 값(VT_DATE/숫자/문자)을 날짜 직렬로. 실패 시 false.
        bool dateSerialOf(const VARIANT& v, int& serial)
        {
            VARIANT out; VariantInit(&out);
            if (SUCCEEDED(VariantChangeType(&out, const_cast<VARIANT*>(&v), 0, VT_DATE)))
            {
                serial = (int)out.date;
                VariantClear(&out);
                return serial > 0;
            }
            if (SUCCEEDED(VariantChangeType(&out, const_cast<VARIANT*>(&v), 0, VT_R8)))
            {
                serial = (int)out.dblVal;
                VariantClear(&out);
                return serial > 0;
            }
            return false;
        }

        constexpr const wchar_t* kCheckboxNF = L"[=1]\"✅\";[=0]\"⬜\";\"\"";

        // 0/1 목록 유효성(xlValidateList) — DailyA/B 체크박스 열
        void addCheckboxValidation(IDispatch* rng)
        {
            setNF(rng, kCheckboxNF);
            if (IDispatch* valid = getObject(rng, L"Validation"))
            {
                Releaser rv{ valid };
                VARIANT a[4];
                a[0] = varLong(3);              // xlValidateList
                a[1] = varLong(1);              // xlValidAlertStop
                a[2] = varLong(1);              // xlBetween
                a[3] = varBStr(L"0,1");
                callMethod(valid, L"Add", a, 4);
                VariantClear(&a[3]);
            }
        }

        // CopyTaskFromYesterday (VB C04_Scheduler.vb:886) — 모든 시트의 A1 메모가
        // "EGCalendar"로 시작하는 시트 중 가장 최근 과거(C1 날짜) 시트에서
        // 미완료(B열 = 빈칸/0/"□") 업무(C열)를 모은다.
        std::vector<std::wstring> copyTasksFromYesterday(IDispatch* wb,
                                                         int todaySerial, long lastRow)
        {
            std::vector<std::wstring> out;
            IDispatch* sheets = getObject(wb, L"Sheets");
            if (!sheets) return out;
            Releaser rs{ sheets };

            IDispatch* best = nullptr;
            int bestSerial = INT_MIN;
            const long n = getLong(sheets, L"Count", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* sh = getObjectIdx(sheets, L"Item", i);
                if (!sh) continue;
                Releaser rSh{ sh };
                IDispatch* a1 = rangeOf(sh, L"A1");     // 차트 시트 등은 실패 → 건너뜀
                if (!a1) continue;
                Releaser rA1{ a1 };
                IDispatch* cm = getObject(a1, L"Comment");
                if (!cm) continue;
                Releaser rCm{ cm };

                bool isCal = false;
                VARIANT tr; VariantInit(&tr);
                if (invokeRaw(cm, L"Text", DISPATCH_METHOD, &tr, nullptr, 0) &&
                    tr.vt == VT_BSTR && tr.bstrVal)
                    isCal = wcsncmp(tr.bstrVal, L"EGCalendar", 10) == 0;
                VariantClear(&tr);
                if (!isCal) continue;

                IDispatch* c1 = rangeOf(sh, L"C1");
                if (!c1) continue;
                Releaser rC1{ c1 };
                VARIANT v; VariantInit(&v);
                getVar(c1, L"Value", &v);
                int serial = 0;
                const bool okDate = dateSerialOf(v, serial);
                VariantClear(&v);
                if (!okDate) continue;

                if (serial < todaySerial && serial > bestSerial)
                {
                    if (best) best->Release();
                    best = sh;
                    best->AddRef();
                    bestSerial = serial;
                }
            }
            if (!best) return out;
            Releaser rBest{ best };

            for (long i = 5; i <= lastRow; ++i)
            {
                IDispatch* c = getObjectIdx2(best, L"Cells", i, 3);
                if (!c) continue;
                Releaser rc{ c };
                VARIANT vc; VariantInit(&vc);
                getVar(c, L"Value2", &vc);
                if (vc.vt == VT_ERROR || isEmptyVal(vc)) { VariantClear(&vc); continue; }
                const std::wstring task = varToString(vc);
                VariantClear(&vc);
                if (task.empty()) continue;

                IDispatch* b = getObjectIdx2(best, L"Cells", i, 2);
                if (!b) continue;
                Releaser rb{ b };
                VARIANT vb; VariantInit(&vb);
                getVar(b, L"Value", &vb);
                bool undone = isEmptyVal(vb);       // VB: Nothing = 0 → True
                if (!undone)
                {
                    VARIANT num; VariantInit(&num);
                    if (SUCCEEDED(VariantChangeType(&num, &vb, 0, VT_R8)))
                        undone = num.dblVal == 0.0;
                    else
                        undone = varToString(vb) == L"□";
                    VariantClear(&num);
                }
                VariantClear(&vb);
                if (undone) out.push_back(task);
            }
            return out;
        }

        // ── 연간 달력 공통 구현 (VB MakeAnnualCalendarA:213 / B:330) ──
        // A: 3열×4행(A:Z, 세로), B: 4열×3행(A:AJ, 가로). 그 외 로직 동일.
        void makeAnnualCalendarImpl(bool typeB)
        {
            using egtools::i18n::t;
            try
            {
                const bool ko = egtools::i18n::current() == L"ko";
                SYSTEMTIME st; GetLocalTime(&st);

                double yv = st.wYear;
                if (!egtools::dialogs::inputNumber(excelHwnd(),
                        t(L"cmd.cal.yearTitle"), t(L"cmd.cal.yearPrompt"),
                        yv, 2000, 2050, st.wYear))
                    return;
                const int year = (int)yv;

                xloil::PauseExcel pause(xloil::thisApp());

                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };
                IDispatch* wb = getObject(ad, L"ActiveWorkbook");
                if (!wb) return;
                Releaser rWb{ wb };
                IDispatch* sh = addSheetAfterActive(ad, wb);
                if (!sh) return;
                Releaser rSh{ sh };

                const auto holidays = holidayList(year, ko);

                const wchar_t* colRange   = typeB ? L"A:AJ"    : L"A:Z";
                const wchar_t* fitRange   = typeB ? L"A1:AJ32" : L"A1:Z40";
                const wchar_t* titleMerge = typeB ? L"K1:Z2"   : L"H1:S2";
                const wchar_t* titleCell  = typeB ? L"K1"      : L"H1";
                const int perRow = typeB ? 4 : 3;

                {
                    VARIANT a = varBStr(colRange);
                    IDispatch* cols = getObject(sh, L"Columns", &a, 1);
                    VariantClear(&a);
                    if (cols) { Releaser r{ cols }; putDouble(cols, L"ColumnWidth", 3); }
                }
                if (IDispatch* r = rangeOf(sh, colRange))
                {
                    Releaser rr{ r };
                    hAlign(r, xlHAlignCenter);
                    vAlign(r, xlVAlignCenter);
                }
                if (IDispatch* r = rangeOf(sh, fitRange))
                { Releaser rr{ r }; shrinkToFit(r); }
                if (IDispatch* r = rangeOf(sh, titleMerge))
                { Releaser rr{ r }; mergeCells(r); }
                if (IDispatch* r = rangeOf(sh, titleCell))
                {
                    Releaser rr{ r };
                    putDate(r, serialFromYMD(year, 1, 1));
                    setNF(r, L"yyyy");
                    fontSize(r, 20);
                }

                for (int month = 1; month <= 12; ++month)
                {
                    const long iRow = ((month - 1) / perRow) * 9 + 4;
                    const long iCol = ((month - 1) % perRow) * 9 + 2;

                    if (IDispatch* title = rangeRC(sh, iRow, iCol, iRow, iCol + 6))
                    {
                        Releaser rt{ title };
                        mergeCells(title);
                        hAlign(title, xlHAlignCenter);
                        interiorColor(title, RGB(0, 255, 0));
                        putText(title, monthName(month, ko));
                    }

                    long R = iRow + 2;
                    long C = dowOf(serialFromYMD(year, month, 1));
                    const int dim = daysInMonth(year, month);
                    for (int day = 1; day <= dim; ++day)
                    {
                        const int serial = serialFromYMD(year, month, day);
                        IDispatch* cell = getObjectIdx2(sh, L"Cells", R, iCol + C);
                        if (cell)
                        {
                            Releaser rc{ cell };
                            putNum(cell, day);
                            setNF(cell, L"d");
                            const int dw = dowOf(serial);
                            if (dw == 0)      fontColor(cell, RGB(255, 0, 0));
                            else if (dw == 6) fontColor(cell, RGB(0, 0, 255));
                            auto it = holidays.find(serial);
                            if (it != holidays.end())
                            {
                                interiorColor(cell, RGB(255, 255, 0));
                                fontColor(cell, RGB(255, 0, 0));
                                addComment(cell, it->second, /*autoSize*/true, /*hide*/false);
                            }
                        }
                        if (++C == 7) { C = 0; ++R; }
                    }

                    R = iRow + 1;   // 요일 헤더
                    for (long c = 0; c <= 6; ++c)
                    {
                        IDispatch* cell = getObjectIdx2(sh, L"Cells", R, iCol + c);
                        if (!cell) continue;
                        Releaser rc{ cell };
                        putText(cell, ko ? kWeekKo[c] : kWeekEn[c]);
                        fontColor(cell, c == 0 ? RGB(255, 0, 0)
                                     : c == 6 ? RGB(0, 0, 255) : RGB(0, 0, 0));
                    }
                }

                pageSetup(sh, typeB ? xlLandscape : xlPortrait);
                finishView(ad, sh, L"A3", /*unfreezeFirst*/true);
            }
            catch (...) {}
        }

        // ── 일일 일정표 공통 구현 (VB MakeDailyCalendarA:749 / B:927) ──
        void makeDailyCalendarImpl(bool typeB)
        {
            using egtools::i18n::t;
            try
            {
                const bool ko = egtools::i18n::current() == L"ko";
                SYSTEMTIME st; GetLocalTime(&st);
                wchar_t def[16];
                swprintf_s(def, L"%04d-%02d-%02d", st.wYear, st.wMonth, st.wDay);
                std::wstring in = def;
                if (!egtools::dialogs::inputText(excelHwnd(),
                        t(L"cmd.cal.dayTitle"), t(L"cmd.cal.dayPrompt"), in))
                    return;
                int y = 0, m = 0, d = 0;
                if (!parseDate(in, y, m, d)) return;    // 파싱 실패 → 조용히 종료
                const int serial = serialFromYMD(y, m, d);

                xloil::PauseExcel pause(xloil::thisApp());

                IDispatch* ad = appDisp();
                if (!ad) return;
                Releaser rApp{ ad };
                IDispatch* wb = getObject(ad, L"ActiveWorkbook");
                if (!wb) return;
                Releaser rWb{ wb };
                IDispatch* sh = addSheetAfterActive(ad, wb);
                if (!sh) return;
                Releaser rSh{ sh };

                const auto holidays = holidayList(y, ko);
                const long taskLast = typeB ? 32 : 19;      // 주요 업무 마지막 행

                // ── 열 너비 ──
                auto setColWidth = [&](const wchar_t* addr, double w)
                {
                    VARIANT a = varBStr(addr);
                    IDispatch* cols = getObject(sh, L"Columns", &a, 1);
                    VariantClear(&a);
                    if (!cols)
                    {
                        // "A:A,E:E" 같은 다중 영역은 Range로
                        if (IDispatch* r = rangeOf(sh, addr))
                        { Releaser rr{ r }; putDouble(r, L"ColumnWidth", w); }
                        return;
                    }
                    Releaser rc{ cols };
                    putDouble(cols, L"ColumnWidth", w);
                };
                if (!typeB)
                {
                    if (IDispatch* r = rangeOf(sh, L"A:A,E:E,G:G"))
                    { Releaser rr{ r }; putDouble(r, L"ColumnWidth", 1); }
                    setColWidth(L"B:B", 3.8);
                    setColWidth(L"C:D", 35);
                    setColWidth(L"F:F", 70);
                }
                else
                {
                    if (IDispatch* r = rangeOf(sh, L"A:A,G:G"))
                    { Releaser rr{ r }; putDouble(r, L"ColumnWidth", 1); }
                    if (IDispatch* r = rangeOf(sh, L"B:B,D:D,E:E"))
                    { Releaser rr{ r }; putDouble(r, L"ColumnWidth", 3); }
                    setColWidth(L"C:C", 45);
                    setColWidth(L"F:F", 25);
                    setColWidth(L"H:H", 60);
                }

                // ── 정렬/축소 ──
                if (IDispatch* r = rangeOf(sh, typeB ? L"A4:H4" : L"A4:G4"))
                { Releaser rr{ r }; hAlign(r, xlHAlignCenter); }
                if (IDispatch* r = rangeOf(sh, typeB ? L"A:H" : L"A:G"))
                { Releaser rr{ r }; vAlign(r, xlVAlignCenter); }
                if (IDispatch* r = rangeOf(sh, typeB ? L"A1:H32" : L"A1:G35"))
                { Releaser rr{ r }; shrinkToFit(r); }

                // ── 제목(날짜/요일/음력) ──
                if (IDispatch* r = rangeOf(sh, L"C1:C2"))
                {
                    Releaser rr{ r };
                    mergeCells(r);
                    putDate(r, serial);
                    setNF(r, ko ? L"yyyy년 M월 d일" : L"MMMM d, yyyy");
                    fontSize(r, 20);
                }
                const std::wstring dayName = ko ? kDayFullKo[dowOf(serial)]
                                                : kDayFullEn[dowOf(serial)];
                if (!typeB)
                {
                    if (IDispatch* r = rangeOf(sh, L"D1:D2"))
                    {
                        Releaser rr{ r };
                        mergeCells(r);
                        const std::wstring lunar = lunarFull(serial);
                        putText(r, lunar.empty() ? dayName
                                                 : dayName + L" (" + lunar + L")");
                        fontSize(r, 16);
                    }
                }
                else
                {
                    if (IDispatch* r = rangeOf(sh, L"D1:E1"))
                    {
                        Releaser rr{ r };
                        mergeCells(r);
                        putText(r, ko ? L"(음력)" : L"(Lunar)");
                    }
                    if (IDispatch* r = rangeOf(sh, L"D2:E2"))
                    {
                        Releaser rr{ r };
                        mergeCells(r);
                        putText(r, L"(" + lunarShort(serial) + L")");
                    }
                    if (IDispatch* r = rangeOf(sh, L"F1:F2"))
                    {
                        Releaser rr{ r };
                        putText(r, dayName);
                        mergeCells(r);
                        fontSize(r, 14);
                    }
                }

                // ── 공휴일 표시줄 ──
                if (IDispatch* r = rangeOf(sh, typeB ? L"B3:F3" : L"B3:D3"))
                {
                    Releaser rr{ r };
                    mergeCells(r);
                    auto it = holidays.find(serial);
                    if (it != holidays.end()) putText(r, it->second);
                }

                // ── 주요 업무 ──
                const wchar_t* taskHead = typeB ? L"B4:C4" : L"B4:D4";
                if (IDispatch* r = rangeOf(sh, taskHead))
                {
                    Releaser rr{ r };
                    mergeCells(r);
                    borderAround(r);
                    interiorColor(r, RGB(200, 200, 200));
                    putText(r, ko ? L"주요 업무" : L"Main Tasks");
                }
                if (IDispatch* r = rangeOf(sh, typeB ? L"B5:B32" : L"B5:B19"))
                { Releaser rr{ r }; addCheckboxValidation(r); }

                // 전일 미완료 업무 이월
                const auto tasks = copyTasksFromYesterday(wb, serial, taskLast);
                for (size_t i = 0; i < tasks.size(); ++i)
                {
                    if (IDispatch* b = getObjectIdx2(sh, L"Cells", 5 + (long)i, 2))
                    { Releaser rb{ b }; putNum(b, 0); }
                    if (IDispatch* c = getObjectIdx2(sh, L"Cells", 5 + (long)i, 3))
                    { Releaser rc{ c }; putText(c, tasks[i]); }
                }

                if (!typeB)
                {
                    if (IDispatch* r = rangeOf(sh, L"C5:D19"))
                    {
                        Releaser rr{ r };
                        mergeCells(r, /*across*/true);
                        hAlign(r, xlHAlignLeft);
                    }
                    if (IDispatch* r = rangeOf(sh, L"B5:D19"))
                    {
                        Releaser rr{ r };
                        bordersHairline(r);
                        borderAround(r);
                    }
                    if (IDispatch* r = rangeOf(sh, L"B5:B19"))
                    { Releaser rr{ r }; hAlign(r, xlHAlignCenter); }

                    // 시간 일정(8~21시)
                    if (IDispatch* r = rangeOf(sh, L"B20:D20"))
                    {
                        Releaser rr{ r };
                        mergeCells(r);
                        borderAround(r);
                        hAlign(r, xlHAlignCenter);
                        interiorColor(r, RGB(200, 200, 200));
                        putText(r, ko ? L"시간 일정" : L"Time Schedule");
                    }
                    for (long i = 1; i <= 14; ++i)
                    {
                        if (IDispatch* c = getObjectIdx2(sh, L"Cells", 20 + i, 2))
                        { Releaser rc{ c }; putNum(c, 7 + i); }
                    }
                    if (IDispatch* r = rangeOf(sh, L"B21:D34"))
                    {
                        Releaser rr{ r };
                        bordersHairline(r);
                        borderAround(r);
                    }

                    // 업무 메모
                    if (IDispatch* r = rangeOf(sh, L"F4"))
                    {
                        Releaser rr{ r };
                        borderAround(r);
                        interiorColor(r, RGB(200, 200, 200));
                        putText(r, ko ? L"업무 메모" : L"Task Notes");
                    }
                    if (IDispatch* r = rangeOf(sh, L"F5:F34"))
                    {
                        Releaser rr{ r };
                        bordersHairline(r);
                        borderAround(r);
                    }
                    // (VB의 F1:F3 명언 셀은 제거 — 사용자 결정)
                }
                else
                {
                    if (IDispatch* r = rangeOf(sh, L"B5:C32"))
                    { Releaser rr{ r }; borderAround(r); }
                    if (IDispatch* r = rangeOf(sh, L"B5:B32"))
                    { Releaser rr{ r }; hAlign(r, xlHAlignCenter); }

                    // 시간 일정(8~21시, 2행 병합)
                    if (IDispatch* r = rangeOf(sh, L"D4:F4"))
                    {
                        Releaser rr{ r };
                        mergeCells(r);
                        putText(r, ko ? L"시간 일정" : L"Time Schedule");
                    }
                    for (long i = 1; i <= 28; i += 2)
                    {
                        if (IDispatch* c = getObjectIdx2(sh, L"Cells", 4 + i, 4))
                        { Releaser rc{ c }; putNum(c, 8 + i / 2); }
                        if (IDispatch* r = rangeRC(sh, 4 + i, 4, 5 + i, 4))
                        { Releaser rr{ r }; mergeCells(r); }
                    }
                    if (IDispatch* r = rangeOf(sh, L"D4:F4"))
                    {
                        Releaser rr{ r };
                        borderAround(r);
                        interiorColor(r, RGB(200, 200, 200));
                    }
                    if (IDispatch* r = rangeOf(sh, L"B5:F32"))
                    { Releaser rr{ r }; bordersHairline(r); }
                    if (IDispatch* r = rangeOf(sh, L"B5:C32"))
                    { Releaser rr{ r }; borderAround(r); }
                    if (IDispatch* r = rangeOf(sh, L"D5:F32"))
                    { Releaser rr{ r }; borderAround(r); }

                    // 업무 메모
                    if (IDispatch* r = rangeOf(sh, L"H4"))
                    {
                        Releaser rr{ r };
                        putText(r, ko ? L"업무 메모" : L"Task Notes");
                        borderAround(r);
                        interiorColor(r, RGB(200, 200, 200));
                    }
                    if (IDispatch* r = rangeOf(sh, L"H5:H32"))
                    {
                        Releaser rr{ r };
                        borderAround(r);
                        bordersHairline(r);
                        borderAround(r);
                    }
                    // (VB의 H1:H3 명언 셀은 제거 — 사용자 결정)
                }

                pageSetup(sh, xlLandscape);

                // A1 메모(시트 식별) + 활성화 + 시트 이름(MMdd, 중복 시 _n)
                wchar_t tag[32];
                swprintf_s(tag, L"EGCalendar:%04d-%02d-%02d", y, m, d);
                if (IDispatch* a1 = rangeOf(sh, L"A1"))
                { Releaser ra{ a1 }; addComment(a1, tag, /*autoSize*/false, /*hide*/true); }

                callMethod(sh, L"Activate");
                if (IDispatch* r = rangeOf(sh, L"C5"))
                { Releaser rr{ r }; callMethod(r, L"Activate"); }
                wchar_t base[8];
                swprintf_s(base, L"%02d%02d", m, d);
                putBStr(sh, L"Name", newSheetName(wb, base));

                IDispatch* win = getObject(ad, L"ActiveWindow");
                if (win)
                {
                    Releaser rw{ win };
                    putBool(win, L"FreezePanes", true);
                    putBool(win, L"DisplayGridlines", false);
                }
            }
            catch (...) {}
        }
    }

    // 월간 달력 (VB MakeMonthlyCalendar, C04_Scheduler.vb:447)
    // 42일(6주) 격자를 2차원 배열로 한 번에 기입 후 서식 지정.
    void makeMonthlyCalendar()
    {
        using egtools::i18n::t;
        try
        {
            const bool ko = egtools::i18n::current() == L"ko";
            SYSTEMTIME st; GetLocalTime(&st);
            wchar_t def[16];
            swprintf_s(def, L"%04d-%02d", st.wYear, st.wMonth);
            std::wstring in = def;
            if (!egtools::dialogs::inputText(excelHwnd(),
                    t(L"cmd.cal.monthTitle"), t(L"cmd.cal.monthPrompt"), in))
                return;
            int y = 0, m = 0;
            if (!parseYearMonth(in, y, m)) return;

            xloil::PauseExcel pause(xloil::thisApp());

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* wb = getObject(ad, L"ActiveWorkbook");
            if (!wb) return;
            Releaser rWb{ wb };

            auto holidays = holidayList(y, ko);
            if (m == 1)         // 앞달에 걸친 주의 성탄절
                holidays[serialFromYMD(y - 1, 12, 25)] = ko ? L"성탄절" : L"Christmas";
            else if (m == 12)   // 뒷달에 걸친 주의 새해첫날
                holidays[serialFromYMD(y + 1, 1, 1)] = ko ? L"새해첫날" : L"New Year's Day";

            // ── 격자 값 구성(VB vCells(49,39) → Resize(49,39)와 동일 크기) ──
            constexpr long kRows = 49, kCols = 39;
            struct MCell { int kind = 0; int serial = 0; std::wstring s; }; // 0빈 1날짜 2문자
            std::vector<std::vector<MCell>> g((size_t)kRows + 1,
                                              std::vector<MCell>((size_t)kCols + 2));

            const int firstSerial = serialFromYMD(y, m, 1);
            g[1][17] = { 1, firstSerial, L"" };     // Q1 (연)
            g[1][20] = { 1, firstSerial, L"" };     // T1 (월)

            for (int i = 0; i <= 6; ++i)            // 요일 헤더(행 4)
                g[4][(size_t)i * 5 + 2] = { 2, 0, ko ? kWeekKo[i] : kWeekEn[i] };

            const int start = firstSerial - dowOf(firstSerial);
            for (int i = 0; i <= 41; ++i)
            {
                const long row = (i / 7) * 7 + 5;
                const long col = (i % 7) * 5 + 2;
                const int serial = start + i;
                g[row][col] = { 1, serial, L"" };
                if (col == 2)                       // 주차번호(수요일 기준)
                    g[row][1] = { 2, 0, L"W" + std::to_wstring(weekNumOf(serial)) };
                auto it = holidays.find(serial);
                if (it != holidays.end())
                    g[row][col + 1] = { 2, 0, it->second };
                int ly, lm, ld;                     // 음력 1·11·21일만 표시(VB 동일)
                if (egtools::dates::lunarOf(serial, ly, lm, ld) && ld % 10 == 1)
                    g[row][col + 4] = { 2, 0, L"(" + lunarShort(serial) + L")" };
            }

            // ── 새 시트에 한 번에 기입 ──
            IDispatch* sh = addSheetAfterActive(ad, wb);
            if (!sh) return;
            Releaser rSh{ sh };
            {
                Arr2DOut arr;
                if (!arr.create(kRows, kCols)) return;
                for (long r = 1; r <= kRows; ++r)
                    for (long c = 1; c <= kCols; ++c)
                    {
                        const MCell& mc = g[r][c];
                        if (mc.kind == 1)      arr.setDate(r, c, mc.serial);
                        else if (mc.kind == 2) arr.setStr(r, c, mc.s);
                    }
                IDispatch* a1 = rangeOf(sh, L"A1");
                if (!a1) return;
                Releaser ra{ a1 };
                IDispatch* dest = rangeResize(a1, kRows, kCols);
                if (!dest) return;
                Releaser rd{ dest };
                arr.writeValue(dest);
            }

            // ── 서식 ──
            {
                VARIANT a = varBStr(L"A:AJ");
                IDispatch* cols = getObject(sh, L"Columns", &a, 1);
                VariantClear(&a);
                if (cols) { Releaser r{ cols }; putDouble(cols, L"ColumnWidth", 4); }
            }
            if (IDispatch* r = rangeOf(sh, L"A:AJ"))
            {
                Releaser rr{ r };
                hAlign(r, xlHAlignCenter);
                vAlign(r, xlVAlignCenter);
            }
            if (IDispatch* r = rangeOf(sh, L"A1:AJ60"))
            { Releaser rr{ r }; shrinkToFit(r); }
            if (IDispatch* r = rangeOf(sh, L"Q1:S2"))
            { Releaser rr{ r }; mergeCells(r); }
            if (IDispatch* r = rangeOf(sh, L"Q1"))
            { Releaser rr{ r }; setNF(r, ko ? L"yyyy년" : L"yyyy"); }
            if (IDispatch* r = rangeOf(sh, L"T1:U2"))
            { Releaser rr{ r }; mergeCells(r); }
            if (IDispatch* r = rangeOf(sh, L"T1"))
            {
                Releaser rr{ r };
                putDate(r, firstSerial);            // VB: T1 = Q1.Value
                setNF(r, ko ? L"M월" : L"MMMM");
            }
            if (IDispatch* r = rangeOf(sh, L"Q1:U2"))
            { Releaser rr{ r }; fontSize(r, 20); }

            // 주차제목(요일 헤더) 병합+테두리
            for (long iC = 2; iC <= 36; iC += 5)
            {
                if (IDispatch* r = rangeRC(sh, 4, iC, 4, iC + 4))
                {
                    Releaser rr{ r };
                    mergeCells(r);
                    borderAround(r);
                }
            }

            // 일자별 정리(주 단위 7행 × 5열 블록)
            for (long iR = 5; iR <= 40; iR += 7)
            {
                for (long iC = 2; iC <= 36; iC += 5)
                {
                    const MCell& dc = g[iR][iC];
                    const bool inMonth = dc.kind == 1 && monthOf(dc.serial) == m;
                    IDispatch* dayCell = getObjectIdx2(sh, L"Cells", iR, iC);
                    if (!dayCell) continue;
                    Releaser rdc{ dayCell };

                    if (inMonth)
                    {
                        setNF(dayCell, L"D");
                        fontSize(dayCell, 12);
                        fontBold(dayCell);
                        const int dw = dowOf(dc.serial);
                        if (dw == 0)      fontColor(dayCell, RGB(255, 0, 0));
                        else if (dw == 6) fontColor(dayCell, RGB(0, 0, 255));
                    }
                    else
                    {
                        if (iC == 2 && iR > 5)      // 월에 속하지 않는 주 → 행 삭제
                        {
                            if (IDispatch* clr = rangeRC(sh, iR, 1, iR, 40))
                            { Releaser rc{ clr }; callMethod(clr, L"Clear"); }
                            break;
                        }
                        putBStr(dayCell, L"NumberFormatLocal", L"M/D");
                        if (IDispatch* r = rangeRC(sh, iR, iC, iR, iC + 4))
                        {
                            Releaser rr{ r };
                            fontSize(r, 8);
                            fontColor(r, RGB(128, 128, 128));
                        }
                    }
                    // 일자 이름(공휴일명) 병합
                    if (IDispatch* r = rangeRC(sh, iR, iC + 1, iR, iC + 3))
                    { Releaser rr{ r }; mergeCells(r); }
                    // 일정 내용 영역
                    if (IDispatch* r = rangeRC(sh, iR + 1, iC, iR + 6, iC + 4))
                    {
                        Releaser rr{ r };
                        mergeCells(r, /*across*/true);
                        hAlign(r, xlHAlignLeft);
                    }
                    // 공휴일인 경우 강조
                    if (g[iR][iC + 1].kind == 2)
                    {
                        if (IDispatch* r = rangeRC(sh, iR, iC, iR, iC + 4))
                        { Releaser rr{ r }; fontColor(r, RGB(255, 0, 0)); }
                        if (IDispatch* r = rangeRC(sh, iR, iC, iR + 6, iC + 4))
                        { Releaser rr{ r }; interiorColor(r, RGB(255, 255, 220)); }
                    }
                    // 일 전체 테두리
                    if (IDispatch* r = rangeRC(sh, iR, iC, iR + 6, iC + 4))
                    { Releaser rr{ r }; borderAround(r); }
                }
            }

            pageSetup(sh, xlLandscape);
            finishView(ad, sh, L"B5", /*unfreezeFirst*/false);
        }
        catch (...) {}
    }

    // 연간 달력 A — 3열×4행 세로형 (VB MakeAnnualCalendarA, C04_Scheduler.vb:213)
    void makeAnnualCalendarA() { makeAnnualCalendarImpl(false); }
    // 연간 달력 B — 4열×3행 가로형 (VB MakeAnnualCalendarB, C04_Scheduler.vb:330)
    void makeAnnualCalendarB() { makeAnnualCalendarImpl(true); }

    // 주간 일정표 (VB MakeWeeklyCalendar, C04_Scheduler.vb:612)
    void makeWeeklyCalendar()
    {
        using egtools::i18n::t;
        try
        {
            const bool ko = egtools::i18n::current() == L"ko";
            SYSTEMTIME st; GetLocalTime(&st);
            wchar_t def[16];
            swprintf_s(def, L"%04d-%02d-%02d", st.wYear, st.wMonth, st.wDay);
            std::wstring in = def;
            if (!egtools::dialogs::inputText(excelHwnd(),
                    t(L"cmd.cal.weekTitle"), t(L"cmd.cal.weekPrompt"), in))
                return;
            int y = 0, m = 0, d = 0;
            if (!parseDate(in, y, m, d)) return;
            const int ownerSerial = serialFromYMD(y, m, d);
            const int startSerial = ownerSerial - dowOf(ownerSerial);   // 일요일

            xloil::PauseExcel pause(xloil::thisApp());

            IDispatch* ad = appDisp();
            if (!ad) return;
            Releaser rApp{ ad };
            IDispatch* wb = getObject(ad, L"ActiveWorkbook");
            if (!wb) return;
            Releaser rWb{ wb };
            IDispatch* sh = addSheetAfterActive(ad, wb);
            if (!sh) return;
            Releaser rSh{ sh };

            auto holidays = holidayList(y, ko);
            if (monthOf(startSerial) == 1 && dayOf(startSerial) < 7)
                holidays[serialFromYMD(y - 1, 12, 25)] = ko ? L"성탄절" : L"Christmas";
            else if (monthOf(startSerial) == 12 && dayOf(startSerial) > 23)
                holidays[serialFromYMD(y + 1, 1, 1)] = ko ? L"새해첫날" : L"New Year's Day";

            {
                VARIANT a = varBStr(L"A:AK");
                IDispatch* cols = getObject(sh, L"Columns", &a, 1);
                VariantClear(&a);
                if (cols) { Releaser r{ cols }; putDouble(cols, L"ColumnWidth", 4); }
            }
            if (IDispatch* r = rangeOf(sh, L"A:AK"))
            {
                Releaser rr{ r };
                hAlign(r, xlHAlignCenter);
                vAlign(r, xlVAlignCenter);
            }
            if (IDispatch* r = rangeOf(sh, L"A1:AK40"))
            { Releaser rr{ r }; shrinkToFit(r); }

            // ── 제목줄 ──
            auto mergeAndDate = [&](const wchar_t* mergeAddr, const wchar_t* cellAddr,
                                    int serial, const std::wstring& nf)
            {
                if (IDispatch* r = rangeOf(sh, mergeAddr))
                { Releaser rr{ r }; mergeCells(r); }
                if (IDispatch* r = rangeOf(sh, cellAddr))
                {
                    Releaser rr{ r };
                    putDate(r, serial);
                    setNF(r, nf);
                }
            };
            mergeAndDate(L"C1:E2", L"C1", ownerSerial, ko ? L"yyyy년" : L"yyyy");
            mergeAndDate(L"F1:G2", L"F1", ownerSerial, ko ? L"M월" : L"MMMM");
            mergeAndDate(L"N1:R2", L"N1", startSerial, L"yyyy-MM-dd");
            if (IDispatch* r = rangeOf(sh, L"S1:S2"))
            { Releaser rr{ r }; mergeCells(r); }
            if (IDispatch* r = rangeOf(sh, L"S1"))
            { Releaser rr{ r }; putText(r, L"~"); }
            mergeAndDate(L"T1:X2", L"T1", startSerial + 6, L"yyyy-MM-dd");
            if (IDispatch* r = rangeOf(sh, L"Y1:AA2"))
            { Releaser rr{ r }; mergeCells(r); }
            if (IDispatch* r = rangeOf(sh, L"Y1"))
            {
                Releaser rr{ r };
                putText(r, L"(" + std::to_wstring(weekNumOf(startSerial)) +
                           (ko ? L"주차" : L"W") + L")");
            }
            if (IDispatch* r = rangeOf(sh, L"C1,F1,N1,S1,T1,Y1"))
            { Releaser rr{ r }; fontSize(r, 20); }

            // ── 왼쪽 제목열 ──
            if (IDispatch* r = rangeOf(sh, L"B4"))
            { Releaser rr{ r }; putText(r, ko ? L"요일" : L"WeekDay"); }
            if (IDispatch* r = rangeOf(sh, L"B5"))
            { Releaser rr{ r }; putText(r, ko ? L"일자" : L"Date"); }
            if (IDispatch* r = rangeOf(sh, L"B11"))
            { Releaser rr{ r }; putText(r, ko ? L"시간" : L"Time"); }
            for (long iR = 1; iR <= 27; ++iR)       // 08:00 ~ 21:00, 30분 간격
            {
                const int minutes = 8 * 60 + (int)(iR - 1) * 30;
                wchar_t buf[8];
                swprintf_s(buf, L"%02d:%02d", minutes / 60, minutes % 60);
                if (IDispatch* c = getObjectIdx2(sh, L"Cells", 11 + iR, 2))
                { Releaser rc{ c }; putText(c, buf); }
            }

            // ── 요일별 열(5열 블록 × 7) ──
            for (int i = 0; i <= 6; ++i)
            {
                const long iRow = 4;
                const long iCol = (long)i * 5 + 3;
                const int thisSerial = startSerial + i;

                // 요일
                if (IDispatch* r = rangeRC(sh, iRow, iCol, iRow, iCol + 4))
                {
                    Releaser rr{ r };
                    mergeCells(r, /*across*/true);
                    borderAround(r);
                }
                if (IDispatch* c = getObjectIdx2(sh, L"Cells", iRow, iCol))
                {
                    Releaser rc{ c };
                    putText(c, std::wstring(ko ? kWeekKo[i] : kWeekEn[i]) +
                               (ko ? L"요일" : L""));
                }
                // 날짜
                if (IDispatch* r = rangeRC(sh, iRow + 1, iCol, iRow + 1, iCol + 2))
                { Releaser rr{ r }; mergeCells(r, true); }
                if (IDispatch* c = getObjectIdx2(sh, L"Cells", iRow + 1, iCol))
                {
                    Releaser rc{ c };
                    putDate(c, thisSerial);
                    setNF(c, L"yyyy-MM-dd");
                }
                // 음력
                if (IDispatch* r = rangeRC(sh, iRow + 1, iCol + 3, iRow + 1, iCol + 4))
                { Releaser rr{ r }; mergeCells(r, true); }
                if (IDispatch* c = getObjectIdx2(sh, L"Cells", iRow + 1, iCol + 3))
                { Releaser rc{ c }; putText(c, L"(" + lunarShort(thisSerial) + L")"); }
                // 주말표시
                const int dw = dowOf(thisSerial);
                if (dw == 0 || dw == 6)
                {
                    if (IDispatch* r = rangeRC(sh, iRow, iCol, iRow + 1, iCol + 4))
                    {
                        Releaser rr{ r };
                        fontColor(r, dw == 0 ? RGB(255, 0, 0) : RGB(0, 0, 255));
                    }
                }
                // 공휴일 표시
                if (IDispatch* r = rangeRC(sh, iRow + 2, iCol, iRow + 2, iCol + 4))
                { Releaser rr{ r }; mergeCells(r, true); }
                auto it = holidays.find(thisSerial);
                if (it != holidays.end())
                {
                    if (IDispatch* c = getObjectIdx2(sh, L"Cells", iRow + 2, iCol))
                    { Releaser rc{ c }; putText(c, it->second); }
                    if (IDispatch* r = rangeRC(sh, iRow, iCol, iRow + 2, iCol + 4))
                    {
                        Releaser rr{ r };
                        interiorColor(r, RGB(255, 255, 220));
                        fontColor(r, RGB(255, 0, 0));
                    }
                }
                if (IDispatch* r = rangeRC(sh, iRow, iCol, iRow + 2, iCol + 4))
                { Releaser rr{ r }; borderAround(r); }
                // 일정내용 공간
                if (IDispatch* r = rangeRC(sh, iRow + 3, iCol, iRow + 7, iCol + 4))
                {
                    Releaser rr{ r };
                    mergeCells(r, true);
                    borderAround(r);
                }
                // 시간 일정표 공간
                if (IDispatch* r = rangeRC(sh, iRow + 8, iCol, iRow + 34, iCol + 4))
                {
                    Releaser rr{ r };
                    mergeCells(r, true);
                    hAlign(r, xlHAlignLeft);
                    borderAround(r);
                    borderInsideHorizontalHairline(r);
                }
            }

            pageSetup(sh, xlLandscape);
            finishView(ad, sh, L"A7", /*unfreezeFirst*/false);
        }
        catch (...) {}
    }

    // 일일 일정표 A — 2단(업무/시간 + 메모) 가로형 (VB MakeDailyCalendarA:749)
    void makeDailyCalendarA() { makeDailyCalendarImpl(false); }
    // 일일 일정표 B — 3단(업무 | 시간 | 메모) 가로형 (VB MakeDailyCalendarB:927)
    void makeDailyCalendarB() { makeDailyCalendarImpl(true); }
}
