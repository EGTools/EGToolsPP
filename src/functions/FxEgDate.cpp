// FxEgDate.cpp — EGTools 전용 날짜 함수 (EGToolsVB FX3_Date.vb + C04 휴일 로직 포팅, plan/20 E4).
//   TOLUNAR / TOSOLAR — 양력↔음력(한국 음력, 918~2050). 데이터는 .NET
//     KoreanLunisolarCalendar에서 생성한 KoreanLunarTable.h 사용.
//   KOREANHOLIDAYS    — 대한민국 공휴일 목록(대체공휴일 규칙 포함, 내장 계산).
//   NETWORKHOUR       — 휴식시간을 반영한 순 근무시간.
//   WEEKNUMOFMONTH / MONTHBHYWEEK — 시작·기준 요일에 따른 월 주차/주차 기준 월.
//   TODATETIME        — 한글·한자 날짜/시간 문자열 → 날짜·시간 값.
// KOREANHOLIDAYS는 [api_key] 인수(또는 저장 키)가 있으면 공공데이터포털 특일정보
// API를 조회해 임시공휴일·선거일 등 내장 계산에 없는 항목을 병합한다(plan/23
// §3-4 — "내장과 동일 결과"라던 종전 제외 판단은 임시공휴일을 놓친 것이라 번복).
// 공휴일/음력 계산 코어는 KoreanHolidays.h로도 노출되어 리본 달력이 공유한다.

#include "../core/Registry.h"
#include "../core/Apply.h"
#include "../core/ArrayUtil.h"
#include "../core/Spill.h"
#include "../core/I18n.h"
#include "KoreanLunarTable.h"
#include "KoreanHolidays.h"

#include <windows.h>
#include <winhttp.h>
#pragma comment(lib, "winhttp.lib")

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <map>
#include <regex>
#include <string>
#include <vector>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // ---- 그레고리력 <-> 직렬값 (OADate: 0 = 1899-12-30) ----------------

        long long daysFromCivil(int y, int m, int d)
        {
            y -= m <= 2;
            const long long era = (y >= 0 ? y : y - 399) / 400;
            const unsigned yoe = (unsigned)(y - era * 400);
            const unsigned doy = (153u * (unsigned)(m + (m > 2 ? -3 : 9)) + 2) / 5 + (unsigned)d - 1;
            const unsigned doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
            return era * 146097 + (long long)doe - 719468;
        }
        void civilFromDays(long long z, int& y, int& m, int& d)
        {
            z += 719468;
            const long long era = (z >= 0 ? z : z - 146096) / 146097;
            const unsigned doe = (unsigned)(z - era * 146097);
            const unsigned yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
            y = (int)(yoe + era * 400);
            const unsigned doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
            const unsigned mp = (5 * doy + 2) / 153;
            d = (int)(doy - (153 * mp + 2) / 5 + 1);
            m = (int)(mp + (mp < 10 ? 3 : -9));
            y += m <= 2;
        }
        const long long kOAEpoch = daysFromCivil(1899, 12, 30);
        int toSerial(int y, int m, int d) { return (int)(daysFromCivil(y, m, d) - kOAEpoch); }
        void fromSerial(int serial, int& y, int& m, int& d)
        { civilFromDays(kOAEpoch + serial, y, m, d); }

        // 요일: 1=일 ~ 7=토 (VB Weekday(date, vbSunday)).
        int weekday1(int serial)
        {
            const int r = serial % 7;   // 0=토, 1=일, 2=월 … 6=금
            return r == 0 ? 7 : r;
        }

        // ---- 음력 변환 -----------------------------------------------------

        int lunarMonthDays(const LunarYear& ly, int m /*1..months*/)
        { return (ly.mask & (1 << (m - 1))) ? 30 : 29; }

        // 양력 직렬값 → 음력 (연, 1..13 월 인덱스, 일). 범위 밖 false.
        bool solarToLunar(int serial, int& ly, int& lm, int& ld)
        {
            const int n = (int)(sizeof(kLunarYears) / sizeof(kLunarYears[0]));
            if (serial < kLunarYears[0].serial) return false;
            int lo = 0, hi = n;   // serial이 속한 마지막 연 시작
            while (lo < hi)
            {
                const int mid = (lo + hi) / 2;
                if (kLunarYears[mid].serial <= serial) lo = mid + 1;
                else hi = mid;
            }
            const int yi = lo - 1;
            if (yi < 0 || yi >= n) return false;
            const LunarYear& Y = kLunarYears[yi];
            int off = serial - Y.serial;
            for (int m = 1; m <= Y.months; ++m)
            {
                const int len = lunarMonthDays(Y, m);
                if (off < len) { ly = kLunarFirstYear + yi; lm = m; ld = off + 1; return true; }
                off -= len;
            }
            return false;   // 다음 해 범위(테이블 끝)
        }

        // 음력 (연, 1..13 월 인덱스, 일) → 양력 직렬값. 실패 false.
        bool lunarToSolar(int ly, int lm, int ld, int& serial)
        {
            if (ly < kLunarFirstYear || ly > kLunarLastYear) return false;
            const LunarYear& Y = kLunarYears[ly - kLunarFirstYear];
            if (lm < 1 || lm > Y.months || ld < 1 || ld > lunarMonthDays(Y, lm)) return false;
            int off = 0;
            for (int m = 1; m < lm; ++m) off += lunarMonthDays(Y, m);
            serial = Y.serial + off + ld - 1;
            return true;
        }

        // 셀 값 → 양력 직렬값 (직렬번호 / "yyyy-mm-dd"류 / yyyymmdd 숫자).
        bool argToSerial(const ExcelObj& v, int& serial)
        {
            if (v.type() == ExcelType::Num || v.type() == ExcelType::Int)
            {
                const double d = v.get<double>(0.0);
                if (d > 0 && d < 10000101) { serial = (int)d; return serial > 0; }
                const long long n = (long long)d;
                const int y = (int)(n / 10000), m = (int)(n / 100 % 100), dd = (int)(n % 100);
                if (y < 100 || m < 1 || m > 12 || dd < 1 || dd > 31) return false;
                serial = toSerial(y, m, dd);
                return true;
            }
            if (v.type() == ExcelType::Str)
            {
                const std::wstring s = v.toString();
                int y, m, d;
                if (swscanf_s(s.c_str(), L"%d-%d-%d", &y, &m, &d) == 3 ||
                    swscanf_s(s.c_str(), L"%d/%d/%d", &y, &m, &d) == 3 ||
                    swscanf_s(s.c_str(), L"%d.%d.%d", &y, &m, &d) == 3)
                {
                    if (m < 1 || m > 12 || d < 1 || d > 31) return false;
                    serial = toSerial(y, m, d);
                    return true;
                }
                if (s.size() == 8 && std::all_of(s.begin(), s.end(), ::iswdigit))
                {
                    y = _wtoi(s.substr(0, 4).c_str());
                    m = _wtoi(s.substr(4, 2).c_str());
                    d = _wtoi(s.substr(6, 2).c_str());
                    if (m < 1 || m > 12 || d < 1 || d > 31) return false;
                    serial = toSerial(y, m, d);
                    return true;
                }
            }
            return false;
        }

        ExcelObj toLunarOne(const ExcelObj& v)
        {
            int serial;
            if (!argToSerial(v, serial)) return ExcelObj(CellError::Value);
            int ly, lm, ld;
            if (!solarToLunar(serial, ly, lm, ld)) return ExcelObj(CellError::Value);
            const LunarYear& Y = kLunarYears[ly - kLunarFirstYear];
            const bool isLeap = Y.leap > 0 && lm == Y.leap;
            int disp = lm;
            if (Y.leap > 0 && lm >= Y.leap) disp -= 1;
            wchar_t buf[32];
            swprintf_s(buf, L"%d-%02d-%02d", ly, disp, ld);
            std::wstring out = buf;
            if (isLeap)
                out += egtools::i18n::current() == L"ko" ? L"(윤)" : L"(LM)";
            return ExcelObj(std::wstring_view(out));
        }

        ExcelObj toSolarOne(const ExcelObj& v)
        {
            int y = 0, m = 0, d = 0;
            bool leap = false;

            if (v.type() == ExcelType::Str)
            {
                std::wstring s = v.toString();
                std::wstring up = s;
                for (auto& c : up) c = (wchar_t)std::towupper(c);
                leap = s.find(L'윤') != std::wstring::npos ||
                       up.find(L"LM") != std::wstring::npos;
                // 윤달 표기 제거.
                std::wstring t;
                for (size_t i = 0; i < s.size(); ++i)
                {
                    if (s[i] == L'윤') continue;
                    if (s[i] == L'(' || s[i] == L')') continue;
                    if ((s[i] == L'L' || s[i] == L'l') && i + 1 < s.size() &&
                        (s[i + 1] == L'M' || s[i + 1] == L'm')) { ++i; continue; }
                    t.push_back(s[i]);
                }
                if (swscanf_s(t.c_str(), L"%d-%d-%d", &y, &m, &d) != 3 &&
                    swscanf_s(t.c_str(), L"%d/%d/%d", &y, &m, &d) != 3 &&
                    swscanf_s(t.c_str(), L"%d.%d.%d", &y, &m, &d) != 3)
                    return ExcelObj(CellError::Value);
            }
            else if (v.type() == ExcelType::Num || v.type() == ExcelType::Int)
            {
                int serial;
                if (!argToSerial(v, serial)) return ExcelObj(CellError::Value);
                fromSerial(serial, y, m, d);
            }
            else
                return ExcelObj(CellError::Value);

            if (y < kLunarFirstYear || y > kLunarLastYear) return ExcelObj(CellError::Value);
            const LunarYear& Y = kLunarYears[y - kLunarFirstYear];
            // 표시 월(윤달 무시) → 1..13 인덱스.
            int idx = m;
            if (Y.leap > 0)
            {
                if (leap && m == Y.leap - 1) idx = Y.leap;
                else if (m >= Y.leap) idx = m + 1;
            }
            int serial;
            if (!lunarToSolar(y, idx, d, serial)) return ExcelObj(CellError::Value);
            return ExcelObj((double)serial);
        }

        // ---- KOREANHOLIDAYS ------------------------------------------------

        bool lunarNewYearSerial(int year, int lm, int ld, int& serial)
        {
            // 표시 월 기준(윤달 아님).
            if (year < kLunarFirstYear || year > kLunarLastYear) return false;
            const LunarYear& Y = kLunarYears[year - kLunarFirstYear];
            int idx = lm;
            if (Y.leap > 0 && lm >= Y.leap) idx = lm + 1;
            return lunarToSolar(year, idx, ld, serial);
        }

        // 내장 공휴일 계산 본체 — KOREANHOLIDAYS UDF와 리본 달력이 공유
        // (KoreanHolidays.h egtools::dates::builtinHolidays가 이 함수로 위임).
        void buildHolidayList(int year, bool mayDay, bool ko,
                              std::map<int, std::wstring>& list)
        {
            auto isHol = [&](int s)
            {
                return list.count(s) > 0 || s % 7 == 0 /*토*/ || s % 7 == 1 /*일*/;
            };
            auto add = [&](int s, const std::wstring& name)
            {
                auto it = list.find(s);
                if (it == list.end()) list[s] = name;
                else it->second += L", " + name;
            };
            auto addSub = [&](int s, const std::wstring& name)
            {
                int i = 1;
                while (isHol(s + i)) ++i;
                add(s + i, name + (ko ? L" 대체공휴일" : L" Substitute Holiday"));
            };

            int d;
            add(toSerial(year, 1, 1), ko ? L"새해첫날" : L"New Year's Day");

            d = toSerial(year, 3, 1);
            if (year >= 2022 && isHol(d)) addSub(d, ko ? L"삼일절" : L"Independence Movement Day");
            add(d, ko ? L"삼일절" : L"Independence Movement Day");

            if (year < 2006) add(toSerial(year, 4, 5), ko ? L"식목일" : L"Arbor Day");

            if ((year >= 1994 || year < 1958) && mayDay)
                add(toSerial(year, 5, 1), ko ? L"노동절" : L"Labor Day");
            else if (year >= 1963 && mayDay)
                add(toSerial(year, 3, 10), ko ? L"근로자의 날" : L"Labor Day");

            d = toSerial(year, 5, 5);
            if (year > 2013 && isHol(d)) addSub(d, ko ? L"어린이날" : L"Children's Day");
            add(d, ko ? L"어린이날" : L"Children's Day");

            add(toSerial(year, 6, 6), ko ? L"현충일" : L"Memorial Day");

            if (year < 2008 || year >= 2026)
                add(toSerial(year, 7, 17), ko ? L"제헌절" : L"Constitution Day");

            d = toSerial(year, 8, 15);
            if (year >= 2022 && isHol(d)) addSub(d, ko ? L"광복절" : L"Liberation Day");
            add(d, ko ? L"광복절" : L"Liberation Day");

            d = toSerial(year, 10, 3);
            if (year >= 2022 && isHol(d)) addSub(d, ko ? L"개천절" : L"National Foundation Day");
            add(d, ko ? L"개천절" : L"National Foundation Day");

            if (year >= 2013)
            {
                d = toSerial(year, 10, 9);
                if (year >= 2022 && isHol(d)) addSub(d, ko ? L"한글날" : L"Hangul Day");
                add(d, ko ? L"한글날" : L"Hangul Day");
            }

            d = toSerial(year, 12, 25);
            if (year >= 2023 && isHol(d)) addSub(d, ko ? L"성탄절" : L"Christmas");
            add(d, ko ? L"성탄절" : L"Christmas");

            // 음력 공휴일: 설날(1/1)·석가탄신일(4/8)·추석(8/15).
            if (lunarNewYearSerial(year, 1, 1, d))
            {
                if (year > 2013 && (isHol(d) || d % 7 == 2 /*월*/))
                    addSub(d + 1, ko ? L"설날" : L"Lunar New Year");
                add(d - 1, ko ? L"설날연휴" : L"Lunar New Year Holiday");
                add(d, ko ? L"설날" : L"Lunar New Year");
                add(d + 1, ko ? L"설날연휴" : L"Lunar New Year Holiday");
            }
            if (lunarNewYearSerial(year, 4, 8, d))
            {
                if (year >= 2023 && isHol(d)) addSub(d, ko ? L"석가탄신일" : L"Buddha's Birthday");
                add(d, ko ? L"석가탄신일" : L"Buddha's Birthday");
            }
            if (lunarNewYearSerial(year, 8, 15, d))
            {
                if (year > 2013 && (isHol(d) || d % 7 == 2))
                    addSub(d + 1, ko ? L"추석" : L"Chuseok");
                add(d - 1, ko ? L"추석연휴" : L"Chuseok Holiday");
                add(d, ko ? L"추석" : L"Chuseok");
                add(d + 1, ko ? L"추석연휴" : L"Chuseok Holiday");
            }
        }

        ExcelObj* koreanHolidays(const ExcelObj& yearObj, const ExcelObj& nameObj,
                                 const ExcelObj& mayDayObj, const ExcelObj& keyObj)
        {
            int year;
            if (yearObj.isMissing())
            {
                SYSTEMTIME st;
                GetLocalTime(&st);
                year = st.wYear;
            }
            else
                year = yearObj.get<int>(0);
            if (year < kLunarFirstYear + 1 || year > kLunarLastYear)
                return returnValue(CellError::Value);
            const bool withName = nameObj.isMissing() ? true : (nameObj.get<double>(1.0) != 0.0);
            const bool mayDay = mayDayObj.isMissing() ? true : (mayDayObj.get<double>(1.0) != 0.0);
            const bool ko = egtools::i18n::current() == L"ko";

            // api_key 인수(4번째)의 3가지 용법 — 기본은 **API 사용**(사용자 지시
            // 2026-08-04):
            //   생략        → 저장된 data.go.kr 키로 병합(키 없거나 실패하면 내장만)
            //   "키 문자열" → 그 키를 저장하고 사용(설정 중이므로 실패는 오류로 보고)
            //   0 / FALSE   → API를 아예 쓰지 않고 내장 계산만(네트워크 접근 없음)
            const bool apiOff =
                !keyObj.isMissing() &&
                (keyObj.isType(ExcelType::Bool) || keyObj.isType(ExcelType::Num)) &&
                keyObj.get<double>(1.0) == 0.0;

            std::map<int, std::wstring> list;
            buildHolidayList(year, mayDay, ko, list);
            if (apiOff)
            {
                std::vector<ExcelObj> only;
                only.reserve(list.size() * (withName ? 2 : 1));
                for (const auto& [s, name] : list)
                {
                    only.push_back(ExcelObj((double)s));
                    if (withName) only.push_back(ExcelObj(std::wstring_view(name)));
                }
                return egtools::core::output(egtools::core::makeArray(
                    (ExcelArrayBuilder::row_t)list.size(), withName ? 2 : 1, only));
            }

            // 공공 API 병합(결정 3): 임시공휴일·선거일 등 내장에 없는 항목 추가.
            std::wstring apiKey;
            if (!keyObj.isMissing()) apiKey = keyObj.toString();

            // 실패 처리: 키를 **문자열로 명시**한 호출은 설정 중이므로 사유를
            // 셀에 돌려준다(FxPublicApi "ERROR: …" 관례 — 조용히 실패하면 원인을
            // 알 수 없다는 사용자 신고 2026-08-04). 생략한 호출은 수식이 깨지지
            // 않도록 내장 결과를 그대로 반환한다.
            std::wstring apiError;
            egtools::dates::mergeApiHolidays(year, list, apiKey, &apiError);
            if (!apiKey.empty() && !apiError.empty())
                return returnValue(ExcelObj(std::wstring_view(
                    L"ERROR: 공휴일 API - " + apiError)));

            std::vector<ExcelObj> vals;
            vals.reserve(list.size() * (withName ? 2 : 1));
            for (const auto& [s, name] : list)
            {
                vals.push_back(ExcelObj((double)s));
                if (withName) vals.push_back(ExcelObj(std::wstring_view(name)));
            }
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)list.size(), withName ? 2 : 1, vals));
        }

        // ---- NETWORKHOUR ---------------------------------------------------

        ExcelObj* netWorkHour(const ExcelObj& startObj, const ExcelObj& endObj,
                              const ExcelObj& breakObj)
        {
            if (startObj.isMissing() || endObj.isMissing())
                return returnValue(CellError::Value);

            auto dims = [](const ExcelObj& o, size_t& r, size_t& c)
            {
                if (o.isType(ExcelType::Multi))
                {
                    ExcelArray a(o);
                    r = a.nRows(); c = a.nCols();
                }
                else { r = 1; c = 1; }
            };
            size_t nR, nC, eR, eC;
            dims(startObj, nR, nC);
            dims(endObj, eR, eC);
            if (nR != eR || nC != eC) return returnValue(CellError::Value);

            auto cellAt = [](const ExcelObj& o, size_t r, size_t c, size_t cols) -> ExcelObj
            {
                if (o.isType(ExcelType::Multi))
                {
                    ExcelArray a(o);
                    return ExcelObj(a.at(r * cols + c));
                }
                return ExcelObj(o);
            };

            // 휴식 구간 정규화: 자정 넘김 분할 + 겹침 제거.
            std::vector<std::pair<double, double>> breaks;
            if (!breakObj.isMissing())
            {
                std::vector<std::pair<double, double>> raw;
                if (breakObj.isType(ExcelType::Multi))
                {
                    ExcelArray b(breakObj);
                    if (b.nCols() < 2) return returnValue(CellError::Value);
                    for (size_t r = 0; r < b.nRows(); ++r)
                    {
                        const ExcelObj& s = b.at(r * b.nCols());
                        const ExcelObj& e = b.at(r * b.nCols() + 1);
                        if (s.type() != ExcelType::Num && s.type() != ExcelType::Int) continue;
                        raw.emplace_back(s.get<double>(0.0), e.get<double>(0.0));
                    }
                }
                for (const auto& [bs, be] : raw)
                {
                    if (bs <= be) breaks.emplace_back(bs, be);
                    else { breaks.emplace_back(bs, 1.0); breaks.emplace_back(0.0, be); }
                }
                // 겹침 제거 (VB 로직 단순화: 구간 정렬 후 병합).
                std::sort(breaks.begin(), breaks.end());
                std::vector<std::pair<double, double>> merged;
                for (const auto& iv : breaks)
                {
                    if (!merged.empty() && iv.first <= merged.back().second)
                        merged.back().second = (std::max)(merged.back().second, iv.second);
                    else
                        merged.push_back(iv);
                }
                breaks.swap(merged);
            }

            std::vector<ExcelObj> vals;
            vals.reserve(nR * nC);
            for (size_t r = 0; r < nR; ++r)
                for (size_t c = 0; c < nC; ++c)
                {
                    const ExcelObj s = cellAt(startObj, r, c, nC);
                    const ExcelObj e = cellAt(endObj, r, c, nC);
                    const bool sNum = s.type() == ExcelType::Num || s.type() == ExcelType::Int;
                    const bool eNum = e.type() == ExcelType::Num || e.type() == ExcelType::Int;
                    if (!sNum || !eNum)
                    { vals.push_back(ExcelObj(std::wstring_view(L""))); continue; }

                    double ws = s.get<double>(0.0), we = e.get<double>(0.0);
                    if (we > 1 && ws > we)
                    { vals.push_back(ExcelObj(CellError::Value)); continue; }
                    if (ws > we) we += 1.0;   // 자정 넘김

                    double net = we - ws;
                    for (const auto& [bs, be] : breaks)
                        for (int day = (int)floor(ws); day <= (int)floor(we); ++day)
                        {
                            const double os = (std::max)(ws, day + bs);
                            const double oe = (std::min)(we, day + be);
                            if (oe > os) net -= (oe - os);
                        }
                    if (net < 0) net = 0;
                    vals.push_back(ExcelObj(net));
                }
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)nR, (ExcelArrayBuilder::col_t)nC, vals));
        }

        // ---- WEEKNUMOFMONTH / MONTHBHYWEEK ---------------------------------

        // 기준(소유) 요일이 속한 날짜의 직렬값 계산. 실패 시 INT_MIN.
        bool weekOwnerSerial(const ExcelObj& dateObj, const ExcelObj& sowObj,
                             const ExcelObj& oowObj, int& owner, int& sow, int& oow)
        {
            sow = sowObj.isMissing() ? 2 : sowObj.get<int>(2);
            oow = oowObj.isMissing() ? 5 : oowObj.get<int>(5);
            if (sow < 1 || sow > 7 || oow < 1 || oow > 7) return false;
            int serial;
            if (!argToSerial(dateObj, serial)) return false;

            owner = serial - (weekday1(serial) - oow);
            if (weekday1(serial) < sow) owner -= 7;
            if (sow > oow) owner += 7;
            return true;
        }

        ExcelObj* weekNumOfMonth(const ExcelObj& dateObj, const ExcelObj& sowObj,
                                 const ExcelObj& oowObj)
        {
            int owner, sow, oow;
            if (!weekOwnerSerial(dateObj, sowObj, oowObj, owner, sow, oow))
                return returnValue(CellError::Value);

            int y, m, d;
            fromSerial(owner, y, m, d);
            int m1 = toSerial(y, m, 1);
            m1 -= weekday1(m1) - oow;
            int my, mm, md;
            fromSerial(m1, my, mm, md);
            if (mm != m) m1 += 7;
            m1 -= weekday1(m1) - oow;

            return returnValue(ExcelObj((double)((owner - m1) / 7 + 1)));
        }

        ExcelObj* monthByWeek(const ExcelObj& dateObj, const ExcelObj& sowObj,
                              const ExcelObj& oowObj)
        {
            int owner, sow, oow;
            if (!weekOwnerSerial(dateObj, sowObj, oowObj, owner, sow, oow))
                return returnValue(CellError::Value);
            int y, m, d;
            fromSerial(owner, y, m, d);
            return returnValue(ExcelObj((double)m));
        }

        // ---- TODATETIME ----------------------------------------------------

        void replaceAll(std::wstring& s, const std::wstring& from, const std::wstring& to)
        {
            size_t p = 0;
            while ((p = s.find(from, p)) != std::wstring::npos)
            {
                s.replace(p, from.size(), to);
                p += to.size();
            }
        }

        ExcelObj* toDateTime(const ExcelObj& textObj)
        {
            if (textObj.isMissing()) return returnValue(CellError::Value);
            const std::wstring orig = textObj.toString();
            std::wstring s = orig;
            for (auto& c : s) c = (wchar_t)std::towupper(c);
            while (s.find(L"  ") != std::wstring::npos) replaceAll(s, L"  ", L" ");
            replaceAll(s, L"/", L"-"); replaceAll(s, L".", L"-");
            replaceAll(s, L"년 ", L"년"); replaceAll(s, L"年 ", L"년"); replaceAll(s, L"年", L"년");
            replaceAll(s, L"년", L"-");
            replaceAll(s, L"월 ", L"월"); replaceAll(s, L"月 ", L"월"); replaceAll(s, L"月", L"월");
            replaceAll(s, L"월", L"-");
            replaceAll(s, L"일", L""); replaceAll(s, L"日", L"");
            replaceAll(s, L"時", L"시"); replaceAll(s, L"分", L"분"); replaceAll(s, L"秒", L"초");
            replaceAll(s, L"시 ", L"시"); replaceAll(s, L"분 ", L"분");
            if (s.find(L"초") != std::wstring::npos)
            { replaceAll(s, L"시", L":"); replaceAll(s, L"분", L":"); replaceAll(s, L"초", L""); }
            else if (s.find(L"분") != std::wstring::npos)
            { replaceAll(s, L"시", L":"); replaceAll(s, L"분", L""); }
            else
                replaceAll(s, L"시", L":00");

            double dateSerial = -1;
            try
            {
                std::wsmatch m;
                if (std::regex_search(s, m, std::wregex(L"(\\d{1,4})-(\\d{1,2})-(\\d{1,4})")))
                {
                    long a = _wtol(m[1].str().c_str());
                    long b = _wtol(m[2].str().c_str());
                    long c = _wtol(m[3].str().c_str());
                    int y = 0, mo = 0, dd = 0;
                    // 연도 위치 판정: 4자리 우선, 다음 31 초과 2자리(2000년대 보정).
                    // (VB 원본은 "yyyy-mm-dd"의 일(dd)을 2자리 연도로 오인하는 버그가
                    //  있어 재설계.)
                    if (a > 31 && c > 31) { /* 날짜 아님 */ }
                    else
                    {
                        if (a >= 1000) { y = (int)a; if (b > 12) { mo = (int)c; dd = (int)b; } else { mo = (int)b; dd = (int)c; } }
                        else if (c >= 1000) { y = (int)c; if (a > 12) { mo = (int)b; dd = (int)a; } else { mo = (int)a; dd = (int)b; } }
                        else if (a > 31) { y = 2000 + (int)a; if (b > 12) { mo = (int)c; dd = (int)b; } else { mo = (int)b; dd = (int)c; } }
                        else if (c > 31) { y = 2000 + (int)c; if (a > 12) { mo = (int)b; dd = (int)a; } else { mo = (int)a; dd = (int)b; } }
                        if (y >= 1000 && mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31)
                            dateSerial = toSerial(y, mo, dd);
                    }
                }
                else if (std::regex_search(s, m, std::wregex(L"(\\d{1,2})-(\\d{1,2})")))
                {
                    SYSTEMTIME st;
                    GetLocalTime(&st);
                    long a = _wtol(m[1].str().c_str());
                    long b = _wtol(m[2].str().c_str());
                    const int mo = a > 12 ? (int)b : (int)a;
                    const int dd = a > 12 ? (int)a : (int)b;
                    if (mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31)
                        dateSerial = toSerial(st.wYear, mo, dd);
                }

                double timePart = 0;
                std::wsmatch tm;
                if (std::regex_search(s, tm, std::wregex(L"(\\d{1,2}):(\\d{1,2})(:(\\d{1,2}))?")))
                {
                    const int h = _wtoi(tm[1].str().c_str());
                    const int mi = _wtoi(tm[2].str().c_str());
                    const int se = tm[4].matched ? _wtoi(tm[4].str().c_str()) : 0;
                    if (h < 24 && mi < 60 && se < 60)
                        timePart = (h * 3600 + mi * 60 + se) / 86400.0;
                }

                // 오전/오후 처리.
                const bool pm = s.find(L"PM") != std::wstring::npos ||
                                orig.find(L"오후") != std::wstring::npos ||
                                orig.find(L"午後") != std::wstring::npos ||
                                orig.find(L"下午") != std::wstring::npos;
                const bool am = s.find(L"AM") != std::wstring::npos ||
                                orig.find(L"오전") != std::wstring::npos ||
                                orig.find(L"午前") != std::wstring::npos ||
                                orig.find(L"上午") != std::wstring::npos;
                if (pm && timePart < 0.5) timePart += 0.5;
                else if (am && timePart >= 0.5 && timePart < 13.0 / 24) timePart -= 0.5;

                if (dateSerial >= 0 && timePart > 0)
                    return returnValue(ExcelObj(dateSerial + timePart));
                if (dateSerial >= 0) return returnValue(ExcelObj(dateSerial));
                if (timePart > 0) return returnValue(ExcelObj(timePart));
                return returnValue(ExcelObj(std::wstring_view(L"")));
            }
            catch (...) { return returnValue(CellError::Value); }
        }
    }

    void registerEgDate()
    {
        egtools::core::registerFn(L"TOLUNAR",
            [](const ExcelObj& v) -> ExcelObj*
            { return egtools::core::mapUnary(v, toLunarOne); });

        egtools::core::registerFn(L"TOSOLAR",
            [](const ExcelObj& v) -> ExcelObj*
            { return egtools::core::mapUnary(v, toSolarOne); });

        egtools::core::registerFn(L"KOREANHOLIDAYS",
            [](const ExcelObj& y, const ExcelObj& n, const ExcelObj& md,
               const ExcelObj& key) -> ExcelObj*
            {
                try { return koreanHolidays(y, n, md, key); }
                catch (...) { return returnValue(CellError::Value); }
            });

        egtools::core::registerFn(L"NETWORKHOUR",
            [](const ExcelObj& s, const ExcelObj& e, const ExcelObj& b) -> ExcelObj*
            { return netWorkHour(s, e, b); });

        egtools::core::registerFn(L"WEEKNUMOFMONTH",
            [](const ExcelObj& d, const ExcelObj& sow, const ExcelObj& oow) -> ExcelObj*
            { return weekNumOfMonth(d, sow, oow); });

        egtools::core::registerFn(L"MONTHBHYWEEK",
            [](const ExcelObj& d, const ExcelObj& sow, const ExcelObj& oow) -> ExcelObj*
            { return monthByWeek(d, sow, oow); });

        egtools::core::registerFn(L"TODATETIME",
            [](const ExcelObj& t) -> ExcelObj*
            {
                try { return toDateTime(t); }
                catch (...) { return returnValue(CellError::Value); }
            });
    }
}

// ── 공용 공휴일/음력 헬퍼 (KoreanHolidays.h) ────────────────────────────────
// 익명 네임스페이스 헬퍼(buildHolidayList/toSerial/solarToLunar)는 같은 TU라
// egtools::functions:: 한정 이름으로 접근 가능.
namespace egtools::dates
{
    void builtinHolidays(int year, bool mayDay, bool korean,
                         std::map<int, std::wstring>& list)
    {
        egtools::functions::buildHolidayList(year, mayDay, korean, list);
    }

    bool lunarOf(int serial, int& lYear, int& lMonth, int& lDay)
    {
        return egtools::functions::solarToLunar(serial, lYear, lMonth, lDay);
    }

    // ── data.go.kr 인증키(단일 슬롯 "datago") ───────────────────────────────
    // 공공데이터포털은 계정당 키 하나를 발급하고 활용신청한 모든 서비스에 그
    // 키를 쓴다. 따라서 서비스별 슬롯은 불필요 — KOREANHOLIDAYS(특일정보)와
    // BRNSTATUS(사업자등록)가 이 값을 공유한다. 구 "odcloud" 슬롯은 최초 조회
    // 시 datago로 옮기고 삭제한다(사용자 지시 2026-08-04, 통일).
    namespace
    {
        constexpr const wchar_t* kApiKeyPath = L"Software\\EGTools\\ApiKeys";

        std::wstring regReadKey(const wchar_t* value)
        {
            HKEY h = nullptr;
            if (RegOpenKeyExW(HKEY_CURRENT_USER, kApiKeyPath, 0, KEY_READ, &h)
                    != ERROR_SUCCESS)
                return L"";
            wchar_t buf[512]{};
            DWORD size = sizeof(buf) - sizeof(wchar_t), type = 0;
            std::wstring out;
            if (RegQueryValueExW(h, value, nullptr, &type, (BYTE*)buf, &size)
                    == ERROR_SUCCESS && type == REG_SZ)
                out = buf;
            RegCloseKey(h);
            return out;
        }
    }

    void saveDataGoKey(const std::wstring& key)
    {
        if (key.empty()) return;
        HKEY h = nullptr;
        if (RegCreateKeyExW(HKEY_CURRENT_USER, kApiKeyPath, 0, nullptr, 0,
                            KEY_WRITE, nullptr, &h, nullptr) != ERROR_SUCCESS)
            return;
        RegSetValueExW(h, L"datago", 0, REG_SZ, (const BYTE*)key.c_str(),
                       (DWORD)((key.size() + 1) * sizeof(wchar_t)));
        RegDeleteValueW(h, L"odcloud");     // 통일 후 구 슬롯 제거
        RegCloseKey(h);
    }

    std::wstring dataGoKey()
    {
        std::wstring key = regReadKey(L"datago");
        if (!key.empty()) return key;
        key = regReadKey(L"odcloud");       // 구 버전 등록분 → 자동 이관
        if (!key.empty()) saveDataGoKey(key);
        return key;
    }

    int mergeApiHolidays(int year, std::map<int, std::wstring>& list,
                         const std::wstring& apiKeyArg, std::wstring* error)
    {
        auto fail = [&](std::wstring msg) { if (error) *error = std::move(msg); return 0; };

        // ── 키 결정: 인수 우선(전달 시 저장 — 공공 API 키 정책), 없으면 레지스트리 ──
        // data.go.kr은 계정당 인증키 하나로 모든 서비스를 쓴다. 종전에는 서비스별로
        // datago(공휴일)/odcloud(사업자등록) 슬롯을 따로 뒀는데 같은 값이라 혼란만
        // 컸다 — **datago 하나로 통일**하고 구 odcloud 값은 자동 이관한다
        // (사용자 지시 2026-08-04: 키를 odcloud로만 등록해 임시공휴일이 병합되지
        // 않던 문제의 근본 해소).
        std::wstring key = apiKeyArg;
        if (!key.empty()) egtools::dates::saveDataGoKey(key);
        else key = egtools::dates::dataGoKey();
        if (key.empty())
            return fail(L"data.go.kr 키가 없습니다. 마지막 인수로 1회 입력하거나 "
                        L"[EGTools] > [API 키 관리]에서 등록하세요.");

        // 포털은 인증키를 Encoding(퍼센트 인코딩)/Decoding 두 형태로 보여준다.
        // Decoding 키를 그대로 붙이면 '+/=' 때문에 서명이 깨지므로 인코딩한다
        // (이미 인코딩된 키는 '%'가 있으므로 건드리지 않는다).
        if (key.find(L'%') == std::wstring::npos)
        {
            std::wstring enc;
            for (wchar_t c : key)
            {
                if (c == L'+') enc += L"%2B";
                else if (c == L'/') enc += L"%2F";
                else if (c == L'=') enc += L"%3D";
                else enc += c;
            }
            key.swap(enc);
        }

        // ── 연도별 세션 캐시: **성공만** 캐시한다. 실패는 캐시하지 않고 매번
        //    다시 조회한다 — 사용자가 키를 고치면 즉시 반영되어야 하기 때문
        //    (사용자 지시 2026-08-04: 재시도 지연을 두면 조치 결과를 바로
        //    확인할 수 없다). 실패가 이어지면 매 재계산마다 타임아웃(최대 약
        //    10초)이 발생하므로 타임아웃 값은 짧게 유지한다.
        static std::map<int, std::vector<std::pair<int, std::wstring>>> cache;
        auto it = cache.find(year);
        if (it == cache.end())
        {
            std::vector<std::pair<int, std::wstring>> fetched;
            std::string body;
            HINTERNET hSession = WinHttpOpen(L"EGTools++",
                WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);
            if (hSession)
            {
                WinHttpSetTimeouts(hSession, 5000, 5000, 10000, 10000);
                HINTERNET hConnect = WinHttpConnect(hSession, L"apis.data.go.kr",
                                                    INTERNET_DEFAULT_HTTPS_PORT, 0);
                if (hConnect)
                {
                    const std::wstring path =
                        L"/B090041/openapi/service/SpcdeInfoService/getRestDeInfo"
                        L"?serviceKey=" + key +
                        L"&solYear=" + std::to_wstring(year) + L"&numOfRows=100";
                    HINTERNET hReq = WinHttpOpenRequest(hConnect, L"GET", path.c_str(),
                        nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES,
                        WINHTTP_FLAG_SECURE);
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
                                if (!WinHttpQueryDataAvailable(hReq, &avail) || !avail)
                                    break;
                                std::string chunk((size_t)avail, '\0');
                                DWORD read = 0;
                                if (!WinHttpReadData(hReq, chunk.data(), avail, &read))
                                    break;
                                body.append(chunk.data(), read);
                            } while (avail > 0);
                        }
                        WinHttpCloseHandle(hReq);
                    }
                    WinHttpCloseHandle(hConnect);
                }
                WinHttpCloseHandle(hSession);
            }

            // <item> 단위로 locdate(YYYYMMDD)/dateName/isHoliday(Y) 추출(UTF-8).
            auto tagText = [](const std::string& s, size_t from, size_t to,
                              const char* tag) -> std::string
            {
                const std::string open = std::string("<") + tag + ">";
                const std::string close = std::string("</") + tag + ">";
                const size_t a = s.find(open, from);
                if (a == std::string::npos || a >= to) return "";
                const size_t b = s.find(close, a);
                if (b == std::string::npos || b > to) return "";
                return s.substr(a + open.size(), b - a - open.size());
            };
            size_t pos = 0;
            while (true)
            {
                const size_t a = body.find("<item>", pos);
                if (a == std::string::npos) break;
                const size_t b = body.find("</item>", a);
                if (b == std::string::npos) break;
                pos = b + 7;

                if (tagText(body, a, b, "isHoliday") != "Y") continue;
                const std::string loc = tagText(body, a, b, "locdate");
                if (loc.size() != 8) continue;
                const int y = atoi(loc.substr(0, 4).c_str());
                const int m = atoi(loc.substr(4, 2).c_str());
                const int d = atoi(loc.substr(6, 2).c_str());
                if (y != year || m < 1 || m > 12 || d < 1 || d > 31) continue;

                const std::string name8 = tagText(body, a, b, "dateName");
                std::wstring name;
                if (!name8.empty())
                {
                    const int n = MultiByteToWideChar(CP_UTF8, 0, name8.c_str(),
                                                      (int)name8.size(), nullptr, 0);
                    name.resize((size_t)n);
                    MultiByteToWideChar(CP_UTF8, 0, name8.c_str(), (int)name8.size(),
                                        name.data(), n);
                }
                fetched.emplace_back(egtools::functions::toSerial(y, m, d), name);
            }

            // 실패 판정: 응답 없음 / resultCode≠00(키 미등록·트래픽 초과 등) /
            //           항목 0건. 사유를 그대로 돌려줘 사용자가 조치할 수 있게 한다.
            if (body.empty())
                return fail(L"서버에 연결하지 못했습니다(네트워크·방화벽 확인).");
            const std::string code = tagText(body, 0, body.size(), "resultCode");
            if (!code.empty() && code != "00")
            {
                std::string msg = tagText(body, 0, body.size(), "resultMsg");
                if (msg.empty()) msg = tagText(body, 0, body.size(), "returnAuthMsg");
                std::wstring wmsg(msg.begin(), msg.end());   // 코드/메시지는 ASCII
                return fail(L"키가 거부되었습니다(" + wmsg + L"). 활용신청 승인 여부와 "
                            L"키 값을 확인하세요.");
            }
            if (fetched.empty())
                return fail(L"응답에 공휴일 항목이 없습니다(특일정보 활용신청 확인).");
            it = cache.emplace(year, std::move(fetched)).first;
        }

        // 내장 계산에 없는 날짜만 추가(임시공휴일·선거일 등).
        int added = 0;
        for (const auto& [serial, name] : it->second)
            if (!list.count(serial))
            {
                list[serial] = name.empty() ? L"임시공휴일" : name;
                ++added;
            }
        return added;
    }
}
