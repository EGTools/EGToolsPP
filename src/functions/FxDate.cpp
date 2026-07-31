// FxDate.cpp — date/time functions.

#include "../core/Registry.h"
#include "../core/Apply.h"

#include <xlOil/xlOil.h>
#include <xlOil/Date.h>
#include <cmath>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        bool isLeap(int y) { return (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0); }

        // ISO weekday 1=Mon..7=Sun (Sakamoto).
        int isoWeekday(int y, int m, int d)
        {
            static const int t[] = { 0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4 };
            int yy = (m < 3) ? y - 1 : y;
            int dow = (yy + yy / 4 - yy / 100 + yy / 400 + t[m - 1] + d) % 7; // 0=Sun
            return ((dow + 6) % 7) + 1;
        }

        int weeksInYear(int y)
        {
            auto p = [](int yr) { return (yr + yr / 4 - yr / 100 + yr / 400) % 7; };
            return (p(y) == 4 || p(y - 1) == 3) ? 53 : 52;
        }

        ExcelObj isoWeekNumOne(const ExcelObj& e)
        {
            const int serial = e.get<int>(-1);
            if (serial < 0) return ExcelObj(CellError::Value);
            int y = 0, m = 0, d = 0;
            if (!excelSerialDateToYMD(serial, y, m, d))
                return ExcelObj(CellError::Num);

            static const int cum[] = { 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334 };
            int ordinal = cum[m - 1] + d + ((m > 2 && isLeap(y)) ? 1 : 0);
            int wd = isoWeekday(y, m, d);
            int week = (ordinal - wd + 10) / 7;
            if (week < 1)        week = weeksInYear(y - 1);
            else if (week > weeksInYear(y)) week = 1;
            return ExcelObj((double)week);
        }
    }

    void registerDate()
    {
        // DAYS(end_date, start_date): whole-day difference.
        egtools::core::registerFn(L"DAYS",
            [](const ExcelObj& endDate, const ExcelObj& startDate) -> ExcelObj*
            {
                if (endDate.isMissing() || startDate.isMissing())
                    return returnValue(CellError::Value);
                const double e = std::floor(endDate.get<double>(0.0));
                const double s = std::floor(startDate.get<double>(0.0));
                return returnValue(ExcelObj(e - s));
            });

        // ISOWEEKNUM(date): ISO 8601 week number (element-wise).
        egtools::core::registerFn(L"ISOWEEKNUM",
            [](const ExcelObj& date) -> ExcelObj* { return egtools::core::mapUnary(date, isoWeekNumOne); });
    }
}
