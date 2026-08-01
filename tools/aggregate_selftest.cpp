// aggregate_selftest.cpp — prints AggCore.h results for a fixed dataset so
// tests can diff them against native Excel formulas. Never shipped.
#include "core/AggCore.h"

#include <cstdio>
#include <vector>

using namespace egtools::agg;

int wmain()
{
    const std::vector<double> v = { 3, 7, 1, 9, 4, 7 };

    struct Case { const wchar_t* name; const wchar_t* fn; double param; bool hasParam; };
    const Case cases[] = {
        { L"MEDIAN",          L"MEDIAN",          0,   false },
        { L"MODE.SNGL",       L"MODE.SNGL",       0,   false },
        { L"STDEV.S",         L"STDEV.S",         0,   false },
        { L"STDEV.P",         L"STDEV.P",         0,   false },
        { L"VAR.S",           L"VAR.S",           0,   false },
        { L"VAR.P",           L"VAR.P",           0,   false },
        { L"LARGE2",          L"LARGE",           2,   true },
        { L"SMALL2",          L"SMALL",           2,   true },
        { L"PERCENTILE.INC3", L"PERCENTILE.INC",  0.3, true },
        { L"PERCENTILE.EXC3", L"PERCENTILE.EXC",  0.3, true },
        { L"QUARTILE.INC1",   L"QUARTILE.INC",    1,   true },
        { L"QUARTILE.EXC3",   L"QUARTILE.EXC",    3,   true },
    };

    for (const auto& c : cases)
    {
        const auto r = aggNumeric(c.fn, v,
            c.hasParam ? std::optional<double>(c.param) : std::nullopt);
        if (r.err == Err::None) wprintf(L"%s=%.12g\n", c.name, r.value);
        else wprintf(L"%s=ERR%d\n", c.name, (int)r.err);
    }
    return 0;
}
