// AggCore.h — pure numeric aggregation core (no xlOil/Excel dependency) shared
// by GROUPBY/PIVOTBY/VISIBLEAGGR via core/Aggregate.h and unit-tested by
// tools/aggregate_selftest.cpp. Semantics follow the native Excel functions.
#pragma once
#include <algorithm>
#include <cmath>
#include <map>
#include <optional>
#include <string>
#include <vector>

namespace egtools::agg
{
    enum class Err { None, Value, Num, Div0, NA };
    struct NumResult { Err err; double value; };

    inline NumResult ok(double v) { return { Err::None, v }; }
    inline NumResult fail(Err e) { return { e, 0.0 }; }

    // PERCENTILE.INC: linear interpolation at rank p*(n-1) over ascending v.
    inline NumResult percentileInc(std::vector<double> v, double p)
    {
        if (v.empty() || p < 0.0 || p > 1.0) return fail(Err::Num);
        std::sort(v.begin(), v.end());
        const double rank = p * (double)(v.size() - 1);
        const size_t lo = (size_t)rank;
        const double frac = rank - (double)lo;
        if (lo + 1 >= v.size()) return ok(v.back());
        return ok(v[lo] + frac * (v[lo + 1] - v[lo]));
    }

    // PERCENTILE.EXC: rank p*(n+1)-1; p must lie in [1/(n+1), n/(n+1)].
    inline NumResult percentileExc(std::vector<double> v, double p)
    {
        const size_t n = v.size();
        if (n == 0) return fail(Err::Num);
        if (p < 1.0 / (double)(n + 1) || p > (double)n / (double)(n + 1)) return fail(Err::Num);
        std::sort(v.begin(), v.end());
        const double rank = p * (double)(n + 1) - 1.0;
        const size_t lo = (size_t)rank;
        const double frac = rank - (double)lo;
        if (lo + 1 >= n) return ok(v.back());
        return ok(v[lo] + frac * (v[lo + 1] - v[lo]));
    }

    // fn must be UPPERCASE. nums in original order (MODE.SNGL first-occurrence
    // tie-break). param: LARGE/SMALL k, PERCENTILE* p, QUARTILE* quart.
    // Returns Err::Value for names this numeric core does not know.
    inline NumResult aggNumeric(const std::wstring& f, const std::vector<double>& nums,
                                std::optional<double> param)
    {
        const size_t n = nums.size();

        if (f == L"SUM")
        {
            double s = 0; for (double d : nums) s += d;
            return ok(s);
        }
        if (f == L"AVERAGE")
        {
            if (n == 0) return fail(Err::Div0);
            double s = 0; for (double d : nums) s += d;
            return ok(s / (double)n);
        }
        if (f == L"MIN" || f == L"MAX")
        {
            if (n == 0) return ok(0.0);
            double best = nums[0];
            for (double d : nums) best = (f == L"MAX") ? std::max(best, d) : std::min(best, d);
            return ok(best);
        }
        if (f == L"PRODUCT")
        {
            if (n == 0) return ok(0.0);
            double p = 1; for (double d : nums) p *= d;
            return ok(p);
        }
        if (f == L"MEDIAN")
            return percentileInc(nums, 0.5);
        if (f == L"MODE.SNGL")
        {
            if (n == 0) return fail(Err::NA);
            // first-occurring value with the highest count >= 2
            std::map<double, size_t> firstIdx;
            std::map<double, int> count;
            for (size_t i = 0; i < n; ++i)
            {
                if (!count.count(nums[i])) firstIdx[nums[i]] = i;
                ++count[nums[i]];
            }
            int bestCnt = 1; double bestVal = 0; size_t bestFirst = n;
            for (auto& [val, cnt] : count)
                if (cnt > bestCnt || (cnt == bestCnt && cnt > 1 && firstIdx[val] < bestFirst))
                    { bestCnt = cnt; bestVal = val; bestFirst = firstIdx[val]; }
            if (bestCnt < 2) return fail(Err::NA);
            return ok(bestVal);
        }
        if (f == L"STDEV.S" || f == L"STDEV.P" || f == L"VAR.S" || f == L"VAR.P")
        {
            const bool sample = (f == L"STDEV.S" || f == L"VAR.S");
            if (n < (sample ? 2u : 1u)) return fail(Err::Div0);
            double mean = 0; for (double d : nums) mean += d;
            mean /= (double)n;
            double ss = 0; for (double d : nums) ss += (d - mean) * (d - mean);
            const double var = ss / (double)(sample ? n - 1 : n);
            return ok((f[0] == L'S') ? std::sqrt(var) : var);
        }
        if (f == L"LARGE" || f == L"SMALL")
        {
            if (!param || n == 0) return fail(Err::Num);
            const long k = (long)*param;
            if (k < 1 || (size_t)k > n) return fail(Err::Num);
            std::vector<double> v = nums;
            std::sort(v.begin(), v.end());
            return ok(f == L"SMALL" ? v[(size_t)k - 1] : v[n - (size_t)k]);
        }
        if (f == L"PERCENTILE" || f == L"PERCENTILE.INC")
        {
            if (!param) return fail(Err::Num);
            return percentileInc(nums, *param);
        }
        if (f == L"PERCENTILE.EXC")
        {
            if (!param) return fail(Err::Num);
            return percentileExc(nums, *param);
        }
        if (f == L"QUARTILE" || f == L"QUARTILE.INC")
        {
            if (!param) return fail(Err::Num);
            const long q = (long)*param;          // Excel truncates
            if (q < 0 || q > 4) return fail(Err::Num);
            return percentileInc(nums, (double)q / 4.0);
        }
        if (f == L"QUARTILE.EXC")
        {
            if (!param) return fail(Err::Num);
            const long q = (long)*param;
            if (q < 1 || q > 3) return fail(Err::Num);
            return percentileExc(nums, (double)q / 4.0);
        }
        return fail(Err::Value);
    }

    // Excel "General"-style number → text (for CONCAT/TEXTJOIN/ARRAYTOTEXT).
    inline std::wstring numToText(double d)
    {
        if (d == (long long)d && std::abs(d) < 1e15)
            return std::to_wstring((long long)d);
        wchar_t buf[40];
        swprintf_s(buf, L"%.15g", d);
        return buf;
    }
}
