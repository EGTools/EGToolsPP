// FxMath.cpp — math / number-format / finance helpers added for older-Excel
// drop-in (mostly Excel 2013) plus PERCENTOF (365/2024).
//
// All scalar (single value in → single value out); invalid args return the
// standard Excel error rather than crashing.

#include "../core/Registry.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <cmath>
#include <string>
#include <algorithm>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // Read a required numeric arg; ok=false if missing/non-numeric.
        double num(const ExcelObj& o, bool& ok)
        {
            auto t = o.type();
            if (t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool)
            { ok = true; return o.get<double>(0.0); }
            // numeric text is accepted by Excel for most of these
            if (t == ExcelType::Str)
            {
                try { size_t p = 0; const std::wstring s = o.toString();
                      double v = std::stod(s, &p); if (p) { ok = true; return v; } }
                catch (...) {}
            }
            ok = false; return 0.0;
        }

        double optNum(const ExcelObj& o, double dflt)
        {
            if (o.isMissing()) return dflt;
            bool ok; double v = num(o, ok); return ok ? v : dflt;
        }

        // Sum all numeric cells in a value/array (for PERCENTOF).
        double sumNumeric(const ExcelObj& o, bool& any)
        {
            double s = 0; any = false;
            if (o.type() == ExcelType::Multi)
            {
                ExcelArray a(o);
                for (ExcelArray::row_t r = 0; r < a.nRows(); ++r)
                    for (ExcelArray::col_t c = 0; c < a.nCols(); ++c)
                    {
                        const ExcelObj& e = a.at(r, c);
                        auto t = e.type();
                        if (t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool)
                        { s += e.get<double>(0.0); any = true; }
                    }
            }
            else
            {
                bool ok; double v = num(o, ok); if (ok) { s = v; any = true; }
            }
            return s;
        }

        // C(n, k) with doubles, guarding overflow reasonably.
        double nCk(double n, double k)
        {
            if (k < 0 || n < 0 || k > n) return 0.0;
            k = std::min(k, n - k);
            double r = 1.0;
            for (double i = 0; i < k; ++i) r = r * (n - i) / (i + 1.0);
            return std::round(r);
        }
    }

    void registerMath()
    {
        // PERCENTOF(data_subset, data_all) = SUM(subset) / SUM(all).
        egtools::core::registerFn(L"PERCENTOF",
            [](const ExcelObj& subset, const ExcelObj& all) -> ExcelObj*
            {
                if (subset.isMissing() || all.isMissing()) return returnValue(CellError::Value);
                bool a1, a2; double s = sumNumeric(subset, a1), t = sumNumeric(all, a2);
                if (!a2 || t == 0.0) return returnValue(CellError::Div0);
                return returnValue(ExcelObj(s / t));
            });

        // COMBINA(number, number_chosen) — combinations with repetition.
        egtools::core::registerFn(L"COMBINA",
            [](const ExcelObj& nA, const ExcelObj& kA) -> ExcelObj*
            {
                bool o1, o2; double n = std::floor(num(nA, o1)), k = std::floor(num(kA, o2));
                if (!o1 || !o2 || n < 0 || k < 0) return returnValue(CellError::Num);
                if (n == 0 && k == 0) return returnValue(ExcelObj(1.0));
                return returnValue(ExcelObj(nCk(n + k - 1, k)));
            });

        // PERMUTATIONA(number, number_chosen) = number ^ number_chosen.
        egtools::core::registerFn(L"PERMUTATIONA",
            [](const ExcelObj& nA, const ExcelObj& kA) -> ExcelObj*
            {
                bool o1, o2; double n = std::floor(num(nA, o1)), k = std::floor(num(kA, o2));
                if (!o1 || !o2 || n < 0 || k < 0) return returnValue(CellError::Num);
                return returnValue(ExcelObj(std::pow(n, k)));
            });

        // CEILING.MATH(number, [significance], [mode]) — round up to multiple.
        egtools::core::registerFn(L"CEILING.MATH",
            [](const ExcelObj& numA, const ExcelObj& sigA, const ExcelObj& modeA) -> ExcelObj*
            {
                bool ok; double x = num(numA, ok); if (!ok) return returnValue(CellError::Value);
                double sig = optNum(sigA, 1.0); double mode = optNum(modeA, 0.0);
                if (sig == 0.0) return returnValue(ExcelObj(0.0));
                sig = std::fabs(sig);
                double base = x / sig;
                double m = (x >= 0 || mode == 0) ? std::ceil(base) : std::floor(base);
                return returnValue(ExcelObj(m * sig));
            });

        // FLOOR.MATH(number, [significance], [mode]) — round down to multiple.
        egtools::core::registerFn(L"FLOOR.MATH",
            [](const ExcelObj& numA, const ExcelObj& sigA, const ExcelObj& modeA) -> ExcelObj*
            {
                bool ok; double x = num(numA, ok); if (!ok) return returnValue(CellError::Value);
                double sig = optNum(sigA, 1.0); double mode = optNum(modeA, 0.0);
                if (sig == 0.0) return returnValue(ExcelObj(0.0));
                sig = std::fabs(sig);
                double base = x / sig;
                double m = (x >= 0 || mode == 0) ? std::floor(base) : std::ceil(base);
                return returnValue(ExcelObj(m * sig));
            });

        // BASE(number, radix, [min_length]) — integer → text in given radix.
        egtools::core::registerFn(L"BASE",
            [](const ExcelObj& numA, const ExcelObj& radixA, const ExcelObj& minLenA) -> ExcelObj*
            {
                bool o1, o2; double nd = std::floor(num(numA, o1)); int radix = (int)std::floor(num(radixA, o2));
                if (!o1 || !o2 || nd < 0 || radix < 2 || radix > 36) return returnValue(CellError::Num);
                long long n = (long long)nd;
                std::wstring out;
                if (n == 0) out = L"0";
                while (n > 0) { int d = (int)(n % radix); out.insert(out.begin(), (wchar_t)(d < 10 ? L'0' + d : L'A' + d - 10)); n /= radix; }
                int minLen = (int)optNum(minLenA, 0.0);
                while ((int)out.size() < minLen) out.insert(out.begin(), L'0');
                return returnValue(ExcelObj(out));
            });

        // DECIMAL(text, radix) — text in given radix → number.
        egtools::core::registerFn(L"DECIMAL",
            [](const ExcelObj& textA, const ExcelObj& radixA) -> ExcelObj*
            {
                bool o2; int radix = (int)std::floor(num(radixA, o2));
                if (!o2 || radix < 2 || radix > 36) return returnValue(CellError::Num);
                std::wstring s = textA.toString();
                double acc = 0;
                for (wchar_t c : s)
                {
                    int d;
                    if (c >= L'0' && c <= L'9') d = c - L'0';
                    else if (c >= L'A' && c <= L'Z') d = c - L'A' + 10;
                    else if (c >= L'a' && c <= L'z') d = c - L'a' + 10;
                    else return returnValue(CellError::Num);
                    if (d >= radix) return returnValue(CellError::Num);
                    acc = acc * radix + d;
                }
                return returnValue(ExcelObj(acc));
            });

        // ARABIC(roman_text) — Roman numeral → number.
        egtools::core::registerFn(L"ARABIC",
            [](const ExcelObj& textA) -> ExcelObj*
            {
                std::wstring s = textA.toString();
                std::wstring t; for (wchar_t c : s) if (!iswspace(c)) t += (wchar_t)towupper(c);
                bool neg = false; size_t i = 0;
                if (!t.empty() && t[0] == L'-') { neg = true; i = 1; }
                auto val = [](wchar_t c) -> int {
                    switch (c) { case L'I': return 1; case L'V': return 5; case L'X': return 10;
                                 case L'L': return 50; case L'C': return 100; case L'D': return 500;
                                 case L'M': return 1000; default: return -1; } };
                long long total = 0; int prev = 0;
                for (size_t j = t.size(); j-- > i; )
                {
                    int v = val(t[j]); if (v < 0) return returnValue(CellError::Value);
                    if (v < prev) total -= v; else { total += v; prev = v; }
                }
                return returnValue(ExcelObj((double)(neg ? -total : total)));
            });

        // RRI(nper, pv, fv) = (fv/pv)^(1/nper) − 1.
        egtools::core::registerFn(L"RRI",
            [](const ExcelObj& nperA, const ExcelObj& pvA, const ExcelObj& fvA) -> ExcelObj*
            {
                bool o1, o2, o3; double nper = num(nperA, o1), pv = num(pvA, o2), fv = num(fvA, o3);
                if (!o1 || !o2 || !o3 || nper <= 0 || pv == 0) return returnValue(CellError::Num);
                return returnValue(ExcelObj(std::pow(fv / pv, 1.0 / nper) - 1.0));
            });

        // PDURATION(rate, pv, fv) = (ln(fv) − ln(pv)) / ln(1+rate).
        egtools::core::registerFn(L"PDURATION",
            [](const ExcelObj& rateA, const ExcelObj& pvA, const ExcelObj& fvA) -> ExcelObj*
            {
                bool o1, o2, o3; double rate = num(rateA, o1), pv = num(pvA, o2), fv = num(fvA, o3);
                if (!o1 || !o2 || !o3 || rate <= 0 || pv <= 0 || fv <= 0) return returnValue(CellError::Num);
                return returnValue(ExcelObj((std::log(fv) - std::log(pv)) / std::log(1.0 + rate)));
            });
    }
}
