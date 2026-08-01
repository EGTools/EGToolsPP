#include "Version.h"

#include <windows.h>
#include <algorithm>
#include <map>
#include <cwctype>

#include <xlOil/State.h>

namespace egtools::core
{
    namespace
    {
        std::wstring upper(std::wstring s)
        {
            std::transform(s.begin(), s.end(), s.begin(),
                           [](wchar_t c) { return (wchar_t)std::towupper(c); });
            return s;
        }

        // 365 sorts above any dated edition.
        int rank(int year) { return year == 365 ? 9999 : year; }

        // Parse the ClickToRun ProductReleaseIds to a year (16.x can be any of
        // 2016/2019/2021/2024/365 — the major version alone can't tell them apart).
        int editionFromClickToRun()
        {
            wchar_t buf[1024]{};
            DWORD sz = sizeof(buf);
            if (RegGetValueW(HKEY_LOCAL_MACHINE,
                    L"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration",
                    L"ProductReleaseIds",
                    RRF_RT_REG_SZ | RRF_SUBKEY_WOW6464KEY,
                    nullptr, buf, &sz) != ERROR_SUCCESS)
                return 0;  // MSI / not Click-to-Run

            const std::wstring s = buf;
            auto has = [&](const wchar_t* t) { return s.find(t) != std::wstring::npos; };
            if (has(L"365") || has(L"O365")) return 365;
            int best = 0;
            if (has(L"2024")) best = (std::max)(best, 2024);
            if (has(L"2021")) best = (std::max)(best, 2021);
            if (has(L"2019")) best = (std::max)(best, 2019);
            if (has(L"2016")) best = (std::max)(best, 2016);
            return best;
        }

        struct Req { int year; bool needsDynamicArrays; };

        // Native availability metadata for the functions we shadow.
        const std::map<std::wstring, Req>& funcTable()
        {
            static const std::map<std::wstring, Req> t = {
                // Dynamic-array era (2021/365) — decided by the DA probe.
                { L"XLOOKUP",   { 2021, true } }, { L"XMATCH",     { 2021, true } },
                { L"FILTER",    { 2021, true } }, { L"SORT",       { 2021, true } },
                { L"SORTBY",    { 2021, true } }, { L"UNIQUE",     { 2021, true } },
                { L"SEQUENCE",  { 2021, true } }, { L"RANDARRAY",  { 2021, true } },
                // Post-DA (2024/365) — need DA *and* edition year ≥ 2024
                // (2021 has dynamic arrays but lacks these, so DA alone is insufficient).
                { L"TEXTSPLIT", { 2024, true } }, { L"TEXTBEFORE", { 2024, true } },
                { L"TEXTAFTER", { 2024, true } }, { L"VSTACK",     { 2024, true } },
                { L"HSTACK",    { 2024, true } }, { L"TAKE",       { 2024, true } },
                { L"DROP",      { 2024, true } }, { L"TOROW",      { 2024, true } },
                { L"TOCOL",     { 2024, true } }, { L"CHOOSEROWS", { 2024, true } },
                { L"CHOOSECOLS",{ 2024, true } }, { L"EXPAND",     { 2024, true } },
                { L"WRAPROWS",  { 2024, true } }, { L"WRAPCOLS",   { 2024, true } },
                // 2024 scalar-text (no dynamic-array output) — decided by year.
                { L"ARRAYTOTEXT", { 2024, false } }, { L"VALUETOTEXT", { 2024, false } },
                // REGEX (365, 2024) — REGEXEXTRACT can spill (DA), others scalar.
                { L"REGEXTEST",   { 2024, false } }, { L"REGEXREPLACE", { 2024, false } },
                { L"REGEXEXTRACT",{ 2024, true } },
                // GROUPBY/PIVOTBY (365) — text-aggregator variant, spills.
                { L"GROUPBY",     { 2024, true } }, { L"PIVOTBY",      { 2024, true } },
                // IMPORTTEXT/IMPORTCSV (365, 2026 롤아웃 — Beta부터) — post-DA로 취급:
                // DA 호스트(2024+)에선 EG. 접두(향후 네이티브와 공존), 구버전은 드롭인.
                { L"IMPORTTEXT",  { 2024, true } }, { L"IMPORTCSV",    { 2024, true } },
                // Pre-dynamic-array — decided by edition year.
                { L"IFS",       { 2019, false } }, { L"SWITCH",    { 2019, false } },
                { L"MAXIFS",    { 2019, false } }, { L"MINIFS",    { 2019, false } },
                { L"TEXTJOIN",  { 2019, false } }, { L"CONCAT",    { 2019, false } },
                // Excel 2013.
                { L"IFNA",      { 2013, false } }, { L"XOR",       { 2013, false } },
                { L"FORMULATEXT", { 2013, false } }, { L"ISFORMULA", { 2013, false } },
                { L"SHEET",     { 2013, false } }, { L"SHEETS",    { 2013, false } },
                { L"UNICHAR",   { 2013, false } }, { L"UNICODE",   { 2013, false } },
                { L"NUMBERVALUE", { 2013, false } }, { L"DAYS",     { 2013, false } },
                { L"ISOWEEKNUM",{ 2013, false } }, { L"ENCODEURL", { 2013, false } },
                { L"FILTERXML", { 2013, false } }, { L"WEBSERVICE",{ 2013, false } },
                { L"COMBINA",   { 2013, false } }, { L"PERMUTATIONA", { 2013, false } },
                { L"CEILING.MATH", { 2013, false } }, { L"FLOOR.MATH", { 2013, false } },
                { L"BASE",      { 2013, false } }, { L"DECIMAL",   { 2013, false } },
                { L"ARABIC",    { 2013, false } }, { L"RRI",       { 2013, false } },
                { L"PDURATION", { 2013, false } },
                { L"BITAND",    { 2013, false } }, { L"BITOR",     { 2013, false } },
                { L"BITXOR",    { 2013, false } }, { L"BITLSHIFT", { 2013, false } },
                { L"BITRSHIFT", { 2013, false } },
                // PERCENTOF (365/2024) — scalar.
                { L"PERCENTOF", { 2024, false } },
                // LET (2021) — year gate: 2016 → bare drop-in, 365/2021 → EG.LET.
                { L"LET", { 2021, false } },
                // IMAGE (365/2024) — 2016 → bare IMAGE, 365 → EG.IMAGE.
                { L"IMAGE", { 2024, false } },
                // TRIMRANGE (365/2024) — value-based edge trim (native is ref-based).
                { L"TRIMRANGE", { 2024, true } },
            };
            return t;
        }
    }

    int excelYear()
    {
        const int major = xloil::Environment::excelProcess().version;
        switch (major)
        {
        case 16: { int y = editionFromClickToRun(); return y ? y : 2016; }
        case 15: return 2013;
        case 14: return 2010;
        case 12: return 2007;
        case 11: return 2003;
        case 10: return 2002;
        case 9:  return 2000;
        case 8:  return 1997;
        default: return 0;
        }
    }

    bool supportsDynamicArrays()
    {
        return xloil::Environment::excelProcess().supportsDynamicArrays;
    }

    bool hasNativeFunction(const std::wstring& bareName)
    {
        const auto& t = funcTable();
        auto it = t.find(upper(bareName));
        if (it == t.end()) return false;                 // EGTools-only
        if (it->second.needsDynamicArrays)
        {
            if (!supportsDynamicArrays()) return false;
            // Functions that shipped *with* dynamic arrays (≤2021): the DA probe
            // is definitive and survives MSI editions where the year can't be read.
            if (it->second.year <= 2021) return true;
            // Post-DA functions: also require the edition year (2021 has DA but
            // not these). MSI-2021 falls back to 2016 → correctly treated as bare.
            return rank(excelYear()) >= rank(it->second.year);
        }
        return rank(excelYear()) >= rank(it->second.year);
    }

    bool needsXPrefix(const std::wstring& bareName)
    {
        const std::wstring u = upper(bareName);
        return u == L"SORT" || u == L"FILTER" || u == L"LET";
    }

    std::wstring registeredName(const std::wstring& bareName)
    {
        if (hasNativeFunction(bareName)) return L"EG." + bareName;
        if (needsXPrefix(bareName))      return L"x" + bareName;
        return bareName;
    }

    std::wstring storedName(const std::wstring& bareName)
    {
        const std::wstring u = upper(bareName);
        if (funcTable().find(u) == funcTable().end())
            return L"";   // EGTools-only / not a shadowed modern function

        // Empirically (tests/extract_stored_names.ps1), FILTER and SORT are
        // stored with the worksheet-only marker; every other shadowed function
        // uses the plain _xlfn. prefix.
        if (u == L"FILTER" || u == L"SORT")
            return L"_xlfn._xlws." + u;
        return L"_xlfn." + u;
    }

    const std::vector<std::wstring>& shadowedFunctionNames()
    {
        // Built once from funcTable() — every shadowed function is a conversion
        // candidate. Kept in insertion-independent (sorted) order for stable UI.
        static const std::vector<std::wstring> names = []
        {
            std::vector<std::wstring> v;
            v.reserve(funcTable().size());
            for (const auto& kv : funcTable()) v.push_back(kv.first);
            std::sort(v.begin(), v.end());
            return v;
        }();
        return names;
    }
}
