#include "Catalog.h"
#include "../core/I18n.h"
#include "../core/Version.h"

#include <algorithm>
#include <cwctype>

namespace egtools::intellisense
{
    namespace
    {
        std::wstring upper(std::wstring s)
        {
            std::transform(s.begin(), s.end(), s.begin(),
                           [](wchar_t c) { return (wchar_t)std::towupper(c); });
            return s;
        }
    }

    bool lookupFunction(const std::wstring& name, FuncInfo& out)
    {
        // The registered name may be bare (XLOOKUP) or EG.-prefixed (EG.XLOOKUP)
        // depending on the Excel version; the i18n catalog is keyed by the bare
        // name, so strip a leading "EG." before lookup.
        const std::wstring key = upper(name);
        std::wstring bare = key;
        if (bare.rfind(L"EG.", 0) == 0) bare = bare.substr(3);
        // Keyword-conflict names register as x<F> on legacy hosts (xSORT…) —
        // map back to the bare catalog key. Guarded by needsXPrefix so real
        // X-functions (XMATCH, XLOOKUP…) are never mis-stripped.
        else if (bare.size() > 1 && bare[0] == L'X' &&
                 egtools::core::needsXPrefix(bare.substr(1)))
            bare = bare.substr(1);

        std::wstring desc;
        std::vector<egtools::i18n::ArgMeta> args;
        egtools::i18n::RepeatSpec repeat;
        if (!egtools::i18n::func(bare, desc, args, &repeat))
            return false;

        out.name = key;   // show what the user typed (EG.XLOOKUP or XLOOKUP)
        out.help = desc;
        out.repeat = repeat;
        out.args.clear();
        for (const auto& a : args)
            out.args.push_back({ a.name, a.help, a.opt });
        return true;
    }

    bool lookupRegistered(const std::wstring& name, FuncInfo& out)
    {
        // Same prefix-stripping as lookupFunction to recover the bare name…
        const std::wstring key = upper(name);
        std::wstring bare = key;
        if (bare.rfind(L"EG.", 0) == 0) bare = bare.substr(3);
        else if (bare.size() > 1 && bare[0] == L'X' &&
                 egtools::core::needsXPrefix(bare.substr(1)))
            bare = bare.substr(1);

        // …then accept only when the host-specific registered name round-trips
        // to what the dropdown shows (bare BARCODE ↔ BARCODE, EG.XLOOKUP ↔
        // EG.XLOOKUP; native XLOOKUP on a modern host maps to EG.XLOOKUP ≠
        // XLOOKUP and is rejected — Excel's own tooltip covers it).
        if (upper(egtools::core::registeredName(bare)) != key)
            return false;
        return lookupFunction(name, out);
    }
}
