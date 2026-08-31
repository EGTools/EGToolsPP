// RegexMatchMode.h — shared helper for XLOOKUP/XMATCH match_mode 3 (regex).
//
// Native (365, 2024-08 rollout) treats lookup_value as a PCRE2 pattern and
// matches it against text cells (partial match, REGEXTEST semantics). The
// compat uses std::wregex (ECMAScript flavor), same as FxRegex.cpp. Patterns
// are case-sensitive; the documented native way to ignore case is a leading
// (?i), which ECMAScript lacks — emulated by stripping it and setting icase.

#pragma once

#include <xlOil/ExcelObj.h>
#include <regex>
#include <string>

namespace egtools::functions
{
    inline std::wregex regexForLookup(std::wstring pattern)
    {
        std::regex_constants::syntax_option_type flags = std::regex::ECMAScript;
        if (pattern.rfind(L"(?i)", 0) == 0)
        {
            flags |= std::regex::icase;
            pattern.erase(0, 4);
        }
        return std::wregex(pattern, flags);   // invalid pattern → std::regex_error
    }

    // TRUE if the cell is text and the pattern matches somewhere in it.
    // Non-text cells never match (native: regex applies to text values only).
    inline bool regexCellMatch(const xloil::ExcelObj& cell, const std::wregex& re)
    {
        return cell.type() == xloil::ExcelType::Str
            && std::regex_search(cell.toString(), re);
    }
}
