// FxRegex.cpp — regular-expression functions (REGEXTEST, REGEXEXTRACT,
// REGEXREPLACE). Uses std::wregex (ECMAScript flavor). $1.. backreferences
// supported in REGEXREPLACE replacement.

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"

#include <xlOil/xlOil.h>
#include <string>
#include <vector>
#include <regex>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        std::wregex buildRegex(const std::wstring& pattern, bool icase)
        {
            std::regex_constants::syntax_option_type flags = std::regex::ECMAScript;
            if (icase) flags |= std::regex::icase;
            return std::wregex(pattern, flags);
        }
        bool caseInsensitive(const ExcelObj& o)
        {
            return o.isMissing() ? false : (o.get<int>(0) != 0);
        }
    }

    void registerRegex()
    {
        // REGEXTEST(text, pattern, [case_insensitivity]) — TRUE if a match exists.
        egtools::core::registerFn(L"REGEXTEST",
            [](const ExcelObj& textA, const ExcelObj& patA, const ExcelObj& ciA) -> ExcelObj*
            {
                try
                {
                    const std::wstring text = textA.toString();
                    auto re = buildRegex(patA.toString(), caseInsensitive(ciA));
                    return returnValue(ExcelObj(std::regex_search(text, re)));
                }
                catch (...) { return returnValue(CellError::Value); }
            });

        // REGEXEXTRACT(text, pattern, [return_mode], [case_insensitivity]).
        // return_mode: 0 first match (default), 1 all matches (column),
        //              2 capturing groups of the first match (row).
        egtools::core::registerFn(L"REGEXEXTRACT",
            [](const ExcelObj& textA, const ExcelObj& patA, const ExcelObj& modeA,
               const ExcelObj& ciA) -> ExcelObj*
            {
                try
                {
                    const std::wstring text = textA.toString();
                    const int mode = modeA.isMissing() ? 0 : modeA.get<int>(0);
                    auto re = buildRegex(patA.toString(), caseInsensitive(ciA));

                    if (mode == 1)
                    {
                        std::vector<ExcelObj> vals;
                        for (std::wsregex_iterator it(text.begin(), text.end(), re), end; it != end; ++it)
                            vals.emplace_back(std::wstring_view(it->str()));
                        if (vals.empty()) return returnValue(CellError::NA);
                        return egtools::core::output(egtools::core::makeArray(
                            (ExcelArrayBuilder::row_t)vals.size(), 1, vals));
                    }

                    std::wsmatch m;
                    if (!std::regex_search(text, m, re)) return returnValue(CellError::NA);

                    if (mode == 2)
                    {
                        std::vector<ExcelObj> vals;
                        for (size_t i = 1; i < m.size(); ++i)
                            vals.emplace_back(std::wstring_view(m[i].str()));
                        if (vals.empty()) return returnValue(CellError::NA);
                        return egtools::core::output(egtools::core::makeArray(
                            1, (ExcelArrayBuilder::col_t)vals.size(), vals));
                    }

                    std::wstring s = m[0].str();   // mode 0
                    return returnValue(ExcelObj(std::wstring_view(s)));
                }
                catch (...) { return returnValue(CellError::Value); }
            });

        // REGEXREPLACE(text, pattern, replacement, [occurrence], [case_insensitivity]).
        // occurrence: 0/omitted = all; N>0 = only the Nth match. $1.. in replacement.
        egtools::core::registerFn(L"REGEXREPLACE",
            [](const ExcelObj& textA, const ExcelObj& patA, const ExcelObj& repA,
               const ExcelObj& occA, const ExcelObj& ciA) -> ExcelObj*
            {
                try
                {
                    const std::wstring text = textA.toString();
                    const std::wstring rep = repA.toString();
                    const int occ = occA.isMissing() ? 0 : occA.get<int>(0);
                    auto re = buildRegex(patA.toString(), caseInsensitive(ciA));

                    if (occ <= 0)
                    {
                        std::wstring out = std::regex_replace(text, re, rep);
                        return returnValue(ExcelObj(std::wstring_view(out)));
                    }

                    // Replace only the Nth occurrence.
                    std::wstring out;
                    out.reserve(text.size());
                    int n = 0;
                    auto last = text.cbegin();
                    for (std::wsregex_iterator it(text.begin(), text.end(), re), end; it != end; ++it)
                    {
                        const auto& mm = *it;
                        ++n;
                        out.append(last, text.cbegin() + mm.position(0));
                        out.append(n == occ ? mm.format(rep) : mm[0].str());
                        last = text.cbegin() + mm.position(0) + mm.length(0);
                    }
                    out.append(last, text.cend());
                    return returnValue(ExcelObj(std::wstring_view(out)));
                }
                catch (...) { return returnValue(CellError::Value); }
            });
    }
}
