// FxRegex.cpp — regular-expression functions (REGEXTEST, REGEXEXTRACT,
// REGEXREPLACE). Uses std::wregex (ECMAScript flavor). $1.. backreferences
// supported in REGEXREPLACE replacement.
//
// text/pattern/replacement lift element-wise with broadcast (네이티브 정합,
// plan/22 T5·T6·U12): array in → array out; REGEXEXTRACT의 mode 1·2처럼
// 원소별 결과가 배열이면 네이티브처럼 첫 값으로 강등된다(mapLift).

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "../core/Apply.h"

#include <xlOil/xlOil.h>
#include <optional>
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

        // Per-element regex: reuse a precompiled regex when the pattern is a
        // scalar; otherwise compile the element's pattern.
        struct RegexSource
        {
            bool icase;
            std::optional<std::wregex> pre;   // set when pattern arg is scalar
            explicit RegexSource(const ExcelObj& patA, bool ic) : icase(ic)
            {
                if (!patA.isType(ExcelType::Multi))
                    pre = buildRegex(patA.toString(), ic);   // invalid → throws → #VALUE!
            }
            std::wregex make(const ExcelObj& p) const
            {
                return buildRegex(p.toString(), icase);
            }
        };
    }

    void registerRegex()
    {
        // REGEXTEST(text, pattern, [case_insensitivity]) — TRUE if a match exists.
        egtools::core::registerFn(L"REGEXTEST",
            [](const ExcelObj& textA, const ExcelObj& patA, const ExcelObj& ciA) -> ExcelObj*
            {
                const RegexSource src(patA, caseInsensitive(ciA));
                return egtools::core::mapLift(
                    [&](const ExcelObj& t, const ExcelObj& p) -> ExcelObj
                    {
                        try
                        {
                            const std::wstring text = t.toString();
                            if (src.pre)
                                return ExcelObj(std::regex_search(text, *src.pre));
                            const auto re = src.make(p);
                            return ExcelObj(std::regex_search(text, re));
                        }
                        catch (...) { return ExcelObj(CellError::Value); }
                    },
                    textA, patA);
            });

        // REGEXEXTRACT(text, pattern, [return_mode], [case_insensitivity]).
        // return_mode: 0 first match (default), 1 all matches (column),
        //              2 capturing groups of the first match (row).
        egtools::core::registerFn(L"REGEXEXTRACT",
            [](const ExcelObj& textA, const ExcelObj& patA, const ExcelObj& modeA,
               const ExcelObj& ciA) -> ExcelObj*
            {
                const int mode = modeA.isMissing() ? 0 : modeA.get<int>(0);
                const RegexSource src(patA, caseInsensitive(ciA));
                return egtools::core::mapLift(
                    [&](const ExcelObj& t, const ExcelObj& p) -> ExcelObj
                    {
                        try
                        {
                            const std::wstring text = t.toString();
                            std::optional<std::wregex> local;
                            const std::wregex& re = src.pre ? *src.pre
                                                            : local.emplace(src.make(p));

                            if (mode == 1)
                            {
                                std::vector<ExcelObj> vals;
                                for (std::wsregex_iterator it(text.begin(), text.end(), re), end;
                                     it != end; ++it)
                                    vals.emplace_back(std::wstring_view(it->str()));
                                if (vals.empty()) return ExcelObj(CellError::NA);
                                return egtools::core::makeArray(
                                    (ExcelArrayBuilder::row_t)vals.size(), 1, vals);
                            }

                            std::wsmatch m;
                            if (!std::regex_search(text, m, re)) return ExcelObj(CellError::NA);

                            if (mode == 2)
                            {
                                std::vector<ExcelObj> vals;
                                for (size_t i = 1; i < m.size(); ++i)
                                    vals.emplace_back(std::wstring_view(m[i].str()));
                                if (vals.empty()) return ExcelObj(CellError::NA);
                                return egtools::core::makeArray(
                                    1, (ExcelArrayBuilder::col_t)vals.size(), vals);
                            }

                            return ExcelObj(std::wstring_view(m[0].str()));   // mode 0
                        }
                        catch (...) { return ExcelObj(CellError::Value); }
                    },
                    textA, patA);
            });

        // REGEXREPLACE(text, pattern, replacement, [occurrence], [case_insensitivity]).
        // occurrence: 0/omitted = all; N>0 = only the Nth match. $1.. in replacement.
        egtools::core::registerFn(L"REGEXREPLACE",
            [](const ExcelObj& textA, const ExcelObj& patA, const ExcelObj& repA,
               const ExcelObj& occA, const ExcelObj& ciA) -> ExcelObj*
            {
                const int occ = occA.isMissing() ? 0 : occA.get<int>(0);
                const RegexSource src(patA, caseInsensitive(ciA));
                return egtools::core::mapLift(
                    [&](const ExcelObj& t, const ExcelObj& p, const ExcelObj& r) -> ExcelObj
                    {
                        try
                        {
                            const std::wstring text = t.toString();
                            const std::wstring rep = r.toString();
                            std::optional<std::wregex> local;
                            const std::wregex& re = src.pre ? *src.pre
                                                            : local.emplace(src.make(p));

                            if (occ <= 0)
                                return ExcelObj(std::wstring_view(
                                    std::regex_replace(text, re, rep)));

                            // Replace only the Nth occurrence.
                            std::wstring out;
                            out.reserve(text.size());
                            int n = 0;
                            auto last = text.cbegin();
                            for (std::wsregex_iterator it(text.begin(), text.end(), re), end;
                                 it != end; ++it)
                            {
                                const auto& mm = *it;
                                ++n;
                                out.append(last, text.cbegin() + mm.position(0));
                                out.append(n == occ ? mm.format(rep) : mm[0].str());
                                last = text.cbegin() + mm.position(0) + mm.length(0);
                            }
                            out.append(last, text.cend());
                            return ExcelObj(std::wstring_view(out));
                        }
                        catch (...) { return ExcelObj(CellError::Value); }
                    },
                    textA, patA, repA);
            });
    }
}
