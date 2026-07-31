#include "FormulaParser.h"

#include <algorithm>
#include <regex>

namespace egtools::intellisense
{
    namespace
    {
        // Characters that cannot appear in a function name (same set as Excel-DNA's
        // forbiddenNameCharacters). Note '.' is NOT forbidden, so EG.HELLO is one name.
        bool isForbidden(wchar_t c)
        {
            static const std::wstring kForbidden =
                L" /-:;!@#$%^&*()+=,<>[]{}|'\"\\";
            return kForbidden.find(c) != std::wstring::npos;
        }

        size_t countCh(const std::wstring& s, wchar_t c)
        {
            return (size_t)std::count(s.begin(), s.end(), c);
        }

        // Function name = trailing run of non-forbidden chars, which must be
        // preceded by a forbidden char (mirrors regex "[forbidden](name)$").
        bool extractFunctionName(const std::wstring& before, std::wstring& name)
        {
            size_t end = before.size();
            size_t i = end;
            while (i > 0 && !isForbidden(before[i - 1])) --i;
            if (i == 0) return false;          // no forbidden char before the name
            if (i == end) return false;        // empty name
            name = before.substr(i);
            return true;
        }

        std::wstring regexReplace(const std::wstring& in, const std::wregex& re,
                                  const std::wstring& rep)
        {
            try { return std::regex_replace(in, re, rep); }
            catch (...) { return in; }
        }
    }

    bool tryGetFormulaInfo(const std::wstring& formulaPrefixIn,
                           std::wstring& functionName,
                           int& currentArgIndex,
                           wchar_t sep)
    {
        functionName.clear();
        currentArgIndex = -1;
        std::wstring p = formulaPrefixIn;

        // 1. Close an unterminated string so its commas/parens are ignored.
        if (countCh(p, L'"') % 2 != 0) p.push_back(L'"');

        // 2. Remove strings.
        p = regexReplace(p, std::wregex(L"\"[^\"]*\""), L"");

        // 3. [user fix] Remove the separator that follows a structured table
        //    reference group, e.g. Table1[[#All],[Field]] -> the inner ',' is not
        //    an argument separator.  Pattern: (\[\[.*?\])\s*<sep> -> $1
        try
        {
            std::wstring pat = L"(\\[\\[.*?\\])\\s*";
            pat.push_back(sep);   // sep is ',' or ';' — regex-safe
            p = regexReplace(p, std::wregex(pat), L"$1");
        }
        catch (...) {}

        // 4. Remove innermost parenthesised groups and spaces.
        p = regexReplace(p, std::wregex(L"(\\([^\\(\\)]*\\))| "), L"");

        // 5. Collapse remaining nested parentheses.
        std::wregex parenRe(L"\\([^\\(\\)]*\\)");
        try
        {
            while (std::regex_search(p, parenRe))
            {
                std::wstring next = std::regex_replace(p, parenRe, L"");
                if (next == p) break;
                p.swap(next);
            }
        }
        catch (...) {}

        // 6. Locate the last opening parenthesis.
        size_t lp = p.rfind(L'(');
        if (lp == std::wstring::npos) return false;

        // 7. Extract the function name immediately before it.
        if (!extractFunctionName(p.substr(0, lp), functionName))
            return false;

        // 8. Argument index = separators in the arguments part (arrays removed).
        std::wstring args = p.substr(lp);
        if (countCh(args, L'{') > countCh(args, L'}')) args.push_back(L'}');
        args = regexReplace(args, std::wregex(L"\\{[^\\}]*\\}"), L"");
        currentArgIndex = (int)countCh(args, sep);
        return true;
    }
}
