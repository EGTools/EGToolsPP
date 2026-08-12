// FxText.cpp — text functions.

#include "../core/Registry.h"
#include "../core/Apply.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <string>
#include <vector>
#include <cwctype>
#include <cwchar>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // Flatten an argument (scalar or range/array) into text pieces.
        // 오류 원소를 만나면 그 오류를 반환(집계 함수의 오류 전파 — 네이티브
        // TEXTJOIN/CONCAT은 입력에 오류가 하나라도 있으면 그 오류가 결과다).
        // 오류가 없으면 nullopt.
        std::optional<ExcelObj> collectText(
            const ExcelObj& v, std::vector<std::wstring>& out, bool skipEmpty)
        {
            std::optional<ExcelObj> err;
            auto add = [&](const ExcelObj& e)
            {
                if (err) return;
                if (e.type() == ExcelType::Missing) return;
                if (e.type() == ExcelType::Err) { err = ExcelObj(e); return; }
                std::wstring s = e.toString();
                if (skipEmpty && s.empty()) return;
                out.push_back(std::move(s));
            };
            if (v.isMissing()) return err;
            if (v.isType(ExcelType::Multi))
            {
                ExcelArray a(v);
                const size_t n = (size_t)a.nRows() * a.nCols();
                for (size_t k = 0; k < n; ++k) { add(a.at(k)); if (err) break; }
            }
            else add(v);
            return err;
        }

        std::wstring lower(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)std::towlower(c);
            return s;
        }

        // Find the i-th (1-based) occurrence of `delim` in `text`. i<0 counts
        // from the end. Returns npos if not found.
        size_t findInstance(const std::wstring& text, const std::wstring& delim,
                            int instance, bool ci)
        {
            if (delim.empty()) return std::wstring::npos;
            const std::wstring t = ci ? lower(text) : text;
            const std::wstring d = ci ? lower(delim) : delim;
            if (instance >= 0)
            {
                size_t pos = 0; int n = 0;
                while ((pos = t.find(d, pos)) != std::wstring::npos)
                    if (++n == instance) return pos; else pos += d.size();
                return std::wstring::npos;
            }
            // negative: from the end
            size_t pos = t.size(); int n = 0;
            while (pos != std::wstring::npos)
            {
                size_t found = t.rfind(d, pos == 0 ? 0 : pos - 1);
                if (found == std::wstring::npos) break;
                if (++n == -instance) return found;
                if (found == 0) break;
                pos = found;
            }
            return std::wstring::npos;
        }

        // Count non-overlapping occurrences of `delim` in `text`.
        int countInstances(const std::wstring& text, const std::wstring& delim, bool ci)
        {
            if (delim.empty()) return 0;
            const std::wstring t = ci ? lower(text) : text;
            const std::wstring d = ci ? lower(delim) : delim;
            int n = 0;
            size_t pos = 0;
            while ((pos = t.find(d, pos)) != std::wstring::npos) { ++n; pos += d.size(); }
            return n;
        }

        // Split `text` by any of the delimiters in `delims`.
        std::vector<std::wstring> splitBy(const std::wstring& text,
                                          const std::vector<std::wstring>& delims, bool ci)
        {
            std::vector<std::wstring> parts;
            const std::wstring t = ci ? lower(text) : text;
            size_t pos = 0;
            while (pos <= text.size())
            {
                size_t best = std::wstring::npos, bestLen = 0;
                for (const auto& draw : delims)
                {
                    if (draw.empty()) continue;
                    const std::wstring d = ci ? lower(draw) : draw;
                    size_t f = t.find(d, pos);
                    if (f != std::wstring::npos && (best == std::wstring::npos || f < best))
                        { best = f; bestLen = d.size(); }
                }
                if (best == std::wstring::npos) { parts.push_back(text.substr(pos)); break; }
                parts.push_back(text.substr(pos, best - pos));
                pos = best + bestLen;
            }
            return parts;
        }

        // VALUETOTEXT semantics: strict mode wraps text in double quotes
        // (internal quotes doubled); concise mode is the plain display text.
        std::wstring valueToText(const ExcelObj& v, bool strict)
        {
            if (strict && v.type() == ExcelType::Str)
            {
                const std::wstring s = v.toString();
                std::wstring out = L"\"";
                for (wchar_t c : s) { if (c == L'"') out += L"\"\""; else out.push_back(c); }
                out += L"\"";
                return out;
            }
            return v.toString();
        }

        // NUMBERVALUE: locale-independent parse. Whitespace ignored, group
        // separators removed, decimal separator normalised, trailing % scaled.
        ExcelObj parseNumberValue(const std::wstring& text, wchar_t dec, wchar_t grp)
        {
            std::wstring s;
            for (wchar_t c : text)
                if (!std::iswspace(c)) s.push_back(c);
            if (s.empty()) return ExcelObj(0.0);

            std::wstring t;
            for (wchar_t c : s)
            {
                if (c == grp) continue;
                t.push_back(c == dec ? L'.' : c);
            }
            int pct = 0;
            while (!t.empty() && t.back() == L'%') { ++pct; t.pop_back(); }
            if (t.empty()) return ExcelObj(CellError::Value);

            wchar_t* end = nullptr;
            double v = std::wcstod(t.c_str(), &end);
            if (end == t.c_str() || (end && *end != L'\0'))
                return ExcelObj(CellError::Value);
            for (int i = 0; i < pct; ++i) v /= 100.0;
            return ExcelObj(v);
        }

        // UNICHAR(number): Unicode code point → character.
        ExcelObj unicharOne(const ExcelObj& e)
        {
            const int code = e.get<int>(0);
            if (code < 1 || code > 0x10FFFF || (code >= 0xD800 && code <= 0xDFFF))
                return ExcelObj(CellError::Value);

            std::wstring s;
            if (code <= 0xFFFF)
                s.push_back((wchar_t)code);
            else
            {
                const int c = code - 0x10000;
                s.push_back((wchar_t)(0xD800 + (c >> 10)));
                s.push_back((wchar_t)(0xDC00 + (c & 0x3FF)));
            }
            return ExcelObj(std::wstring_view(s));
        }

        // UNICODE(text): code point of the first character.
        ExcelObj unicodeOne(const ExcelObj& e)
        {
            const std::wstring s = e.toString();
            if (s.empty())
                return ExcelObj(CellError::Value);

            unsigned cp = (unsigned)(wchar_t)s[0];
            if (cp >= 0xD800 && cp <= 0xDBFF && s.size() >= 2)   // high surrogate
            {
                const unsigned lo = (unsigned)(wchar_t)s[1];
                if (lo >= 0xDC00 && lo <= 0xDFFF)
                    cp = 0x10000 + ((cp - 0xD800) << 10) + (lo - 0xDC00);
            }
            return ExcelObj((double)cp);
        }
    }

    void registerText()
    {
        egtools::core::registerFn(L"UNICHAR",
            [](const ExcelObj& n) -> ExcelObj* { return egtools::core::mapUnary(n, unicharOne); });
        egtools::core::registerFn(L"UNICODE",
            [](const ExcelObj& t) -> ExcelObj* { return egtools::core::mapUnary(t, unicodeOne); });

        // NUMBERVALUE(text, [decimal_separator], [group_separator]).
        egtools::core::registerFn(L"NUMBERVALUE",
            [](const ExcelObj& text, const ExcelObj& decSep, const ExcelObj& grpSep) -> ExcelObj*
            {
                wchar_t dec = L'.', grp = L',';
                if (!decSep.isMissing()) { auto s = decSep.toString(); if (!s.empty()) dec = s[0]; }
                if (!grpSep.isMissing()) { auto s = grpSep.toString(); if (!s.empty()) grp = s[0]; }
                return egtools::core::mapUnary(text, [dec, grp](const ExcelObj& e)
                {
                    return parseNumberValue(e.toString(), dec, grp);
                });
            });

        // TEXTJOIN(delimiter, ignore_empty, text1, …) — variadic up to 255 args.
        egtools::core::registerRawFn(L"TEXTJOIN",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                // 오류 전파(네이티브 정합): 구분자·플래그·연결값 어디든 오류가
                // 있으면 그 오류가 결과다(ignore_empty와 무관).
                if (args[0] && args[0]->type() == ExcelType::Err)
                    return returnValue(ExcelObj(*args[0]));
                if (args[1] && args[1]->type() == ExcelType::Err)
                    return returnValue(ExcelObj(*args[1]));
                const bool skip = args[0] && args[1] && !args[1]->isMissing()
                    ? args[1]->get<double>(1.0) != 0.0 : true;
                // 구분자 배열 = 로테이션(원소를 순환하며 사용, 네이티브 정합 U17).
                std::vector<std::wstring> delims;
                if (args[0] && args[0]->isType(ExcelType::Multi))
                {
                    ExcelArray da(*args[0]);
                    const size_t n = (size_t)da.nRows() * da.nCols();
                    for (size_t i = 0; i < n; ++i)
                    {
                        if (da.at(i).type() == ExcelType::Err)
                            return returnValue(ExcelObj(da.at(i)));
                        delims.push_back(da.at(i).toString());
                    }
                }
                if (delims.empty())
                    delims.push_back(args[0] ? args[0]->toString() : std::wstring());
                std::vector<std::wstring> parts;
                for (size_t i = 2; i < info.numArgs(); ++i)
                    if (auto err = collectText(*args[i], parts, skip))
                        return returnValue(std::move(*err));
                std::wstring out;
                for (size_t i = 0; i < parts.size(); ++i)
                    { if (i) out += delims[(i - 1) % delims.size()]; out += parts[i]; }
                return returnValue(ExcelObj(std::wstring_view(out)));
            });

        // CONCAT(text1, …) — concatenate all text (incl. ranges), variadic.
        egtools::core::registerRawFn(L"CONCAT",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                std::vector<std::wstring> parts;
                for (size_t i = 0; i < info.numArgs(); ++i)
                    if (auto err = collectText(*args[i], parts, false))
                        return returnValue(std::move(*err));   // 오류 전파(네이티브 정합)
                std::wstring out;
                for (auto& s : parts) out += s;
                return returnValue(ExcelObj(std::wstring_view(out)));
            });

        // TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end],
        //            [if_not_found]) — 네이티브와 동일 6인수.
        // match_end=1이면 텍스트의 끝(양수 instance)/시작(음수 instance)을 가상의
        // 구분자 1개로 간주: 개수+1번째 instance까지 허용된다.
        egtools::core::registerFn(L"TEXTBEFORE",
            [](const ExcelObj& textA, const ExcelObj& delimA, const ExcelObj& instA,
               const ExcelObj& modeA, const ExcelObj& endA, const ExcelObj& nfA) -> ExcelObj*
            {
                // 구분자 배열(네이티브: 대체 후보 목록)은 미구현 — 오답 대신 거부.
                if (delimA.isType(ExcelType::Multi)) return returnValue(CellError::Value);
                const std::wstring delim = delimA.toString();
                const int inst = instA.isMissing() ? 1 : instA.get<int>(1);
                if (inst == 0) return returnValue(CellError::Value);
                const bool ci = modeA.isMissing() ? false : (modeA.get<double>(0.0) != 0.0);
                const bool matchEnd = endA.isMissing() ? false : (endA.get<double>(0.0) != 0.0);

                // text는 원소별 리프팅(네이티브 정합, plan/22 U12).
                return egtools::core::mapLift([&](const ExcelObj& t) -> ExcelObj
                {
                    const std::wstring text = t.toString();
                    const size_t at = findInstance(text, delim, inst, ci);
                    if (at != std::wstring::npos)
                        return ExcelObj(text.substr(0, at));
                    if (matchEnd)
                    {
                        const int n = countInstances(text, delim, ci);
                        if (inst == n + 1)   // 가상 구분자 = 텍스트 끝
                            return ExcelObj(std::wstring_view(text));
                        if (inst == -(n + 1))   // 가상 구분자 = 텍스트 시작
                            return ExcelObj(std::wstring_view(L""));
                    }
                    return nfA.isMissing() ? ExcelObj(CellError::NA)
                                           : egtools::core::demote(ExcelObj(nfA));
                }, textA);
            });

        // TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end],
        //           [if_not_found]) — 네이티브와 동일 6인수.
        egtools::core::registerFn(L"TEXTAFTER",
            [](const ExcelObj& textA, const ExcelObj& delimA, const ExcelObj& instA,
               const ExcelObj& modeA, const ExcelObj& endA, const ExcelObj& nfA) -> ExcelObj*
            {
                if (delimA.isType(ExcelType::Multi)) return returnValue(CellError::Value);
                const std::wstring delim = delimA.toString();
                const int inst = instA.isMissing() ? 1 : instA.get<int>(1);
                if (inst == 0) return returnValue(CellError::Value);
                const bool ci = modeA.isMissing() ? false : (modeA.get<double>(0.0) != 0.0);
                const bool matchEnd = endA.isMissing() ? false : (endA.get<double>(0.0) != 0.0);

                return egtools::core::mapLift([&](const ExcelObj& t) -> ExcelObj
                {
                    const std::wstring text = t.toString();
                    const size_t at = findInstance(text, delim, inst, ci);
                    if (at != std::wstring::npos)
                        return ExcelObj(text.substr(at + delim.size()));
                    if (matchEnd)
                    {
                        const int n = countInstances(text, delim, ci);
                        if (inst == n + 1)   // 가상 구분자 = 텍스트 끝
                            return ExcelObj(std::wstring_view(L""));
                        if (inst == -(n + 1))   // 가상 구분자 = 텍스트 시작
                            return ExcelObj(std::wstring_view(text));
                    }
                    return nfA.isMissing() ? ExcelObj(CellError::NA)
                                           : egtools::core::demote(ExcelObj(nfA));
                }, textA);
            });

        // TEXTSPLIT(text, col_delim, [row_delim], [ignore_empty], [match_mode], [pad_with]).
        egtools::core::registerFn(L"TEXTSPLIT",
            [](const ExcelObj& textA, const ExcelObj& colD, const ExcelObj& rowD,
               const ExcelObj& ignoreA, const ExcelObj& modeA, const ExcelObj& padA) -> ExcelObj*
            {
                const bool ci = modeA.isMissing() ? false : (modeA.get<double>(0.0) != 0.0);
                const bool ignore = ignoreA.isMissing() ? false : (ignoreA.get<double>(0.0) != 0.0);
                const ExcelObj pad = padA.isMissing() ? ExcelObj(CellError::NA) : ExcelObj(padA);

                // 구분자 배열 = 대체 후보 목록(네이티브 정합, plan/22 T2).
                auto collectDelims = [](const ExcelObj& o, std::vector<std::wstring>& out)
                {
                    if (o.isMissing()) return;
                    if (o.isType(ExcelType::Multi))
                    {
                        ExcelArray a(o);
                        const size_t n = (size_t)a.nRows() * a.nCols();
                        for (size_t i = 0; i < n; ++i)
                        {
                            std::wstring s = a.at(i).toString();
                            if (!s.empty()) out.push_back(std::move(s));
                        }
                        return;
                    }
                    out.push_back(o.toString());
                };
                std::vector<std::wstring> colDelims, rowDelims;
                collectDelims(colD, colDelims);
                collectDelims(rowD, rowDelims);

                // text 배열 → 강등 리프팅(원소별 첫 토큰, 네이티브 정합 T1).
                return egtools::core::mapLift([&](const ExcelObj& te) -> ExcelObj
                {
                    const std::wstring text = te.toString();

                    // Rows: split by row_delim (or whole text as one row).
                    std::vector<std::wstring> rowStrs = rowDelims.empty()
                        ? std::vector<std::wstring>{ text } : splitBy(text, rowDelims, ci);

                    std::vector<std::vector<std::wstring>> grid;
                    size_t maxCols = 0;
                    for (auto& rs : rowStrs)
                    {
                        std::vector<std::wstring> cells = colDelims.empty()
                            ? std::vector<std::wstring>{ rs } : splitBy(rs, colDelims, ci);
                        if (ignore)
                        {
                            std::vector<std::wstring> kept;
                            for (auto& c : cells) if (!c.empty()) kept.push_back(c);
                            cells.swap(kept);
                        }
                        if (!cells.empty() || !ignore) { maxCols = std::max(maxCols, cells.size()); grid.push_back(std::move(cells)); }
                    }
                    if (ignore)
                    {
                        std::vector<std::vector<std::wstring>> kept;
                        for (auto& r : grid) if (!r.empty()) kept.push_back(std::move(r));
                        grid.swap(kept);
                    }
                    if (grid.empty() || maxCols == 0) return ExcelObj(CellError::Value);

                    std::vector<ExcelObj> vals;
                    vals.reserve(grid.size() * maxCols);
                    for (auto& r : grid)
                        for (size_t j = 0; j < maxCols; ++j)
                            vals.emplace_back(j < r.size() ? ExcelObj(std::wstring_view(r[j])) : pad);
                    return egtools::core::makeArray(
                        (xloil::ExcelArrayBuilder::row_t)grid.size(),
                        (xloil::ExcelArrayBuilder::col_t)maxCols, vals);
                }, textA);
            });

        // VALUETOTEXT(value, [format]) — 0 concise (default), 1 strict (quote text).
        // value는 원소별 리프팅(네이티브 정합, plan/22 U14).
        egtools::core::registerFn(L"VALUETOTEXT",
            [](const ExcelObj& value, const ExcelObj& fmt) -> ExcelObj*
            {
                const bool strict = fmt.isMissing() ? false : (fmt.get<int>(0) == 1);
                return egtools::core::mapUnary(value, [strict](const ExcelObj& e)
                {
                    return ExcelObj(std::wstring_view(valueToText(e, strict)));
                });
            });

        // ARRAYTOTEXT(array, [format]) — 0 list "a, b, c"; 1 strict "{a,b;c,d}".
        egtools::core::registerFn(L"ARRAYTOTEXT",
            [](const ExcelObj& array, const ExcelObj& fmt) -> ExcelObj*
            {
                const bool strict = fmt.isMissing() ? false : (fmt.get<int>(0) == 1);
                if (!array.isType(ExcelType::Multi))
                {
                    std::wstring s = valueToText(array, strict);
                    return returnValue(ExcelObj(std::wstring_view(s)));
                }
                ExcelArray a(array);
                std::wstring out;
                if (strict) out = L"{";
                for (ExcelArray::row_t r = 0; r < a.nRows(); ++r)
                {
                    if (r) out += strict ? L";" : L", ";
                    for (ExcelArray::col_t c = 0; c < a.nCols(); ++c)
                    {
                        if (c) out += strict ? L"," : L", ";
                        out += valueToText(a.at(r, c), strict);
                    }
                }
                if (strict) out += L"}";
                return returnValue(ExcelObj(std::wstring_view(out)));
            });
    }
}
