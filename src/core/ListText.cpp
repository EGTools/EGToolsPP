// ListText.cpp — 셀 내 목록/배열 텍스트 표기 파서·직렬화 (ListText.h 참조).

#include "ListText.h"

#include <xlOil/ExcelArray.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <optional>

using namespace xloil;

namespace egtools::core
{
    namespace
    {
        struct ErrName { const wchar_t* text; CellError err; };
        constexpr ErrName kErrNames[] = {
            { L"#NULL!",        CellError::Null },
            { L"#DIV/0!",       CellError::Div0 },
            { L"#VALUE!",       CellError::Value },
            { L"#REF!",         CellError::Ref },
            { L"#NAME?",        CellError::Name },
            { L"#NUM!",         CellError::Num },
            { L"#N/A",          CellError::NA },
            { L"#GETTING_DATA", CellError::GettingData },
        };

        bool isSpace(wchar_t c) { return c == L' ' || c == L'\t' || c == L'\r' || c == L'\n'; }

        std::wstring trim(const std::wstring& s)
        {
            size_t a = 0, b = s.size();
            while (a < b && isSpace(s[a])) ++a;
            while (b > a && isSpace(s[b - 1])) --b;
            return s.substr(a, b - a);
        }

        bool iequals(const std::wstring& a, const wchar_t* b)
        {
            size_t i = 0;
            for (; i < a.size() && b[i]; ++i)
                if (std::towlower(a[i]) != std::towlower(b[i])) return false;
            return i == a.size() && !b[i];
        }

        // 숫자/논리/오류로 읽히는 토큰을 그 값으로, 아니면 nullopt(텍스트).
        std::optional<ExcelObj> classify(const std::wstring& t)
        {
            if (t.empty()) return std::nullopt;
            if (iequals(t, L"TRUE"))  return ExcelObj(true);
            if (iequals(t, L"FALSE")) return ExcelObj(false);
            if (t[0] == L'#')
            {
                for (const auto& e : kErrNames)
                    if (iequals(t, e.text)) return ExcelObj(e.err);
                return std::nullopt;
            }
            const wchar_t c0 = t[0];
            if (!(std::iswdigit(c0) || c0 == L'+' || c0 == L'-' || c0 == L'.'))
                return std::nullopt;
            wchar_t* end = nullptr;
            const double d = std::wcstod(t.c_str(), &end);
            if (!end || *end != L'\0') return std::nullopt;
            // wcstod가 받는 inf/nan 표기는 Excel 값이 아니므로 텍스트로 남긴다.
            for (wchar_t c : t)
                if (std::iswalpha(c) && std::towlower(c) != L'e') return std::nullopt;
            return ExcelObj(d);
        }

        // `"…"` 인용 토큰이면 내부 `""`를 풀어 돌려준다(아니면 nullopt).
        std::optional<std::wstring> unquote(const std::wstring& t)
        {
            if (t.size() < 2 || t.front() != L'"' || t.back() != L'"') return std::nullopt;
            std::wstring out;
            for (size_t i = 1; i + 1 < t.size(); ++i)
            {
                if (t[i] == L'"')
                {
                    if (i + 2 < t.size() && t[i + 1] == L'"') { out.push_back(L'"'); ++i; }
                    else return std::nullopt;   // 홀로 선 따옴표 — 인용 토큰 아님
                }
                else out.push_back(t[i]);
            }
            return out;
        }

        // 깊이 0·따옴표 밖의 `delim`으로 분할. 따옴표/중괄호가 안 맞으면 그대로
        // 한 토큰이 된다(파싱 불능 → 텍스트 1개 규약).
        std::vector<std::wstring> splitTop(const std::wstring& s, wchar_t delim)
        {
            std::vector<std::wstring> out;
            std::wstring cur;
            int depth = 0;
            bool inQ = false;
            for (size_t i = 0; i < s.size(); ++i)
            {
                const wchar_t c = s[i];
                if (inQ)
                {
                    cur.push_back(c);
                    if (c == L'"')
                    {
                        if (i + 1 < s.size() && s[i + 1] == L'"') { cur.push_back(L'"'); ++i; }
                        else inQ = false;
                    }
                    continue;
                }
                if (c == L'"') { inQ = true; cur.push_back(c); continue; }
                if (c == L'{') { ++depth; cur.push_back(c); continue; }
                if (c == L'}') { if (depth > 0) --depth; cur.push_back(c); continue; }
                if (c == delim && depth == 0) { out.push_back(cur); cur.clear(); continue; }
                cur.push_back(c);
            }
            out.push_back(cur);
            return out;
        }

        // `{`로 시작해 짝이 맞는 `}`로 끝나는가(바깥 중괄호 한 쌍이 전체를 감쌈).
        bool isBraced(const std::wstring& t)
        {
            if (t.size() < 2 || t.front() != L'{' || t.back() != L'}') return false;
            int depth = 0;
            bool inQ = false;
            for (size_t i = 0; i < t.size(); ++i)
            {
                const wchar_t c = t[i];
                if (inQ) { if (c == L'"') inQ = false; continue; }
                if (c == L'"') { inQ = true; continue; }
                if (c == L'{') ++depth;
                else if (c == L'}')
                {
                    --depth;
                    if (depth == 0 && i + 1 != t.size()) return false;   // 중간에 닫힘
                }
            }
            return depth == 0 && !inQ;
        }

        ListNode parseBraced(const std::wstring& braced);

        ListNode parseItem(const std::wstring& tok)
        {
            const std::wstring t = trim(tok);
            if (isBraced(t)) return parseBraced(t);
            if (auto q = unquote(t)) return ListNode::scalar(ExcelObj(std::wstring_view(*q)));
            if (auto v = classify(t)) return ListNode::scalar(std::move(*v));
            return ListNode::scalar(ExcelObj(std::wstring_view(t)));
        }

        // `{a,b;c,d}` → 행우선 2D. 들쭉날쭉 행은 #N/A로 채운다.
        ListNode parseBraced(const std::wstring& braced)
        {
            const std::wstring inner = braced.substr(1, braced.size() - 2);
            std::vector<std::vector<ListNode>> rows;
            size_t cols = 0;
            for (const auto& r : splitTop(inner, L';'))
            {
                std::vector<ListNode> row;
                for (const auto& tok : splitTop(r, L',')) row.push_back(parseItem(tok));
                cols = std::max(cols, row.size());
                rows.push_back(std::move(row));
            }
            std::vector<ListNode> items;
            items.reserve(rows.size() * cols);
            for (auto& r : rows)
            {
                for (size_t c = 0; c < cols; ++c)
                    items.push_back(c < r.size() ? std::move(r[c])
                                                 : ListNode::scalar(ExcelObj(CellError::NA)));
            }
            return ListNode::array(rows.size(), cols, std::move(items));
        }

        bool needsQuoteConcise(const std::wstring& s)
        {
            if (s.empty()) return false;
            if (isSpace(s.front()) || isSpace(s.back())) return true;
            for (wchar_t c : s)
                if (c == L',' || c == L';' || c == L'{' || c == L'}' || c == L'"') return true;
            return classify(s).has_value();   // 숫자/논리/오류로 읽힐 텍스트
        }

        std::wstring quoted(const std::wstring& s)
        {
            std::wstring out = L"\"";
            for (wchar_t c : s) { if (c == L'"') out += L"\"\""; else out.push_back(c); }
            out += L"\"";
            return out;
        }

        void formatStrict(const ListNode& n, std::wstring& out)
        {
            if (!n.isArray) { out += scalarText(n.value, true); return; }
            out += L"{";
            for (size_t r = 0; r < n.rows; ++r)
            {
                if (r) out += L";";
                for (size_t c = 0; c < n.cols; ++c)
                {
                    if (c) out += L",";
                    formatStrict(n.at(r, c), out);
                }
            }
            out += L"}";
        }
    }

    ListNode ListNode::scalar(ExcelObj v)
    {
        ListNode n;
        n.value = std::move(v);
        return n;
    }

    ListNode ListNode::array(size_t rows, size_t cols, std::vector<ListNode> items)
    {
        ListNode n;
        n.isArray = true;
        n.rows = rows;
        n.cols = cols;
        n.items = std::move(items);
        return n;
    }

    bool ListNode::hasNested() const
    {
        if (!isArray) return false;
        for (const auto& it : items) if (it.isArray) return true;
        return false;
    }

    ListNode parseListText(const std::wstring& text)
    {
        const std::wstring t = trim(text);
        if (isBraced(t)) return parseBraced(t);
        // concise: `,`로 나눈 N×1 목록(항목 안의 `{…}`는 중첩 배열).
        std::vector<ListNode> items;
        const auto toks = splitTop(t, L',');
        if (toks.size() == 1)
        {
            // 콤마도 중괄호도 없는 단일 텍스트는 분류하지 않고 텍스트 그대로
            // (일반 셀 "1"은 텍스트 "1" — 변환 명령은 항목 하나짜리 숫자/논리
            // 목록을 값 자체로 쓰므로 왕복이 유지된다). 인용 토큰만 푼다.
            const std::wstring one = trim(toks[0]);
            if (isBraced(one)) return parseBraced(one);
            if (auto q = unquote(one))
                items.push_back(ListNode::scalar(ExcelObj(std::wstring_view(*q))));
            else
                items.push_back(ListNode::scalar(ExcelObj(std::wstring_view(one))));
            return ListNode::array(1, 1, std::move(items));
        }
        for (const auto& tok : toks) items.push_back(parseItem(tok));
        const size_t n = items.size();
        return ListNode::array(n, 1, std::move(items));
    }

    ListNode listFromValue(const ExcelObj& v)
    {
        if (v.isType(ExcelType::Multi))
        {
            ExcelArray a(v, /*trim*/ false);
            const size_t R = a.nRows(), C = a.nCols();
            std::vector<ListNode> items;
            items.reserve(R * C);
            for (size_t r = 0; r < R; ++r)
                for (size_t c = 0; c < C; ++c)
                    items.push_back(listFromValue(a.at((ExcelArray::row_t)r, (ExcelArray::col_t)c)));
            return ListNode::array(R, C, std::move(items));
        }
        if (v.isType(ExcelType::Str)) return parseListText(v.toString());
        return ListNode::scalar(ExcelObj(v));
    }

    ExcelObj scalarFromToken(const std::wstring& token)
    {
        if (auto v = classify(token)) return std::move(*v);
        return ExcelObj(std::wstring_view(token));
    }

    std::wstring scalarText(const ExcelObj& v, bool strict)
    {
        switch (v.type())
        {
        case ExcelType::Str:
            return strict ? quoted(v.toString()) : v.toString();
        case ExcelType::Nil:
        case ExcelType::Missing:
            return L"";
        default:
            return v.toString();
        }
    }

    std::wstring formatListText(const ListNode& n)
    {
        if (!n.isArray) return scalarText(n.value, false);
        if (n.cols == 1 && !n.hasNested())
        {
            std::wstring out;
            for (size_t r = 0; r < n.rows; ++r)
            {
                if (r) out += L", ";
                const ExcelObj& v = n.at(r, 0).value;
                if (v.isType(ExcelType::Str))
                {
                    const std::wstring s = v.toString();
                    out += needsQuoteConcise(s) ? quoted(s) : s;
                }
                else out += scalarText(v, false);
            }
            return out;
        }
        std::wstring out;
        formatStrict(n, out);
        return out;
    }

    bool listItemEquals(const ExcelObj& a, const ExcelObj& b)
    {
        auto isText = [](const ExcelObj& o)
        { return o.isType(ExcelType::Str) || o.isType(ExcelType::Nil) || o.isType(ExcelType::Missing); };
        auto isNum = [](const ExcelObj& o)
        { return o.isType(ExcelType::Num) || o.isType(ExcelType::Int); };

        if (isText(a) && isText(b))
        {
            const std::wstring sa = a.isType(ExcelType::Str) ? a.toString() : L"";
            const std::wstring sb = b.isType(ExcelType::Str) ? b.toString() : L"";
            if (sa.size() != sb.size()) return false;
            for (size_t i = 0; i < sa.size(); ++i)
                if (std::towlower(sa[i]) != std::towlower(sb[i])) return false;
            return true;
        }
        if (isNum(a) && isNum(b))
            return a.get<double>() == b.get<double>();
        if (a.isType(ExcelType::Bool) && b.isType(ExcelType::Bool))
            return a.get<bool>() == b.get<bool>();
        return false;
    }
}
