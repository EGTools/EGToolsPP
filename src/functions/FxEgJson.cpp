// FxEgJson.cpp — JSONFILTER (EGToolsVB FX2_JSON.vb 포팅, plan/20 E7).
//   JSON 텍스트를 파싱해 (경로, 값) 2열 표로 평탄화. Key_path("a.b[0].c")를
//   지정하면 해당 토큰 이하만 출력. Newtonsoft 대신 자체 경량 파서 사용.
//   값 타입 개선: 숫자→숫자, true/false→불린, null→빈 문자열 (VB는 전부 문자열).
// 순수 계산(비매크로형).

#include "../core/Registry.h"
#include "../core/ArrayUtil.h"
#include "../core/Spill.h"

#include <xlOil/xlOil.h>

#include <cwctype>
#include <memory>
#include <string>
#include <vector>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        struct JVal
        {
            enum Kind { Obj, Arr, Str, Num, Boolean, Null } kind = Null;
            double num = 0;
            bool b = false;
            std::wstring str;
            std::vector<std::pair<std::wstring, JVal>> obj;
            std::vector<JVal> arr;
        };

        struct Parser
        {
            const std::wstring& s;
            size_t p = 0;
            bool ok = true;

            explicit Parser(const std::wstring& text) : s(text) {}

            void ws() { while (p < s.size() && iswspace(s[p])) ++p; }
            bool eat(wchar_t c)
            {
                ws();
                if (p < s.size() && s[p] == c) { ++p; return true; }
                return false;
            }

            JVal parse()
            {
                JVal v = value();
                ws();
                if (p != s.size()) ok = false;
                return v;
            }

            JVal value()
            {
                ws();
                if (p >= s.size()) { ok = false; return {}; }
                const wchar_t c = s[p];
                if (c == L'{') return object();
                if (c == L'[') return array();
                if (c == L'"') { JVal v; v.kind = JVal::Str; v.str = string(); return v; }
                if (c == L't' || c == L'f')
                {
                    JVal v; v.kind = JVal::Boolean;
                    if (s.compare(p, 4, L"true") == 0) { v.b = true; p += 4; }
                    else if (s.compare(p, 5, L"false") == 0) { v.b = false; p += 5; }
                    else ok = false;
                    return v;
                }
                if (c == L'n')
                {
                    JVal v;
                    if (s.compare(p, 4, L"null") == 0) p += 4; else ok = false;
                    return v;
                }
                // 숫자
                {
                    JVal v; v.kind = JVal::Num;
                    wchar_t* end = nullptr;
                    v.num = wcstod(s.c_str() + p, &end);
                    if (!end || end == s.c_str() + p) { ok = false; return {}; }
                    p = (size_t)(end - s.c_str());
                    return v;
                }
            }

            std::wstring string()
            {
                std::wstring out;
                ++p;   // opening quote
                while (p < s.size() && s[p] != L'"')
                {
                    wchar_t c = s[p++];
                    if (c == L'\\' && p < s.size())
                    {
                        const wchar_t e = s[p++];
                        switch (e)
                        {
                        case L'n': out += L'\n'; break;
                        case L't': out += L'\t'; break;
                        case L'r': out += L'\r'; break;
                        case L'b': out += L'\b'; break;
                        case L'f': out += L'\f'; break;
                        case L'u':
                            if (p + 4 <= s.size())
                            {
                                out += (wchar_t)wcstoul(s.substr(p, 4).c_str(), nullptr, 16);
                                p += 4;
                            }
                            else ok = false;
                            break;
                        default: out += e; break;
                        }
                    }
                    else out += c;
                }
                if (p >= s.size()) ok = false; else ++p;   // closing quote
                return out;
            }

            JVal object()
            {
                JVal v; v.kind = JVal::Obj;
                ++p;   // '{'
                ws();
                if (eat(L'}')) return v;
                while (ok)
                {
                    ws();
                    if (p >= s.size() || s[p] != L'"') { ok = false; break; }
                    std::wstring key = string();
                    if (!eat(L':')) { ok = false; break; }
                    v.obj.emplace_back(std::move(key), value());
                    if (eat(L',')) continue;
                    if (eat(L'}')) break;
                    ok = false;
                }
                return v;
            }

            JVal array()
            {
                JVal v; v.kind = JVal::Arr;
                ++p;   // '['
                ws();
                if (eat(L']')) return v;
                while (ok)
                {
                    v.arr.push_back(value());
                    if (eat(L',')) continue;
                    if (eat(L']')) break;
                    ok = false;
                }
                return v;
            }
        };

        // "a.b[0].c" 경로 탐색. 실패 시 nullptr.
        const JVal* selectPath(const JVal* v, const std::wstring& path)
        {
            size_t i = 0;
            while (i < path.size() && v)
            {
                if (path[i] == L'.') { ++i; continue; }
                if (path[i] == L'[')
                {
                    const size_t e = path.find(L']', i);
                    if (e == std::wstring::npos) return nullptr;
                    const long idx = wcstol(path.substr(i + 1, e - i - 1).c_str(), nullptr, 10);
                    if (v->kind != JVal::Arr || idx < 0 || (size_t)idx >= v->arr.size())
                        return nullptr;
                    v = &v->arr[(size_t)idx];
                    i = e + 1;
                    continue;
                }
                size_t e = i;
                while (e < path.size() && path[e] != L'.' && path[e] != L'[') ++e;
                const std::wstring key = path.substr(i, e - i);
                if (v->kind != JVal::Obj) return nullptr;
                const JVal* next = nullptr;
                for (const auto& [k, child] : v->obj)
                    if (k == key) { next = &child; break; }
                v = next;
                i = e;
            }
            return v;
        }

        void flatten(const JVal& v, const std::wstring& prefix,
                     std::vector<std::pair<std::wstring, ExcelObj>>& out)
        {
            switch (v.kind)
            {
            case JVal::Obj:
                for (const auto& [k, child] : v.obj)
                    flatten(child, prefix.empty() ? k : prefix + L"." + k, out);
                break;
            case JVal::Arr:
                for (size_t i = 0; i < v.arr.size(); ++i)
                    flatten(v.arr[i], prefix + L"[" + std::to_wstring(i) + L"]", out);
                break;
            case JVal::Str:
                out.emplace_back(prefix, ExcelObj(std::wstring_view(v.str)));
                break;
            case JVal::Num:
                out.emplace_back(prefix, ExcelObj(v.num));
                break;
            case JVal::Boolean:
                out.emplace_back(prefix, ExcelObj(v.b));
                break;
            case JVal::Null:
                out.emplace_back(prefix, ExcelObj(std::wstring_view(L"")));
                break;
            }
        }

        ExcelObj* jsonFilter(const ExcelObj& textObj, const ExcelObj& pathObj)
        {
            if (textObj.isMissing()) return returnValue(CellError::Value);
            const std::wstring text = textObj.toString();
            if (text.empty()) return returnValue(CellError::Value);

            Parser parser(text);
            const JVal root = parser.parse();
            if (!parser.ok) return returnValue(CellError::Value);

            const JVal* target = &root;
            if (!pathObj.isMissing())
            {
                const std::wstring path = pathObj.toString();
                if (!path.empty())
                {
                    target = selectPath(&root, path);
                    if (!target) return returnValue(CellError::Value);
                }
            }

            std::vector<std::pair<std::wstring, ExcelObj>> rows;
            flatten(*target, L"", rows);
            if (rows.empty()) return returnValue(CellError::NA);

            std::vector<ExcelObj> vals;
            vals.reserve(rows.size() * 2);
            for (auto& [path, val] : rows)
            {
                vals.push_back(ExcelObj(std::wstring_view(path)));
                vals.push_back(std::move(val));
            }
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)rows.size(), 2, vals));
        }
    }

    void registerEgJson()
    {
        egtools::core::registerFn(L"JSONFILTER",
            [](const ExcelObj& text, const ExcelObj& path) -> ExcelObj*
            {
                try { return jsonFilter(text, path); }
                catch (...) { return returnValue(CellError::Value); }
            });
    }
}
