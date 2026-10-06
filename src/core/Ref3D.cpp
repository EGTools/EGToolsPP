#include "Ref3D.h"
#include "Version.h"

#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelRef.h>
#include <cwctype>
#include <utility>

using namespace xloil;

namespace egtools::core
{
    namespace
    {
        bool eqNoCase(const std::wstring& a, const std::wstring& b)
        {
            if (a.size() != b.size()) return false;
            for (size_t i = 0; i < a.size(); ++i)
                if (towupper(a[i]) != towupper(b[i])) return false;
            return true;
        }

        std::wstring trimWs(const std::wstring& s)
        {
            const size_t a = s.find_first_not_of(L" \t\r\n");
            if (a == std::wstring::npos) return L"";
            return s.substr(a, s.find_last_not_of(L" \t\r\n") - a + 1);
        }

        bool isIdentChar(wchar_t c) { return iswalnum(c) || c == L'_' || c == L'.'; }

        // Top-level argument texts of every `fname(...)` call in `f`, in order.
        // Skips "strings" and 'quoted sheet names'; respects () {} [] nesting.
        std::vector<std::vector<std::wstring>> callArgs(const std::wstring& f, const std::wstring& fname)
        {
            std::vector<std::vector<std::wstring>> calls;
            auto skipQuoted = [&](size_t i) -> size_t   // i at opening quote → index past closing
            {
                const wchar_t q = f[i];
                for (++i; i < f.size(); ++i)
                    if (f[i] == q) { if (i + 1 < f.size() && f[i + 1] == q) ++i; else return i + 1; }
                return f.size();
            };
            for (size_t i = 0; i < f.size(); )
            {
                const wchar_t c = f[i];
                if (c == L'"' || c == L'\'') { i = skipQuoted(i); continue; }
                const size_t open = i + fname.size();
                if (open < f.size() && f[open] == L'(' && (i == 0 || !isIdentChar(f[i - 1]))
                    && eqNoCase(f.substr(i, fname.size()), fname))
                {
                    std::vector<std::wstring> args;
                    std::wstring cur;
                    int depth = 0;
                    for (size_t j = open + 1; j < f.size(); ++j)
                    {
                        const wchar_t d = f[j];
                        if (d == L'"' || d == L'\'')
                        {
                            const size_t e = skipQuoted(j);
                            cur.append(f, j, e - j); j = e - 1; continue;
                        }
                        if (d == L'(' || d == L'{' || d == L'[') ++depth;
                        else if (d == L')' || d == L'}' || d == L']')
                        {
                            if (depth == 0) break;   // closing paren of this call
                            --depth;
                        }
                        else if (d == L',' && depth == 0) { args.push_back(trimWs(cur)); cur.clear(); continue; }
                        cur += d;
                    }
                    args.push_back(trimWs(cur));
                    calls.push_back(std::move(args));
                    i = open + 1;   // nested calls of the same function are scanned too
                    continue;
                }
                ++i;
            }
            return calls;
        }

        // "[Book]Sheet1:Sheet3!A1:B2" / "'Sheet 1:Sheet 3'!A1" → parts.
        bool parse3D(const std::wstring& arg, std::wstring& book, std::wstring& s1,
                     std::wstring& s2, std::wstring& cells)
        {
            const size_t bang = arg.find_last_of(L'!');
            if (bang == std::wstring::npos || bang == 0) return false;
            std::wstring sheets = arg.substr(0, bang);
            cells = trimWs(arg.substr(bang + 1));
            if (cells.empty()) return false;
            if (sheets.size() >= 2 && sheets.front() == L'\'' && sheets.back() == L'\'')
            {
                std::wstring u;
                for (size_t i = 1; i + 1 < sheets.size(); ++i)
                {
                    u += sheets[i];
                    if (sheets[i] == L'\'' && sheets[i + 1] == L'\'') ++i;   // '' → '
                }
                sheets = u;
            }
            book.clear();
            const size_t lb = sheets.find(L'['), rb = sheets.find(L']');
            if (lb != std::wstring::npos && rb != std::wstring::npos && lb < rb)
            {
                book = sheets.substr(lb + 1, rb - lb - 1);
                sheets = sheets.substr(rb + 1);
            }
            const size_t colon = sheets.find(L':');
            if (colon == std::wstring::npos || sheets.find(L':', colon + 1) != std::wstring::npos)
                return false;
            s1 = sheets.substr(0, colon);
            s2 = sheets.substr(colon + 1);
            return !s1.empty() && !s2.empty();
        }

        // Each sheet's range value of a 3D reference text, in workbook order.
        // Empty on failure.
        std::vector<ExcelObj> read3D(const std::wstring& arg, const ExcelObj& caller)
        {
            std::wstring book, s1, s2, cells;
            if (!parse3D(arg, book, s1, s2, cells)) return {};
            if (book.empty())
            {
                auto nm = tryCallExcel(msxll::xlSheetNm, caller);   // "[Book]Sheet"
                if (nm.second != 0) return {};
                const std::wstring full = nm.first.toString();
                const size_t rb = full.find(L']');
                if (full.empty() || full[0] != L'[' || rb == std::wstring::npos) return {};
                book = full.substr(1, rb - 1);
            }
            auto list = tryCallExcel(msxll::xlfGetWorkbook, 1, book);   // ordered "[Book]Sheet"
            if (list.second != 0 || !list.first.isType(ExcelType::Multi)) return {};
            ExcelArray names(list.first);
            std::vector<std::wstring> sheets;
            for (size_t k = 0; k < (size_t)names.nRows() * names.nCols(); ++k)
            {
                const std::wstring full = names.at(k).toString();
                const size_t rb = full.find(L']');
                sheets.push_back(rb == std::wstring::npos ? full : full.substr(rb + 1));
            }
            size_t i1 = sheets.size(), i2 = sheets.size();
            for (size_t k = 0; k < sheets.size(); ++k)
            {
                if (i1 == sheets.size() && eqNoCase(sheets[k], s1)) i1 = k;
                if (i2 == sheets.size() && eqNoCase(sheets[k], s2)) i2 = k;
            }
            if (i1 == sheets.size() || i2 == sheets.size()) return {};
            if (i1 > i2) std::swap(i1, i2);

            std::vector<ExcelObj> out;
            for (size_t k = i1; k <= i2; ++k)
            {
                try { out.push_back(ExcelRef(L"[" + book + L"]" + sheets[k] + L"!" + cells).value()); }
                catch (...) { return {}; }
            }
            return out;
        }
    }

    Ref3D::Ref3D(const wchar_t* bareName, size_t nArgs)
        : _name(registeredName(bareName)), _n(nArgs)
    {}

    std::vector<ExcelObj> Ref3D::sheets(size_t i, const ExcelObj& arg)
    {
        if (!(arg.isType(ExcelType::Err) && arg.val.err == (int)CellError::Value)) return {};
        if (!_tried)
        {
            _tried = true;
            auto c = tryCallExcel(msxll::xlfCaller);
            if (c.second == 0 && c.first.isType(ExcelType::RangeRef))
            {
                _caller = c.first;
                auto f = tryCallExcel(msxll::xlfGetCell, 41, _caller);
                if (f.second == 0)
                {
                    auto calls = callArgs(f.first.toString(), _name);
                    for (auto& call : calls)   // prefer the call whose arity matches
                        if (call.size() == _n) { _texts = std::move(call); break; }
                    if (_texts.empty() && !calls.empty()) _texts = std::move(calls[0]);
                }
            }
        }
        if (i >= _texts.size()) return {};
        return read3D(_texts[i], _caller);
    }

    size_t passedArgs(const ExcelObj* const* args, size_t n)
    {
        while (n > 0 && (!args[n - 1] || args[n - 1]->isMissing())) --n;
        return n;
    }
}
