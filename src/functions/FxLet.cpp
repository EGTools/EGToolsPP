// FxLet.cpp — LET (macro-type), ported from EGTools VB xLET.
//
// Native LET binds names in the formula language and cannot be a normal UDF.
// Instead we read the CALLING CELL's own formula text, split the LET arguments,
// substitute each name with its (value) in the final calculation, and evaluate
// the resulting expression. The declared arguments below exist only so Excel/
// IntelliSense show a useful signature — their runtime values are ignored.
//
// Evaluation uses COM Application.Evaluate (NOT xlfEvaluate): the XLM macro
// xlfEvaluate re-enters Excel's calc engine and corrupts the caller context when
// several LET cells recalculate together; COM Evaluate runs outside that engine.
//
// Fallback (per spec): if evaluation fails, return the final substituted formula
// string itself (for debugging / manual inspection) rather than #VALUE!.
//
// Note: names "R"/"C" (and any single R/C) are reserved by Excel's R1C1 parser
// and cannot be used as LET names — Excel rewrites them to row/column references.

#include "../core/Registry.h"
#include "../core/Spill.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/AppObjects.h>   // thisApp()
#include <xlOil/ExcelUI.h>      // variantToExcelObj
#include <windows.h>
#include <oleauto.h>
#include <string>
#include <vector>

using namespace xloil;

// Internal xlOil COM connector (linked from xlOil-COM.lib). A static XLL does not
// auto-connect COM, so thisApp() throws "COM Connection not ready" until we do.
namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::functions
{
    namespace
    {
        constexpr int kXlfCaller  = 89;
        constexpr int kXlfGetCell = 185;

        // Split top-level comma-separated args, respecting () [] {} and "" strings.
        std::vector<std::wstring> splitTopLevel(const std::wstring& s)
        {
            std::vector<std::wstring> out;
            std::wstring cur;
            int paren = 0, brace = 0, brack = 0;
            bool inStr = false;
            for (size_t i = 0; i < s.size(); ++i)
            {
                wchar_t c = s[i];
                if (inStr)
                {
                    cur += c;
                    if (c == L'"')
                    {
                        if (i + 1 < s.size() && s[i + 1] == L'"') cur += s[++i];
                        else inStr = false;
                    }
                    continue;
                }
                switch (c)
                {
                case L'"': inStr = true; cur += c; break;
                case L'(': ++paren; cur += c; break;
                case L')': --paren; cur += c; break;
                case L'{': ++brace; cur += c; break;
                case L'}': --brace; cur += c; break;
                case L'[': ++brack; cur += c; break;
                case L']': --brack; cur += c; break;
                case L',':
                    if (paren == 0 && brace == 0 && brack == 0) { out.push_back(cur); cur.clear(); }
                    else cur += c;
                    break;
                default: cur += c; break;
                }
            }
            out.push_back(cur);
            return out;
        }

        std::wstring trim(const std::wstring& s)
        {
            size_t a = s.find_first_not_of(L" \t\r\n");
            if (a == std::wstring::npos) return L"";
            size_t b = s.find_last_not_of(L" \t\r\n");
            return s.substr(a, b - a + 1);
        }

        bool isIdentChar(wchar_t c) { return iswalnum(c) || c == L'_' || c == L'.'; }

        // Replace whole-word identifier `name` with `(value)` everywhere in `f`,
        // skipping inside string literals and not matching partial identifiers.
        std::wstring replaceName(const std::wstring& f, const std::wstring& name,
                                 const std::wstring& value)
        {
            if (name.empty()) return f;
            const std::wstring rep = L"(" + value + L")";
            std::wstring out;
            bool inStr = false;
            for (size_t i = 0; i < f.size(); )
            {
                wchar_t c = f[i];
                if (inStr)
                {
                    out += c;
                    if (c == L'"')
                    {
                        if (i + 1 < f.size() && f[i + 1] == L'"') { out += f[i + 1]; i += 2; continue; }
                        inStr = false;
                    }
                    ++i; continue;
                }
                if (c == L'"') { inStr = true; out += c; ++i; continue; }

                if (c == name[0] && f.compare(i, name.size(), name) == 0)
                {
                    bool leftOk = (i == 0) || !isIdentChar(f[i - 1]);
                    size_t end = i + name.size();
                    bool rightOk = (end >= f.size()) || !isIdentChar(f[end]);
                    if (leftOk && rightOk) { out += rep; i = end; continue; }
                }
                out += c; ++i;
            }
            return out;
        }

        // Locate the LET call's argument list inside the caller's formula (handles
        // both bare "LET(" and "EG.LET(") and return the text inside the parens.
        bool extractLetInside(const std::wstring& formula, std::wstring& inside)
        {
            std::wstring up = formula;
            for (auto& c : up) c = (wchar_t)towupper(c);
            size_t pos = up.find(L"LET(");
            if (pos == std::wstring::npos) return false;
            size_t open = pos + 3;            // index of '('
            int depth = 0; bool inStr = false;
            for (size_t i = open; i < formula.size(); ++i)
            {
                wchar_t c = formula[i];
                if (inStr) { if (c == L'"') { if (i + 1 < formula.size() && formula[i + 1] == L'"') ++i; else inStr = false; } continue; }
                if (c == L'"') inStr = true;
                else if (c == L'(') ++depth;
                else if (c == L')') { if (--depth == 0) { inside = formula.substr(open + 1, i - open - 1); return true; } }
            }
            return false;
        }

        // Evaluate a formula string via COM Application.Evaluate (late-bound).
        bool comEvaluate(const std::wstring& formula, ExcelObj& out)
        {
            IUnknown* unk = nullptr;
            try { xloil::COM::connectCom(); unk = xloil::thisApp().ptr(); }
            catch (...) { return false; }
            if (!unk) return false;

            IDispatch* disp = nullptr;
            if (FAILED(unk->QueryInterface(IID_IDispatch, (void**)&disp)) || !disp) return false;

            bool ok = false;
            OLECHAR* nm = (OLECHAR*)L"Evaluate";
            DISPID dispid = 0;
            if (SUCCEEDED(disp->GetIDsOfNames(IID_NULL, &nm, 1, LOCALE_USER_DEFAULT, &dispid)))
            {
                VARIANT arg; VariantInit(&arg);
                arg.vt = VT_BSTR; arg.bstrVal = SysAllocString(formula.c_str());
                DISPPARAMS params{ &arg, nullptr, 1, 0 };
                VARIANT result; VariantInit(&result);
                HRESULT hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                          DISPATCH_METHOD, &params, &result, nullptr, nullptr);
                if (SUCCEEDED(hr))
                {
                    out = xloil::variantToExcelObj(result, false);
                    ok = (out.type() != ExcelType::Err);
                }
                VariantClear(&result);
                VariantClear(&arg);
            }
            disp->Release();
            return ok;
        }

        ExcelObj* doLet()
        {
            // Caller cell + its formula text (macro-sheet calls).
            auto [caller, rc1] = tryCallExcel(kXlfCaller);
            if (rc1 != 0) return returnValue(CellError::Value);
            auto [fml, rc2] = tryCallExcel(kXlfGetCell, 41, caller);
            if (rc2 != 0) return returnValue(CellError::Value);

            std::wstring inside;
            if (!extractLetInside(fml.toString(), inside))
                return returnValue(CellError::Value);

            std::vector<std::wstring> args = splitTopLevel(inside);
            for (auto& a : args) a = trim(a);
            if (args.size() < 3 || (args.size() % 2) == 0)   // name,value,...,calc (odd ≥ 3)
                return returnValue(CellError::Value);

            std::wstring calc = args.back();
            // Substitute each name → (value); later pairs first so a value may use
            // an earlier name (one level of dependency, matching the VB version).
            for (int i = (int)args.size() - 2; i >= 1; i -= 2)
                calc = replaceName(calc, args[i - 1], args[i]);

            const std::wstring finalFormula = L"=" + trim(calc);

            ExcelObj result;
            if (!comEvaluate(finalFormula, result))
                return returnValue(ExcelObj(std::wstring_view(finalFormula)));   // fallback

            if (result.type() == ExcelType::Multi)
                return egtools::core::output(std::move(result));
            return returnValue(std::move(result));
        }
    }

    void registerLet()
    {
        // Variadic: name/value pairs + final calculation, up to Excel's 255-arg
        // limit (127 pairs + calculation; see en.json "repeat"). The declared args
        // are placeholders — their runtime values are ignored; doLet() reads the
        // caller's formula text. macro=true so xlfCaller/xlfGetCell work.
        egtools::core::registerRawFn(L"LET",
            [](const FuncInfo&, const ExcelObj**) -> ExcelObj*
            {
                try { return doLet(); }
                catch (...) { return returnValue(CellError::Value); }
            },
            /*macro*/ true);
    }
}
