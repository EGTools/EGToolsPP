// Convert.cpp — compatibility function-name conversion engine.
//
// Implements the 3-class conversion model (approved plan / plan/14):
//   Class N (name re-tokenise): every shadowed function except IMAGE/GROUPBY/
//     PIVOTBY. Direction A: `_xlfn[._xlws].F` / `EG.F` → bare F. Direction B:
//     `EG.F` → bare F (native on a capable host) + force re-tokenise of legacy
//     bare-F cells.
//   Class E (IMAGE): never converted to native. Direction A: `_xlfn.IMAGE` /
//     `EG.IMAGE` → bare IMAGE. Direction B: bare IMAGE → EG.IMAGE (keep EGTools);
//     genuine native `_xlfn.IMAGE` is left alone.
//   Class G (GROUPBY/PIVOTBY): the function NAME goes to bare (both directions),
//     and the aggregator argument toggles quotes — Direction A adds them
//     ("SUM"→text), Direction B removes them (SUM→function ref) — but ONLY for the
//     supported aggregator whitelist and ONLY when no trailing optional args make
//     the positions diverge; otherwise the whole call is left untouched.
//
// The conversion-target set is core::shadowedFunctionNames() (single source of
// truth with registration). Formula rewriting is identifier-aware and skips
// string literals (adapted from FxLet.cpp splitTopLevel / replaceName), so bare
// names never match partially or inside text.
//
// COM: typed xlOil AppObjects (thisApp / Workbook / Worksheet / ExcelRange) are
// safe here — Ribbon callbacks run on Excel's main thread, unlike the deferred
// COM_API context that forced late-bound calls in FxImage. Defined-names access
// and Workbook.SaveCopyAs (backup) use minimal late-bound IDispatch.

#include "Convert.h"
#include "../core/Version.h"

#include <xlOil/AppObjects.h>
#include <xlOil/ExcelUI.h>      // statusBarMsg
#include <xlOil/ExcelObj.h>

#include <windows.h>
#include <oleauto.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <map>
#include <set>
#include <functional>
#include <algorithm>
#include <cwctype>

#pragma comment(lib, "shell32.lib")

using namespace xloil;

// Internal xlOil COM connector (static XLL does not auto-connect — see FxLet.cpp).
namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::ribbon
{
    namespace core = egtools::core;

    namespace
    {
        // ── string helpers (adapted from FxLet.cpp) ──────────────────────────
        std::wstring upperOf(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)std::towupper(c);
            return s;
        }

        bool isIdentChar(wchar_t c) { return iswalnum(c) || c == L'_' || c == L'.'; }
        bool isIdentStart(wchar_t c) { return iswalpha(c) || c == L'_'; }

        std::wstring trim(const std::wstring& s)
        {
            size_t a = s.find_first_not_of(L" \t\r\n");
            if (a == std::wstring::npos) return L"";
            size_t b = s.find_last_not_of(L" \t\r\n");
            return s.substr(a, b - a + 1);
        }

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
                case L'"': inStr = true;  cur += c; break;
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

        // Whole-identifier replace of `from`→`to`, skipping string literals and
        // partial matches. Returns the number of replacements. `from` may contain
        // '.' (e.g. "_xlfn.XLOOKUP", "EG.IMAGE"); it is treated as one token whose
        // neighbours must be non-identifier characters.
        int replaceToken(std::wstring& f, const std::wstring& from, const std::wstring& to)
        {
            if (from.empty()) return 0;
            std::wstring out; out.reserve(f.size());
            bool inStr = false; int n = 0;
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

                if (c == from[0] && i + from.size() <= f.size() &&
                    f.compare(i, from.size(), from) == 0)
                {
                    bool leftOk = (i == 0) || !isIdentChar(f[i - 1]);
                    size_t end = i + from.size();
                    bool rightOk = (end >= f.size()) || !isIdentChar(f[end]);
                    if (leftOk && rightOk) { out += to; i = end; ++n; continue; }
                }
                out += c; ++i;
            }
            if (n) f.swap(out);
            return n;
        }

        // Count whole-identifier occurrences of `tok` (skips string literals).
        int countToken(const std::wstring& f, const std::wstring& tok)
        {
            if (tok.empty()) return 0;
            bool inStr = false; int n = 0;
            for (size_t i = 0; i < f.size(); )
            {
                wchar_t c = f[i];
                if (inStr)
                {
                    if (c == L'"') { if (i + 1 < f.size() && f[i + 1] == L'"') { i += 2; continue; } inStr = false; }
                    ++i; continue;
                }
                if (c == L'"') { inStr = true; ++i; continue; }
                if (c == tok[0] && i + tok.size() <= f.size() && f.compare(i, tok.size(), tok) == 0)
                {
                    bool leftOk = (i == 0) || !isIdentChar(f[i - 1]);
                    size_t end = i + tok.size();
                    bool rightOk = (end >= f.size()) || !isIdentChar(f[end]);
                    if (leftOk && rightOk) { ++n; i = end; continue; }
                }
                ++i;
            }
            return n;
        }

        // ── classification ───────────────────────────────────────────────────
        bool isGroupName(const std::wstring& bareUpper)
        { return bareUpper == L"GROUPBY" || bareUpper == L"PIVOTBY"; }
        bool isImageName(const std::wstring& bareUpper)
        { return bareUpper == L"IMAGE"; }

        const std::set<std::wstring>& aggWhitelist()
        {   // Mirror of FxGroup.cpp aggregate() supported text aggregators.
            static const std::set<std::wstring> w = {
                L"SUM", L"AVERAGE", L"COUNT", L"COUNTA",
                L"MAX", L"MIN", L"PRODUCT", L"PERCENTOF" };
            return w;
        }

        std::wstring stripKnownPrefix(std::wstring upperIdent)
        {
            const wchar_t* pfx[] = { L"_XLFN._XLWS.", L"_XLFN.", L"EG." };
            for (const auto* p : pfx)
            {
                size_t n = wcslen(p);
                if (upperIdent.compare(0, n, p) == 0) return upperIdent.substr(n);
            }
            return upperIdent;
        }

        // ── Class G — GROUPBY/PIVOTBY aggregator quote toggle ────────────────
        // Returns true and fills `rebuiltInside` (args re-joined with the toggled
        // aggregator) when the call is convertible; false to leave it untouched.
        bool tryConvertGroup(const std::wstring& bare, const std::wstring& inside,
                             bool toCompat, std::wstring& rebuiltInside)
        {
            std::vector<std::wstring> args = splitTopLevel(inside);
            const size_t aggIdx = (bare == L"GROUPBY") ? 2 : 3;   // 3rd / 4th arg
            // Leave when malformed OR when trailing optional args are present
            // (e.g. EGTools sort_order vs native field_headers → position diverges).
            if (args.size() != aggIdx + 1) return false;

            std::wstring a = trim(args[aggIdx]);
            const bool quoted = a.size() >= 2 && a.front() == L'"' && a.back() == L'"';
            std::wstring inner = quoted ? a.substr(1, a.size() - 2) : a;
            const std::wstring up = upperOf(trim(inner));
            if (aggWhitelist().find(up) == aggWhitelist().end()) return false;  // lambda / unsupported

            std::wstring newAgg = toCompat
                ? (quoted ? a : (L"\"" + trim(inner) + L"\""))   // ensure quoted (text)
                : (quoted ? trim(inner) : a);                    // ensure bare (function ref)

            std::wstring res;
            for (size_t x = 0; x < args.size(); ++x)
            {
                if (x) res += L",";
                res += (x == aggIdx) ? newAgg : args[x];
            }
            rebuiltInside = res;
            return true;
        }

        // Scan for GROUPBY/PIVOTBY calls and rewrite name (→ bare) + aggregator.
        std::wstring convertGroupCalls(const std::wstring& f, bool toCompat,
                                       std::map<std::wstring, int>& counts)
        {
            std::wstring out; out.reserve(f.size());
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

                if (isIdentStart(c))
                {
                    size_t j = i;
                    while (j < f.size() && isIdentChar(f[j])) ++j;
                    const std::wstring ident = f.substr(i, j - i);
                    const std::wstring bare = stripKnownPrefix(upperOf(ident));

                    size_t k = j;
                    while (k < f.size() && iswspace(f[k])) ++k;

                    if (isGroupName(bare) && k < f.size() && f[k] == L'(')
                    {
                        // matching close paren (respecting nested parens + strings)
                        int depth = 0; bool s2 = false; size_t close = std::wstring::npos;
                        for (size_t p = k; p < f.size(); ++p)
                        {
                            wchar_t d = f[p];
                            if (s2) { if (d == L'"') { if (p + 1 < f.size() && f[p + 1] == L'"') ++p; else s2 = false; } continue; }
                            if (d == L'"') { s2 = true; continue; }
                            if (d == L'(') ++depth;
                            else if (d == L')') { if (--depth == 0) { close = p; break; } }
                        }
                        if (close != std::wstring::npos)
                        {
                            const std::wstring original = f.substr(i, close + 1 - i);
                            const std::wstring inside = f.substr(k + 1, close - k - 1);
                            std::wstring rebuiltInside;
                            if (tryConvertGroup(bare, inside, toCompat, rebuiltInside))
                            {
                                std::wstring newCall = bare + L"(" + rebuiltInside + L")";
                                if (newCall != original) { counts[bare]++; out += newCall; }
                                else out += original;
                            }
                            else out += original;   // leave untouched
                            i = close + 1;
                            continue;
                        }
                    }
                    out += ident; i = j; continue;
                }
                out += c; ++i;
            }
            return out;
        }

        // ── full per-formula transform ───────────────────────────────────────
        // Returns the rewritten formula and increments `counts`. `forceReset` is
        // set when a Direction-B formula still holds a bare native Class-N name
        // that should be re-tokenised even without a text change.
        std::wstring transform(const std::wstring& formula, bool toCompat,
                               std::map<std::wstring, int>& counts, bool& forceReset)
        {
            std::wstring f = convertGroupCalls(formula, toCompat, counts);

            for (const auto& F : core::shadowedFunctionNames())
            {
                if (isGroupName(F)) continue;               // handled above

                if (isImageName(F))
                {
                    if (toCompat)
                    {
                        counts[F] += replaceToken(f, core::storedName(F), F);  // _xlfn.IMAGE → IMAGE
                        counts[F] += replaceToken(f, L"EG." + F, F);            // EG.IMAGE   → IMAGE
                    }
                    else if (core::hasNativeFunction(F))
                    {
                        counts[F] += replaceToken(f, F, L"EG." + F);            // bare IMAGE → EG.IMAGE
                    }
                    continue;
                }

                // Class N. Keyword-conflict names (SORT/FILTER/LET) register as
                // x<F> on legacy hosts — direction A must target xF (bare F is
                // refused/hijacked by modern-binary hosts with older licenses),
                // and direction B must accept xF as a source token.
                const bool xPfx = core::needsXPrefix(F);
                const std::wstring compatName = xPfx ? (L"x" + F) : F;
                if (toCompat)
                {
                    counts[F] += replaceToken(f, core::storedName(F), compatName); // _xlfn[._xlws].F → F/xF
                    counts[F] += replaceToken(f, L"EG." + F, compatName);          // EG.F            → F/xF
                    if (xPfx)
                        counts[F] += replaceToken(f, F, compatName);               // old-doc bare F  → xF
                }
                else if (core::hasNativeFunction(F))
                {
                    const int bare = countToken(f, F);                          // legacy bare-F (native)
                    counts[F] += replaceToken(f, L"EG." + F, F) + bare;         // EG.F → F (native)
                    if (xPfx)
                        counts[F] += replaceToken(f, L"x" + F, F);              // xF   → F (native)
                    if (bare > 0) forceReset = true;                            // force re-tokenise
                }
            }
            return f;
        }

        // ── minimal late-bound IDispatch helpers (backup + defined names) ─────
        DISPID dispidOf(IDispatch* d, const wchar_t* name)
        {
            DISPID id = DISPID_UNKNOWN;
            OLECHAR* n = const_cast<OLECHAR*>(name);
            if (FAILED(d->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id)))
                return DISPID_UNKNOWN;
            return id;
        }

        bool invokeRaw(IDispatch* d, const wchar_t* name, WORD flags,
                       VARIANT* result, VARIANT* args, UINT nArgs)
        {
            DISPID id = dispidOf(d, name);
            if (id == DISPID_UNKNOWN) return false;
            std::vector<VARIANT> rev(nArgs ? nArgs : 1);
            for (UINT i = 0; i < nArgs; ++i) rev[i] = args[nArgs - 1 - i];
            DISPPARAMS dp{ nArgs ? rev.data() : nullptr, nullptr, nArgs, 0 };
            DISPID putId = DISPID_PROPERTYPUT;
            if (flags & DISPATCH_PROPERTYPUT) { dp.rgdispidNamedArgs = &putId; dp.cNamedArgs = 1; }
            return SUCCEEDED(d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                                       &dp, result, nullptr, nullptr));
        }

        IDispatch* getObject(IDispatch* d, const wchar_t* name, VARIANT* args = nullptr, UINT n = 0)
        {
            VARIANT r; VariantInit(&r);
            WORD flags = DISPATCH_PROPERTYGET | (n ? DISPATCH_METHOD : 0);
            if (!invokeRaw(d, name, flags, &r, args, n)) return nullptr;
            IDispatch* obj = (r.vt == VT_DISPATCH) ? r.pdispVal : nullptr;
            if (!obj) VariantClear(&r);
            return obj;
        }

        long getLong(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return 0;
            VARIANT out; VariantInit(&out);
            long v = SUCCEEDED(VariantChangeType(&out, &r, 0, VT_I4)) ? out.lVal : 0;
            VariantClear(&out); VariantClear(&r);
            return v;
        }

        std::wstring getBStr(IDispatch* d, const wchar_t* name, VARIANT* args = nullptr, UINT n = 0)
        {
            VARIANT r; VariantInit(&r);
            WORD flags = DISPATCH_PROPERTYGET | (n ? DISPATCH_METHOD : 0);
            if (!invokeRaw(d, name, flags, &r, args, n)) return L"";
            std::wstring s = (r.vt == VT_BSTR && r.bstrVal) ? r.bstrVal : L"";
            VariantClear(&r);
            return s;
        }

        void putBStr(IDispatch* d, const wchar_t* name, const std::wstring& val)
        {
            VARIANT a; VariantInit(&a); a.vt = VT_BSTR; a.bstrVal = SysAllocString(val.c_str());
            invokeRaw(d, name, DISPATCH_PROPERTYPUT, nullptr, &a, 1);
            VariantClear(&a);
        }

        bool callBstr(IDispatch* d, const wchar_t* name, const std::wstring& val)
        {
            VARIANT a; VariantInit(&a); a.vt = VT_BSTR; a.bstrVal = SysAllocString(val.c_str());
            bool ok = invokeRaw(d, name, DISPATCH_METHOD, nullptr, &a, 1);
            VariantClear(&a);
            return ok;
        }

        struct Releaser { IDispatch* p; ~Releaser() { if (p) p->Release(); } };

        IDispatch* asDisp(const detail::UnknownObject& obj)
        {
            IDispatch* d = nullptr;
            if (obj.ptr()) obj.ptr()->QueryInterface(IID_IDispatch, (void**)&d);
            return d;
        }

        // ── iteration over every formula-bearing location ────────────────────
        using Setter  = std::function<void(const std::wstring&)>;
        using Visitor = std::function<void(const std::wstring&, const Setter&)>;

        void forEachFormula(ExcelWorkbook& wb, const Visitor& visit)
        {
            // Cell formulas — only formula cells via SpecialCells(xlCellTypeFormulas).
            for (auto ws : wb.worksheets().list())
            {
                try
                {
                    ExcelRange formulaCells = ws.usedRange().specialCells(SpecialCells::Formulas);
                    // No formula cells (COM error 0x800A03EC) does NOT throw:
                    // xlOil swallows it and returns a NULL ExcelRange. In Release
                    // AppObject::com() skips its null check (TCheck=false), so
                    // begin() below would be a raw null dereference — an SEH
                    // access violation the catch(...) cannot stop (Excel crash).
                    if (!formulaCells.valid()) continue;
                    try { statusBarMsg(std::wstring(L"EGTools++: ") + ws.name()); } catch (...) {}
                    // ComIterator has operator== but no operator!= (and we build as
                    // C++17), so drive the loop with an explicit == comparison.
                    for (auto it = formulaCells.begin(); !(it == formulaCells.end()); ++it)
                    {
                        ExcelRange cell = *it;
                        std::wstring fml;
                        try { fml = cell.formula().toString(); } catch (...) { continue; }
                        if (fml.empty()) continue;
                        // ExcelObj overload only: the wstring_view overload always
                        // calls PutFormula2, which fails (0x800A03EC) on every
                        // pre-dynamic-array host, silently converting 0 cells.
                        // This one falls back to PutFormula when the host lacks
                        // dynamic arrays (supportsDynamicArrays check in xlOil).
                        visit(fml, [&cell](const std::wstring& nf) { cell.setFormula(ExcelObj(nf)); });
                    }
                }
                catch (...) { continue; }   // sheet has no formula cells / COM error
            }

            // Defined names — RefersTo formula text (late-bound).
            IDispatch* wbDisp = asDisp(wb);
            if (!wbDisp) return;
            Releaser wbr{ wbDisp };
            IDispatch* names = getObject(wbDisp, L"Names");
            if (!names) return;
            Releaser nr{ names };
            const long cnt = getLong(names, L"Count");
            for (long ix = 1; ix <= cnt; ++ix)
            {
                VARIANT idx; VariantInit(&idx); idx.vt = VT_I4; idx.lVal = ix;
                IDispatch* nm = getObject(names, L"Item", &idx, 1);
                VariantClear(&idx);
                if (!nm) continue;
                Releaser mr{ nm };
                std::wstring refers;
                try { refers = getBStr(nm, L"RefersTo"); } catch (...) { continue; }
                if (refers.empty()) continue;
                visit(refers, [nm](const std::wstring& nf) { putBStr(nm, L"RefersTo", nf); });
            }
        }

        // ── backup (SaveCopyAs) ──────────────────────────────────────────────
        std::wstring timestamp()
        {
            SYSTEMTIME t; GetLocalTime(&t);
            wchar_t buf[32];
            swprintf(buf, 32, L"%04d%02d%02d_%02d%02d%02d",
                     t.wYear, t.wMonth, t.wDay, t.wHour, t.wMinute, t.wSecond);
            return buf;
        }

        bool backupWorkbook(ExcelWorkbook& wb, std::wstring& outPath)
        {
            std::wstring path, name;
            try { path = wb.path(); } catch (...) {}
            try { name = wb.name(); } catch (...) {}
            const std::wstring ts = timestamp();

            std::wstring backup;
            if (!path.empty())
            {
                // xlOil ExcelWorkbook::path() may return either the folder or the
                // full path incl. file name — derive the directory robustly.
                std::wstring dir = path;
                if (!name.empty() && dir.size() >= name.size() &&
                    dir.compare(dir.size() - name.size(), name.size(), name) == 0)
                    dir = dir.substr(0, dir.size() - name.size());      // strip file name
                while (!dir.empty() && (dir.back() == L'\\' || dir.back() == L'/'))
                    dir.pop_back();                                     // strip separators

                std::wstring file = name.empty() ? L"Book.xlsx" : name;
                size_t dot = file.find_last_of(L'.');
                std::wstring base = (dot == std::wstring::npos) ? file : file.substr(0, dot);
                std::wstring ext  = (dot == std::wstring::npos) ? L".xlsx" : file.substr(dot);
                backup = dir + L"\\" + base + L"_egbak_" + ts + ext;
            }
            else
            {
                wchar_t docs[MAX_PATH]{};
                if (FAILED(SHGetFolderPathW(nullptr, CSIDL_PERSONAL, nullptr, 0, docs)))
                    return false;
                std::wstring base = name.empty() ? L"Book" : name;
                size_t dot = base.find_last_of(L'.');
                if (dot != std::wstring::npos) base = base.substr(0, dot);
                backup = std::wstring(docs) + L"\\" + base + L"_egbak_" + ts + L".xlsx";
            }

            IDispatch* wbDisp = asDisp(wb);
            if (!wbDisp) return false;
            Releaser r{ wbDisp };
            if (!callBstr(wbDisp, L"SaveCopyAs", backup)) return false;
            outPath = backup;
            return true;
        }

        // ── UI text ──────────────────────────────────────────────────────────
        HWND excelHwnd() { return GetActiveWindow(); }

        std::wstring summaryText(const std::map<std::wstring, int>& counts,
                                 size_t locations, bool toCompat)
        {
            std::wstring s = std::wstring(L"활성 통합문서에서 변환 대상 ")
                + std::to_wstring(locations) + L"곳을 찾았습니다.\n\n";
            for (const auto& kv : counts)
                if (kv.second > 0)
                    s += L"  • " + kv.first + L" : " + std::to_wstring(kv.second) + L"\n";
            s += L"\n";
            s += toCompat
                ? L"모던/네이티브 토큰을 EGTools 호환 함수명으로 바꿉니다.\n"
                : L"EGTools 함수명을 네이티브 내장 함수로 되돌립니다(IMAGE는 EG.IMAGE 유지).\n";
            s += L"변환 직전 백업 파일이 자동 저장됩니다.\n\n계속하시겠습니까?";
            return s;
        }

        // ── driver ───────────────────────────────────────────────────────────
        void runConversion(bool toCompat)
        {
            const wchar_t* title = toCompat
                ? L"EGTools++ — 구버전 호환 적용" : L"EGTools++ — 내장 함수로 복원";
            const HWND hwnd = excelHwnd();

            try { COM::connectCom(); } catch (...) {}

            ExcelWorkbook wb(nullptr);
            try { wb = thisApp().workbooks().active(); }
            catch (...) {}
            // ActiveWorkbook is null WITHOUT an exception when no workbook is
            // open (Excel idle after closing all) — and a null AppObject is a
            // release-mode null dereference on first use, not a catchable error.
            if (!wb.valid())
            {
                MessageBoxW(hwnd, L"열려 있는 통합문서를 찾을 수 없습니다.", title,
                            MB_OK | MB_ICONWARNING);
                return;
            }

            // ── saved-state gate (사용자 결정 2026-07-31) ────────────────────
            // The conversion rewrites formulas in place, so it must start from
            // a saved workbook. Unsaved → ask permission, save, then proceed;
            // declined or cancelled → abort.
            {
                IDispatch* wbDisp = asDisp(wb);
                if (!wbDisp) return;
                Releaser r{ wbDisp };
                std::wstring path;
                try { path = wb.path(); } catch (...) {}
                const bool saved = getLong(wbDisp, L"Saved") != 0;
                if (path.empty() || !saved)
                {
                    if (MessageBoxW(hwnd,
                        L"이 기능은 통합문서를 저장한 상태에서 실행해야 합니다.\n"
                        L"지금 저장하고 계속하시겠습니까?",
                        title, MB_YESNO | MB_ICONQUESTION) != IDYES)
                        return;

                    if (path.empty())
                    {
                        // Never saved (Book1…): Excel's own Save As dialog so
                        // the user picks name/location (xlDialogSaveAs = 5).
                        IDispatch* appDisp = asDisp(thisApp());
                        if (appDisp)
                        {
                            Releaser ar{ appDisp };
                            IDispatch* dialogs = getObject(appDisp, L"Dialogs");
                            if (dialogs)
                            {
                                Releaser dr{ dialogs };
                                VARIANT ix; VariantInit(&ix); ix.vt = VT_I4; ix.lVal = 5;
                                IDispatch* dlg = getObject(dialogs, L"Item", &ix, 1);
                                VariantClear(&ix);
                                if (dlg)
                                {
                                    Releaser gr{ dlg };
                                    invokeRaw(dlg, L"Show", DISPATCH_METHOD,
                                              nullptr, nullptr, 0);
                                }
                            }
                        }
                    }
                    else
                        invokeRaw(wbDisp, L"Save", DISPATCH_METHOD, nullptr, nullptr, 0);

                    // Re-verify — the user may have cancelled the save dialog.
                    std::wstring newPath;
                    try { newPath = wb.path(); } catch (...) {}
                    if (newPath.empty() || getLong(wbDisp, L"Saved") == 0)
                    {
                        MessageBoxW(hwnd,
                            L"통합문서가 저장되지 않아 변환을 취소했습니다.",
                            title, MB_OK | MB_ICONINFORMATION);
                        return;
                    }
                }
            }

            // PASS 1 — scan / count (read-only).
            std::map<std::wstring, int> counts;
            size_t locations = 0;
            try
            {
                forEachFormula(wb, [&](const std::wstring& fml, const Setter&)
                {
                    bool fr = false;
                    std::map<std::wstring, int> local;
                    std::wstring nf = transform(fml, toCompat, local, fr);
                    if (nf != fml || fr)
                    {
                        ++locations;
                        for (auto& kv : local) counts[kv.first] += kv.second;
                    }
                });
            }
            catch (...) {}
            try { statusBarMsg(L""); } catch (...) {}

            if (locations == 0)
            {
                MessageBoxW(hwnd, L"변환할 대상이 없습니다.", title, MB_OK | MB_ICONINFORMATION);
                return;
            }

            if (MessageBoxW(hwnd, summaryText(counts, locations, toCompat).c_str(),
                            title, MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON1) != IDYES)
                return;

            // Backup before any change.
            std::wstring backupPath;
            if (!backupWorkbook(wb, backupPath))
            {
                MessageBoxW(hwnd,
                    L"백업 파일 저장에 실패하여 변환을 중단했습니다.\n"
                    L"통합문서를 먼저 저장한 뒤 다시 시도하세요.",
                    title, MB_OK | MB_ICONERROR);
                return;
            }

            // PASS 2 — apply.
            size_t applied = 0;
            {
                Application& app = thisApp();
                PauseExcel pause(app);
                try
                {
                    forEachFormula(wb, [&](const std::wstring& fml, const Setter& setFormula)
                    {
                        bool fr = false;
                        std::map<std::wstring, int> local;
                        std::wstring nf = transform(fml, toCompat, local, fr);
                        if (nf != fml || fr)
                        {
                            try { setFormula(nf); ++applied; } catch (...) {}
                        }
                    });
                }
                catch (...) {}
            }
            try { statusBarMsg(L""); } catch (...) {}
            try { thisApp().calculate(true, false); } catch (...) {}

            std::wstring done = std::to_wstring(applied)
                + L"곳을 변환했습니다.\n\n백업 파일:\n" + backupPath
                + L"\n\n원복하려면 반대편 버튼을 실행하거나 위 백업 파일을 여세요.";
            MessageBoxW(hwnd, done.c_str(), title, MB_OK | MB_ICONINFORMATION);
        }
    }

    void convertToCompat() { runConversion(/*toCompat*/ true); }
    void convertToNative() { runConversion(/*toCompat*/ false); }
}
