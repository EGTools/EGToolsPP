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
//     (SUM / _xleta.SUM → "SUM"), Direction B removes them (SUM→function ref) —
//     but ONLY for the native 16-aggregator set (core::isGroupByAggregator);
//     lambdas and unknown names leave the whole call untouched. Trailing
//     optional args are fine: argument order is native-parity since plan/21 R1.
//
// The conversion-target set is core::shadowedFunctionNames() (single source of
// truth with registration). Formula rewriting is identifier-aware and skips
// string literals (adapted from FxLet.cpp splitTopLevel / replaceName), so bare
// names never match partially or inside text.
//
// Storage-token artifacts beyond the function NAME (both fixed 2026-08-03):
//   * `_xlpm.` — LET/LAMBDA parameter-name storage prefix. It surfaces in COM
//     formula text whenever the surrounding call is unresolved on the host
//     (e.g. `xLET(_xlpm.x, ...)` on a modern host, `_xlfn.LET(_xlpm.x, ...)` on
//     a legacy host). Excel REFUSES a Formula/Formula2 put containing `_xlpm.`
//     (0x800A03EC) — Direction B used to rewrite `xLET(_xlpm.x,…)` to
//     `LET(_xlpm.x,…)` whose write failed, leaving the cell unconverted. Both
//     directions now strip the prefix from every identifier (stripXlpmPrefix),
//     counted in the summary as `_xlpm.*`.
//   * `@` — implicit-intersection marker. The bulk scanner reads `.Formula`
//     (legacy view, no `@`), but the per-cell fallback reads xlOil formula() =
//     `.Formula2` on dynamic-array hosts, which SHOWS `@` in front of calls
//     from legacy-authored cells (e.g. `=TAKE(@xSORT(…),3)`). Rewriting that
//     to `=TAKE(@SORT(…),3)` "succeeds" but `@SORT` collapses the array to a
//     single value. Direction B now drops `@` directly in front of any name it
//     converts to native (stripImplicitAt), so both scan paths converge on the
//     spilled-native semantics the button promises. `@` elsewhere (cell refs,
//     non-target functions) is left untouched. Known limit: a modern `@SORT(…)`
//     down-saved by a legacy host as `_xlfn.SINGLE(_xlfn._xlws.SORT(…))` keeps
//     its `_xlfn.SINGLE(` wrapper (call-unwrapping, not token replacement).
//   * `_xll.` / `_xludf.` — XLL / unknown-UDF storage prefixes (fixed
//     2026-08-03). Excel records these on any call to an add-in UDF that is
//     not registered on the opening host — DEFINED NAMES keep them (a 2016-
//     saved name `xSORT(...)` reads back as `_xll.xSORT(...)` / `_xludf.
//     TAKE(...)` via RefersTo on MS365). The token rules match `xF`/`EG.F`/
//     bare `F` with a non-identifier left boundary, and the prefix's '.' IS an
//     identifier char — so prefixed calls never matched and names were never
//     converted. Both directions now strip the prefix FIRST, but only when the
//     identifier that follows is one of our own tokens (stripUdfStoragePrefix;
//     foreign add-ins' UDFs stay untouched). Counted as `_xll.*`.
//
// Native-availability safety net (Direction B, added 2026-08-03): funcTable()'s
// year metadata says when Microsoft INTRODUCED a function, but staged rollouts
// mean a table-eligible host may still lack the native name (IMPORTTEXT/
// IMPORTCSV are 365-preview-only as of 2026-08). Functions flagged
// core::isStagedRollout() are probed once per session via Application.
// Evaluate("F()") and convert to native ONLY on a positive result (anything
// but #NAME?/failure); unconfirmed ones get the IMAGE (Class E) treatment —
// bare F re-points to EG.F so legacy-authored cells keep calculating, EG.F
// cells stay as they are. ONLY staged functions are probed: a first version
// probed EVERY native and Evaluate("F()") proved context-sensitive (measured
// 2026-08-03: "SORT()" → FALSE / #VALUE! / #NAME? depending on host state —
// SORT/FILTER also collide with XLM command names), so a long-GA function
// could be misjudged as missing, silently disabling its conversion (regression:
// xSORT survived restore). GA functions therefore trust the table, full stop.
//
// CSE arrays: a modern spill down-converts to a multi-cell CSE array formula on
// legacy hosts, which rejects per-cell Formula writes ("cannot change part of an
// array" — swallowed inside the COM put, so it LOOKS converted but isn't). Such
// regions are detected via HasArray/CurrentArray, visited once per region, and
// re-entered whole-region (FormulaArray on legacy hosts, native spill from the
// anchor on dynamic-array hosts) — see setArrayFormula().
//
// Performance (measured on tests/sample_200k.xlsx, 20k formula cells — scan
// 7.9s → 0.63s): the workbook is scanned exactly ONCE (collectTargets builds a
// work list of address → rewritten formula) and the apply pass resolves just
// those addresses. Within the scan, the load-bearing pieces are:
//   * bulk area reads — each SpecialCells(Formulas) AREA's .Formula arrives as
//     one 2-D variant array (~50ms for 20k cells) instead of per-cell reads;
//   * the queuedRects skip — once a CSE region is queued, its remaining ~10⁴
//     member cells are rejected in memory, not via 4-5 COM probes each;
//   * TransformPlan — token rules built once per run, and an upper-cased
//     find() prescreen so replaceToken's char-walk only runs on formulas that
//     can actually contain the token.
//
// COM: typed xlOil AppObjects (thisApp / Workbook / Worksheet / ExcelRange) are
// safe here — Ribbon callbacks run on Excel's main thread, unlike the deferred
// COM_API context that forced late-bound calls in FxImage. Defined-names access
// and Workbook.SaveCopyAs (backup) use minimal late-bound IDispatch.

#include "Convert.h"
#include "../core/Version.h"
#include "../core/Aggregate.h"

#include <xlOil/AppObjects.h>
#include <xlOil/ExcelUI.h>      // statusBarMsg
#include <xlOil/ExcelObj.h>

#include <windows.h>
#include <oleauto.h>
#include <shlobj.h>
#include <stdexcept>
#include <string>
#include <vector>
#include <map>
#include <set>
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

        // Strip the LET/LAMBDA parameter storage prefix from every identifier:
        // `_xlpm.x` → `x`. Excel refuses formula writes containing `_xlpm.`
        // (0x800A03EC), so any rewritten formula must shed it. Skips string
        // literals; requires a word boundary on the left and a real identifier
        // start after the prefix. Returns the number of prefixes removed.
        int stripXlpmPrefix(std::wstring& f)
        {
            const wchar_t kPfx[] = L"_xlpm.";
            const size_t  nPfx   = wcslen(kPfx);
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

                if (c == L'_' && i + nPfx < f.size() &&
                    _wcsnicmp(f.c_str() + i, kPfx, nPfx) == 0 &&
                    ((i == 0) || !isIdentChar(f[i - 1])) &&
                    isIdentStart(f[i + nPfx]))
                { i += nPfx; ++n; continue; }   // drop the prefix, keep the name

                out += c; ++i;
            }
            if (n) f.swap(out);
            return n;
        }

        // Strip the XLL / unknown-UDF storage prefixes `_xll.` / `_xludf.` when
        // the identifier that follows is one of OUR conversion tokens (bare F,
        // xF, EG.F — `srcUpper`). Excel prepends these to any add-in UDF call
        // it cannot resolve on the opening host, and defined names keep them in
        // RefersTo (e.g. a 2016-saved name reads back as `_xll.xSORT(...)` on
        // MS365) — the '.' inside the prefix defeats the token rules' left-
        // boundary check, so the prefix must go before they run. Prefixed calls
        // to foreign add-ins' UDFs are left untouched. Returns removals.
        int stripUdfStoragePrefix(std::wstring& f, const std::set<std::wstring>& srcUpper)
        {
            static const wchar_t* kPfx[] = { L"_xll.", L"_xludf." };
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

                if (c == L'_' && ((i == 0) || !isIdentChar(f[i - 1])))
                {
                    size_t adv = 0;
                    for (const auto* p : kPfx)
                    {
                        const size_t np = wcslen(p);
                        if (i + np < f.size() &&
                            _wcsnicmp(f.c_str() + i, p, np) == 0 &&
                            isIdentStart(f[i + np]))
                        { adv = np; break; }
                    }
                    if (adv)
                    {
                        size_t j = i + adv;
                        while (j < f.size() && isIdentChar(f[j])) ++j;
                        if (srcUpper.count(upperOf(f.substr(i + adv, j - i - adv))))
                        { i += adv; ++n; continue; }   // drop the prefix, keep the call
                    }
                }
                out += c; ++i;
            }
            if (n) f.swap(out);
            return n;
        }

        // Direction B: drop an implicit-intersection `@` that sits directly in
        // front of a converted-to-native function name (`@SORT(` → `SORT(`).
        // The `@` is a legacy-load artifact on those calls (the per-cell
        // fallback scanner reads Formula2, which shows it); keeping it would
        // collapse the native function's array result to a single value.
        // `@` in front of anything else is preserved. Returns removals.
        int stripImplicitAt(std::wstring& f, const std::vector<std::wstring>& namesUpper)
        {
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

                if (c == L'@')
                {
                    size_t j = i + 1;
                    while (j < f.size() && isIdentChar(f[j])) ++j;
                    if (j > i + 1)
                    {
                        const std::wstring ident = upperOf(f.substr(i + 1, j - i - 1));
                        if (std::find(namesUpper.begin(), namesUpper.end(), ident)
                            != namesUpper.end())
                        { ++n; ++i; continue; }   // skip the '@'
                    }
                }
                out += c; ++i;
            }
            if (n) f.swap(out);
            return n;
        }

        // ── classification ───────────────────────────────────────────────────
        bool isGroupName(const std::wstring& bareUpper)
        { return bareUpper == L"GROUPBY" || bareUpper == L"PIVOTBY"; }
        bool isImageName(const std::wstring& bareUpper)
        { return bareUpper == L"IMAGE"; }

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
            const size_t aggIdx  = (bare == L"GROUPBY") ? 2 : 3;   // 3rd / 4th arg
            const size_t maxArgs = (bare == L"GROUPBY") ? 8 : 11;  // native signature
            // Trailing optional args are position-compatible (native-parity order,
            // plan/21 R1) — only reject when the aggregator slot is missing or
            // there are more args than the native signature allows (malformed).
            if (args.size() < aggIdx + 1 || args.size() > maxArgs) return false;

            std::wstring a = trim(args[aggIdx]);
            const bool quoted = a.size() >= 2 && a.front() == L'"' && a.back() == L'"';
            std::wstring bareAgg = trim(quoted ? a.substr(1, a.size() - 2) : a);
            // A modern-authored eta-lambda aggregator surfaces on a legacy host
            // as `_xleta.SUM` — strip the prefix before the whitelist check.
            const wchar_t kEta[] = L"_XLETA.";
            const size_t  nEta   = wcslen(kEta);
            const bool hadEta = upperOf(bareAgg).compare(0, nEta, kEta) == 0;
            if (hadEta) bareAgg = bareAgg.substr(nEta);
            if (!core::isGroupByAggregator(upperOf(bareAgg)))
                return false;   // lambda / parameterised extension / unsupported

            std::wstring newAgg = toCompat
                ? ((quoted && !hadEta) ? a : (L"\"" + bareAgg + L"\""))  // ensure quoted (text)
                : ((quoted || hadEta) ? bareAgg : a);                    // ensure bare (function ref)

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
        // `nativeUpper` gates Direction B: a group name absent from it (native
        // missing on this host per nativeResolves) leaves the call untouched.
        std::wstring convertGroupCalls(const std::wstring& f, bool toCompat,
                                       const std::vector<std::wstring>& nativeUpper,
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

                    const bool convertible = toCompat ||
                        std::find(nativeUpper.begin(), nativeUpper.end(), bare)
                            != nativeUpper.end();
                    if (isGroupName(bare) && convertible && k < f.size() && f[k] == L'(')
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
        // The (from → to) token rules are host- and direction-dependent but
        // formula-independent, so they are built ONCE per conversion run
        // (TransformPlan) instead of re-deriving storedName/needsXPrefix/
        // hasNativeFunction for every one of possibly 10⁵ formulas. Each rule
        // carries its upper-cased needle: transform() upper-cases the formula
        // once and runs a plain find() prescreen, so the expensive
        // character-by-character replaceToken pass only runs for formulas that
        // can actually contain the token (case-insensitive superset — a false
        // positive just wastes one pass).
        struct TokenRepl { std::wstring from, to, fromUpper; };
        struct TokenRule
        {
            std::wstring F;                 // bare name (counts key, upper)
            std::vector<TokenRepl> repl;
            bool retokenBare = false;       // Direction B: bare native F → forceReset
        };
        struct TransformPlan
        {
            bool toCompat = true;
            // Direction B on a native-IMAGE host: run the source-aware IMAGE
            // pass (convertImageCalls) instead of a blanket IMAGE→EG.IMAGE
            // token rule.
            bool imageAware = false;
            std::vector<TokenRule> rules;
            // Direction B: upper-cased bare names that end up NATIVE after the
            // conversion — an `@` directly in front of one of these is a
            // legacy-load artifact and gets removed (stripImplicitAt). Also
            // gates convertGroupCalls' to-native rewrite (GROUPBY/PIVOTBY).
            std::vector<std::wstring> nativeUpper;
            // Upper-cased identifiers that may legitimately follow an `_xll.`/
            // `_xludf.` storage prefix (bare F, xF, EG.F) — the strip set for
            // stripUdfStoragePrefix. Both directions.
            std::set<std::wstring> udfSrcUpper;
        };

        // Runtime probe, defined after the COM helpers below: is the NATIVE
        // function POSITIVELY confirmed to resolve on this host? Called only
        // for core::isStagedRollout() functions — see the file-top note.
        bool nativeConfirmed(const std::wstring& bareF);

        // Source-aware IMAGE restore (Direction B), defined after the COM
        // helpers below. `evalCtx` is a Worksheet (cells) or Application
        // (defined names) IDispatch whose Evaluate resolves references.
        // Sets `forceReset` when a call STAYS native: the text may be
        // unchanged, but the cell must still be re-entered so the legacy-load
        // implicit-intersection `@` (visible only in the Formula2 view) is
        // cleared — same re-tokenise idea as TokenRule::retokenBare.
        std::wstring convertImageCalls(const std::wstring& f, IDispatch* evalCtx,
                                       std::map<std::wstring, int>& counts,
                                       bool& forceReset);

        // Direction B: convert `F` to native? GA functions trust the metadata
        // table; staged-rollout ones additionally need the positive probe.
        bool convertsToNative(const std::wstring& F)
        {
            return core::hasNativeFunction(F) &&
                   (!core::isStagedRollout(F) || nativeConfirmed(F));
        }

        TransformPlan buildPlan(bool toCompat)
        {
            TransformPlan plan;
            plan.toCompat = toCompat;
            for (const auto& F : core::shadowedFunctionNames())
            {
                plan.udfSrcUpper.insert(F);
                plan.udfSrcUpper.insert(L"EG." + F);
                if (core::needsXPrefix(F)) plan.udfSrcUpper.insert(L"X" + F);

                if (isGroupName(F))                         // convertGroupCalls
                {
                    if (!toCompat && convertsToNative(F))
                        plan.nativeUpper.push_back(F);
                    continue;
                }

                TokenRule r{ F, {}, false };
                auto add = [&r](std::wstring from, std::wstring to)
                {
                    std::wstring up = upperOf(from);
                    r.repl.push_back({ std::move(from), std::move(to), std::move(up) });
                };

                if (isImageName(F))
                {
                    if (toCompat)
                    {
                        add(core::storedName(F), F);        // _xlfn.IMAGE → IMAGE
                        add(L"EG." + F, F);                 // EG.IMAGE   → IMAGE
                    }
                    else if (core::hasNativeFunction(F))
                        plan.imageAware = true;             // convertImageCalls
                }
                else
                {
                    // Class N. Keyword-conflict names (SORT/FILTER/LET) register
                    // as x<F> on legacy hosts — direction A must target xF (bare
                    // F is refused/hijacked by modern-binary hosts with older
                    // licenses), and direction B must accept xF as a source.
                    const bool xPfx = core::needsXPrefix(F);
                    const std::wstring compatName = xPfx ? (L"x" + F) : F;
                    if (toCompat)
                    {
                        add(core::storedName(F), compatName);  // _xlfn[._xlws].F → F/xF
                        add(L"EG." + F, compatName);           // EG.F            → F/xF
                        if (xPfx) add(F, compatName);          // old-doc bare F  → xF
                    }
                    else if (core::hasNativeFunction(F))
                    {
                        if (convertsToNative(F))
                        {
                            add(L"EG." + F, F);                // EG.F → F (native)
                            if (xPfx) add(L"x" + F, F);        // xF   → F (native)
                            r.retokenBare = true;              // bare F → re-tokenise
                            plan.nativeUpper.push_back(F);     // `@F` → `F` (strip)
                        }
                        else
                            // Staged rollout without a positive native probe
                            // (IMPORTTEXT/IMPORTCSV): keep the EGTools
                            // implementation. Legacy bare-F cells re-point to
                            // the registered EG.F (Class E treatment); EG.F
                            // cells stay put.
                            add(F, L"EG." + F);                // bare F → EG.F
                    }
                }
                if (!r.repl.empty() || r.retokenBare)
                    plan.rules.push_back(std::move(r));
            }
            return plan;
        }

        // Returns the rewritten formula and increments `counts`. `forceReset` is
        // set when a Direction-B formula still holds a bare native Class-N name
        // that should be re-tokenised even without a text change.
        std::wstring transform(const std::wstring& formula, const TransformPlan& plan,
                               std::map<std::wstring, int>& counts, bool& forceReset,
                               IDispatch* evalCtx = nullptr)
        {
            const std::wstring fu = upperOf(formula);       // prescreen needle-space
            std::wstring f = formula;

            // XLL / unknown-UDF storage prefixes (`_xll.xSORT`, `_xludf.TAKE`):
            // recorded by Excel on hosts where the UDF wasn't registered —
            // defined names keep them in RefersTo. Strip FIRST so the token
            // rules below see the clean names (the prefix's '.' defeats their
            // left-boundary check). Both directions; foreign UDFs untouched.
            if (fu.find(L"_XLL.") != std::wstring::npos ||
                fu.find(L"_XLUDF.") != std::wstring::npos)
            {
                const int n = stripUdfStoragePrefix(f, plan.udfSrcUpper);
                if (n) counts[L"_xll.*"] += n;
            }

            if (fu.find(L"GROUPBY") != std::wstring::npos ||
                fu.find(L"PIVOTBY") != std::wstring::npos)
                f = convertGroupCalls(f, plan.toCompat, plan.nativeUpper, counts);

            // Class E IMAGE, Direction B — source-aware: https sources go
            // native, everything else (and anything unprovable) → EG.IMAGE.
            if (plan.imageAware && fu.find(L"IMAGE") != std::wstring::npos)
                f = convertImageCalls(f, evalCtx, counts, forceReset);

            for (const auto& rule : plan.rules)
            {
                if (rule.retokenBare && fu.find(rule.F) != std::wstring::npos)
                {
                    const int bare = countToken(f, rule.F);  // legacy bare-F (native)
                    if (bare > 0) { counts[rule.F] += bare; forceReset = true; }
                }
                for (const auto& rp : rule.repl)
                    if (fu.find(rp.fromUpper) != std::wstring::npos)
                        counts[rule.F] += replaceToken(f, rp.from, rp.to);
            }

            // LET/LAMBDA parameter storage prefix (`_xlpm.x` → `x`): Excel
            // rejects formula writes containing `_xlpm.`, so it must go in
            // BOTH directions (compat xLET keeps working with the bare name;
            // native LET re-binds it as its own parameter).
            if (fu.find(L"_XLPM.") != std::wstring::npos)
            {
                const int n = stripXlpmPrefix(f);
                if (n) counts[L"_xlpm.*"] += n;
            }

            // Direction B: `@` in front of a now-native name would collapse its
            // array result to one value — drop it (legacy-load artifact).
            if (!plan.toCompat && !plan.nativeUpper.empty() &&
                f.find(L'@') != std::wstring::npos)
            {
                const int n = stripImplicitAt(f, plan.nativeUpper);
                if (n) counts[L"@(암시적 교차)"] += n;
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

        bool getBoolProp(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return false;
            VARIANT o; VariantInit(&o);
            bool v = SUCCEEDED(VariantChangeType(&o, &r, 0, VT_BOOL)) && o.boolVal != VARIANT_FALSE;
            VariantClear(&o); VariantClear(&r);
            return v;
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

        bool putBStr(IDispatch* d, const wchar_t* name, const std::wstring& val)
        {
            VARIANT a; VariantInit(&a); a.vt = VT_BSTR; a.bstrVal = SysAllocString(val.c_str());
            bool ok = invokeRaw(d, name, DISPATCH_PROPERTYPUT, nullptr, &a, 1);
            VariantClear(&a);
            return ok;
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

        // POSITIVE confirmation that the native function resolves on this host.
        // Called ONLY for core::isStagedRollout() functions (IMPORTTEXT/
        // IMPORTCSV): Application.Evaluate("F()") returns the #NAME? error
        // value (xlErrName 2029) when the name is unknown, and some OTHER
        // value/error (measured: #VALUE! 2015) when the function exists but
        // the args are missing. Only that other-result case confirms; #NAME?,
        // a failed Invoke, or any exception leaves the function unconfirmed —
        // i.e. it KEEPS its EGTools implementation, the safe direction for a
        // staged function. NEVER probe GA functions with this: Evaluate is
        // context-sensitive (see file-top note) and a false negative would
        // silently disable their conversion. Cached per session; runs on the
        // main thread (ribbon callback), where COM is safe.
        bool nativeConfirmed(const std::wstring& bareF)
        {
            static std::map<std::wstring, bool> cache;
            const auto it = cache.find(bareF);
            if (it != cache.end()) return it->second;

            bool confirmed = false;
            try
            {
                if (IDispatch* app = asDisp(thisApp()))
                {
                    Releaser ar{ app };
                    VARIANT a; VariantInit(&a); a.vt = VT_BSTR;
                    a.bstrVal = SysAllocString((bareF + L"()").c_str());
                    VARIANT r; VariantInit(&r);
                    if (invokeRaw(app, L"Evaluate", DISPATCH_METHOD, &r, &a, 1))
                    {
                        confirmed = !(r.vt == VT_ERROR &&
                                      (r.scode & 0xFFFF) == 2029);  // not #NAME?
                        VariantClear(&r);
                    }
                    VariantClear(&a);
                }
            }
            catch (...) {}
            cache[bareF] = confirmed;
            return confirmed;
        }

        // ── Class E — IMAGE source-aware restore (Direction B) ───────────────
        // Native IMAGE renders ONLY https URLs; EG.IMAGE also renders local
        // paths (and http) — see FxImage.cpp. Restoring therefore keeps
        // EG.IMAGE unless the SOURCE is https (사용자 결정 2026-08-03):
        //   * "https://…" string literal            → stay native (no rewrite)
        //   * other literal (local path, http://)   → EG.IMAGE
        //   * reference/expression                  → EVALUATE it now and use
        //     the value as the criterion; https string → native, else EG.IMAGE.
        // The traced decision is a snapshot: if the referenced cell later
        // changes to a local path, the native IMAGE shows #VALUE! — accepted;
        // any evaluation failure falls back to EG.IMAGE, which always works.

        // `a` is a trimmed argument. True if it is a quoted string literal.
        bool isStringArg(const std::wstring& a)
        { return a.size() >= 2 && a.front() == L'"' && a.back() == L'"'; }

        bool httpsLiteral(const std::wstring& a)
        {
            if (!isStringArg(a)) return false;
            std::wstring s;                       // unquote ("" → ")
            for (size_t i = 1; i + 1 < a.size(); ++i)
            {
                s += a[i];
                if (a[i] == L'"' && i + 2 < a.size() && a[i + 1] == L'"') ++i;
            }
            return _wcsnicmp(s.c_str(), L"https://", 8) == 0;
        }

        bool evaluatesToHttps(const std::wstring& expr, IDispatch* evalCtx)
        {
            if (!evalCtx || expr.empty()) return false;
            bool https = false;
            try
            {
                // `(expr)&""` forces a scalar TEXT result (a bare reference
                // would otherwise come back as a Range object); errors stay
                // errors and simply fail the https check.
                const std::wstring probe = L"(" + expr + L")&\"\"";
                VARIANT a; VariantInit(&a); a.vt = VT_BSTR;
                a.bstrVal = SysAllocString(probe.c_str());
                VARIANT r; VariantInit(&r);
                if (invokeRaw(evalCtx, L"Evaluate", DISPATCH_METHOD, &r, &a, 1))
                {
                    if (r.vt == VT_BSTR && r.bstrVal)
                        https = _wcsnicmp(r.bstrVal, L"https://", 8) == 0;
                    VariantClear(&r);
                }
                VariantClear(&a);
            }
            catch (...) {}
            return https;
        }

        std::wstring convertImageCalls(const std::wstring& f, IDispatch* evalCtx,
                                       std::map<std::wstring, int>& counts,
                                       bool& forceReset)
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

                    size_t k = j;
                    while (k < f.size() && iswspace(f[k])) ++k;

                    // Bare IMAGE calls only — EG.IMAGE / _xlfn.IMAGE idents
                    // contain '.' and do not match.
                    if (upperOf(ident) == L"IMAGE" && k < f.size() && f[k] == L'(')
                    {
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
                            const std::wstring inside = f.substr(k + 1, close - k - 1);
                            std::vector<std::wstring> args = splitTopLevel(inside);
                            const std::wstring a0 = args.empty() ? L"" : trim(args[0]);
                            const bool stayNative = isStringArg(a0)
                                ? httpsLiteral(a0)
                                : evaluatesToHttps(a0, evalCtx);
                            // A directly-preceding `@` is a legacy-load
                            // implicit-intersection artifact — drop it with
                            // the conversion (either way the call never needs
                            // it: native IMAGE spills, EG.IMAGE is scalar).
                            if (!out.empty() && out.back() == L'@')
                            { out.pop_back(); counts[L"@(암시적 교차)"]++; }
                            if (stayNative)
                            {
                                out += ident;                      // native IMAGE
                                // Text may be unchanged, but re-enter the cell
                                // anyway: the Formula2 view can still hold a
                                // hidden legacy `@` that only a rewrite clears.
                                forceReset = true;
                            }
                            else { out += L"EG.IMAGE"; counts[L"IMAGE"]++; }
                            out += f.substr(k, close + 1 - k);     // "(args)" as-is
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

        // ── CSE array-region re-entry ────────────────────────────────────────
        // A multi-cell CSE array (how a modern spill surfaces on a legacy host)
        // refuses per-cell Formula writes ("cannot change part of an array"), so
        // the rewritten formula must go back over the WHOLE region. Ordering is
        // chosen so a failure never wipes the region: whole-region FormulaArray
        // first (replacing an entire array is legal without clearing), clear+
        // retry second, plain anchor entry last (the legacy spill engine in
        // core/Spill re-expands it to CSE on recalc). Throws when nothing stuck
        // so the caller does not count the location as converted.
        void setArrayFormula(IDispatch* region, const std::wstring& nf)
        {
            VARIANT one2[2];
            for (auto& v : one2) { VariantInit(&v); v.vt = VT_I4; v.lVal = 1; }

            if (core::supportsDynamicArrays())
            {
                // Native-spill host: dissolve the CSE, let the anchor spill.
                invokeRaw(region, L"ClearContents", DISPATCH_METHOD, nullptr, nullptr, 0);
                IDispatch* anchor = getObject(region, L"Item", one2, 2);
                if (!anchor) throw std::runtime_error("array anchor");
                Releaser ar{ anchor };
                // Formula2 avoids implicit-intersection '@' insertion; hosts
                // without it fall through to the classic Formula property.
                if (!putBStr(anchor, L"Formula2", nf) && !putBStr(anchor, L"Formula", nf))
                    throw std::runtime_error("anchor formula");
                return;
            }

            if (putBStr(region, L"FormulaArray", nf)) return;
            invokeRaw(region, L"ClearContents", DISPATCH_METHOD, nullptr, nullptr, 0);
            if (putBStr(region, L"FormulaArray", nf)) return;
            // FormulaArray rejects >255-char formulas — plain anchor entry.
            IDispatch* anchor = getObject(region, L"Item", one2, 2);
            if (!anchor) throw std::runtime_error("array anchor");
            Releaser ar{ anchor };
            if (!putBStr(anchor, L"Formula", nf))
                throw std::runtime_error("anchor formula");
        }

        // ── scan → work list → apply ─────────────────────────────────────────
        // The workbook is scanned ONCE: every formula is transformed in memory
        // and only the changed locations are recorded (address + rewritten
        // formula). The apply pass then touches just those targets — no second
        // full-sheet iteration, and the HasArray/CurrentArray COM probes run
        // only for cells whose formula actually changes.
        struct Target
        {
            enum class Kind { Cell, Array, Name };
            Kind kind;
            std::wstring sheet;     // Cell / Array
            std::wstring address;   // local address on `sheet` ($D$1 or $D$1:$E$4)
            long nameIndex;         // Name — workbook Names.Item index (logging)
            std::wstring formula;   // rewritten formula to write
            // Name only — captured at SCAN time so the name can be re-created
            // from scratch at apply time (see the delete-and-re-add note in
            // applyTargets): scope-qualified name ("Sheet1!x" for sheet scope),
            // visibility and comment.
            std::wstring nameText;
            std::wstring nameComment;
            bool nameVisible = true;
        };

        // Decide cell-vs-array-region for a CHANGED cell and fetch its address.
        // CSE array members refuse per-cell writes ("cannot change part of an
        // array" is silently swallowed in the COM put), so those queue the WHOLE
        // CurrentArray region — once (seenArrays dedupe; every member cell
        // reports the same formula). Returns false when already queued.
        bool resolveTarget(IDispatch* cd, const std::wstring& wsName,
                           std::set<std::wstring>& seenArrays,
                           bool& isArray, std::wstring& addr)
        {
            isArray = false; addr.clear();
            if (!cd) return false;
            if (getBoolProp(cd, L"HasArray"))
            {
                if (IDispatch* region = getObject(cd, L"CurrentArray"))
                {
                    Releaser rr{ region };
                    isArray = true;
                    addr = getBStr(region, L"Address");
                    return !addr.empty() &&
                           seenArrays.insert(wsName + L"!" + addr).second;
                }
            }
            addr = getBStr(cd, L"Address");
            return !addr.empty();
        }

        // Parse a single-rectangle A1 address ("$D$2:$E$5002" or "$D$2") into
        // 1-based row/col bounds. Returns false on anything unexpected.
        struct Rect { long r1, c1, r2, c2; };
        bool parseA1Rect(const std::wstring& addr, Rect& out)
        {
            long r[2] = { 0, 0 }, c[2] = { 0, 0 };
            int part = 0;
            for (size_t i = 0; i < addr.size(); ++i)
            {
                wchar_t ch = addr[i];
                if (ch == L'$') continue;
                if (ch == L':') { if (++part > 1) return false; continue; }
                if (ch >= L'A' && ch <= L'Z') c[part] = c[part] * 26 + (ch - L'A' + 1);
                else if (ch >= L'a' && ch <= L'z') c[part] = c[part] * 26 + (ch - L'a' + 1);
                else if (ch >= L'0' && ch <= L'9') r[part] = r[part] * 10 + (ch - L'0');
                else return false;   // sheet-qualified / R1C1 / union — bail
            }
            if (r[0] <= 0 || c[0] <= 0) return false;
            if (part == 0) { r[1] = r[0]; c[1] = c[0]; }
            else if (r[1] <= 0 || c[1] <= 0) return false;
            out = { r[0], c[0], r[1], c[1] };
            return true;
        }

        // Fallback scanner: per-cell iteration over SpecialCells(Formulas).
        // One COM round-trip per formula cell — kept as the fallback when the
        // bulk Areas path is unavailable, and as the benchmark baseline.
        void scanSheetPerCell(ExcelWorksheet& ws, const std::wstring& wsName,
                              const TransformPlan& plan, std::map<std::wstring, int>& counts,
                              std::vector<Target>& targets)
        {
            std::set<std::wstring> seenArrays;      // regions never span sheets
            IDispatch* wsEval = asDisp(ws);         // IMAGE source tracing
            Releaser we_{ wsEval };
            ExcelRange formulaCells = ws.usedRange().specialCells(SpecialCells::Formulas);
            // No formula cells (COM error 0x800A03EC) does NOT throw:
            // xlOil swallows it and returns a NULL ExcelRange. In Release
            // AppObject::com() skips its null check (TCheck=false), so
            // begin() below would be a raw null dereference — an SEH
            // access violation the catch(...) cannot stop (Excel crash).
            if (!formulaCells.valid()) return;
            // ComIterator has operator== but no operator!= (and we build as
            // C++17), so drive the loop with an explicit == comparison.
            for (auto it = formulaCells.begin(); !(it == formulaCells.end()); ++it)
            {
                ExcelRange cell = *it;
                std::wstring fml;
                try { fml = cell.formula().toString(); } catch (...) { continue; }
                if (fml.empty()) continue;

                bool fr = false;
                std::map<std::wstring, int> local;
                std::wstring nf = transform(fml, plan, local, fr, wsEval);
                if (nf == fml && !fr) continue;             // not a target

                IDispatch* cd = asDisp(cell);
                Releaser cr{ cd };
                bool isArray = false; std::wstring addr;
                if (!resolveTarget(cd, wsName, seenArrays, isArray, addr)) continue;

                for (auto& kv : local) counts[kv.first] += kv.second;
                targets.push_back({ isArray ? Target::Kind::Array : Target::Kind::Cell,
                                    wsName, addr, 0, nf });
            }
        }

        // Bulk scanner: read each SpecialCells(Formulas) AREA's .Formula as one
        // 2-D variant array — one COM round-trip per contiguous area instead of
        // one per cell. Per-cell COM (the HasArray probe in resolveTarget) then
        // runs only for cells whose formula actually changes. Returns false to
        // make the caller fall back to scanSheetPerCell; counts/targets are only
        // written on success (all-or-nothing so the fallback cannot double-queue).
        bool scanSheetBulk(ExcelWorksheet& ws, const std::wstring& wsName,
                           const TransformPlan& plan, std::map<std::wstring, int>& counts,
                           std::vector<Target>& targets)
        {
            ExcelRange formulaCells = ws.usedRange().specialCells(SpecialCells::Formulas);
            if (!formulaCells.valid()) return true;         // no formulas — done

            IDispatch* fc = asDisp(formulaCells);
            if (!fc) return false;
            Releaser fr_{ fc };
            IDispatch* areas = getObject(fc, L"Areas");
            if (!areas) return false;
            Releaser ar_{ areas };
            const long nAreas = getLong(areas, L"Count");
            if (nAreas <= 0) return true;

            IDispatch* wsDisp = asDisp(ws);
            if (!wsDisp) return false;
            Releaser wr_{ wsDisp };
            IDispatch* cells = getObject(wsDisp, L"Cells");
            if (!cells) return false;
            Releaser cr_{ cells };

            std::set<std::wstring> seenArrays;
            std::map<std::wstring, int> localCounts;
            std::vector<Target> localTargets;
            // Rectangles of CSE regions already queued: every member cell of a
            // multi-cell array WILL transform to the same change, so once the
            // region is queued the remaining ~10⁴ member cells are skipped here
            // in memory — without this, each of them costs 4-5 COM probes in
            // resolveTarget just to be rejected by the seenArrays dedupe.
            std::vector<Rect> queuedRects;

            // Transform one formula; on change, probe the live cell for its
            // CSE-region/cell address and queue the target.
            auto handleOne = [&](long row, long col, const std::wstring& fml)
            {
                if (fml.empty() || fml[0] != L'=') return;
                for (const auto& q : queuedRects)
                    if (row >= q.r1 && row <= q.r2 && col >= q.c1 && col <= q.c2)
                        return;                             // inside a queued CSE region
                bool frs = false;
                std::map<std::wstring, int> local;
                std::wstring nf = transform(fml, plan, local, frs, wsDisp);
                if (nf == fml && !frs) return;              // not a target

                VARIANT rc[2];
                for (auto& v : rc) VariantInit(&v);
                rc[0].vt = VT_I4; rc[0].lVal = row;
                rc[1].vt = VT_I4; rc[1].lVal = col;
                IDispatch* cd = getObject(cells, L"Item", rc, 2);
                Releaser cdr{ cd };
                bool isArray = false; std::wstring addr;
                if (!resolveTarget(cd, wsName, seenArrays, isArray, addr)) return;
                if (isArray)
                {
                    Rect q;
                    if (parseA1Rect(addr, q)) queuedRects.push_back(q);
                }

                for (auto& kv : local) localCounts[kv.first] += kv.second;
                localTargets.push_back({ isArray ? Target::Kind::Array : Target::Kind::Cell,
                                         wsName, addr, 0, nf });
            };

            for (long a = 1; a <= nAreas; ++a)
            {
                VARIANT ix; VariantInit(&ix); ix.vt = VT_I4; ix.lVal = a;
                IDispatch* area = getObject(areas, L"Item", &ix, 1);
                VariantClear(&ix);
                if (!area) return false;
                Releaser al_{ area };
                const long row0 = getLong(area, L"Row");
                const long col0 = getLong(area, L"Column");
                if (row0 <= 0 || col0 <= 0) return false;

                VARIANT fv; VariantInit(&fv);
                if (!invokeRaw(area, L"Formula", DISPATCH_PROPERTYGET, &fv, nullptr, 0))
                    return false;

                if (fv.vt == VT_BSTR)                        // 1×1 area → scalar
                    handleOne(row0, col0, fv.bstrVal ? fv.bstrVal : L"");
                else if (fv.vt == (VT_ARRAY | VT_VARIANT) && fv.parray &&
                         SafeArrayGetDim(fv.parray) == 2)
                {
                    SAFEARRAY* sa = fv.parray;
                    LONG rlb = 0, rub = -1, clb = 0, cub = -1;
                    SafeArrayGetLBound(sa, 1, &rlb); SafeArrayGetUBound(sa, 1, &rub);
                    SafeArrayGetLBound(sa, 2, &clb); SafeArrayGetUBound(sa, 2, &cub);
                    VARIANT* data = nullptr;
                    if (SUCCEEDED(SafeArrayAccessData(sa, (void**)&data)))
                    {
                        const LONG nR = rub - rlb + 1, nC = cub - clb + 1;
                        // SAFEARRAY is column-major: element (i,j) = data[j*nR + i].
                        for (LONG j = 0; j < nC; ++j)
                            for (LONG i = 0; i < nR; ++i)
                            {
                                const VARIANT& v = data[(size_t)j * nR + i];
                                if (v.vt == VT_BSTR && v.bstrVal)
                                    handleOne(row0 + i, col0 + j, v.bstrVal);
                            }
                        SafeArrayUnaccessData(sa);
                    }
                }
                VariantClear(&fv);
            }

            for (auto& kv : localCounts) counts[kv.first] += kv.second;
            targets.insert(targets.end(), localTargets.begin(), localTargets.end());
            return true;
        }

        void collectTargets(ExcelWorkbook& wb, bool toCompat,
                            std::map<std::wstring, int>& counts,
                            std::vector<Target>& targets, bool bulkScan = true)
        {
            const TransformPlan plan = buildPlan(toCompat);

            // Cell formulas — bulk area scan first, per-cell iteration fallback.
            for (auto ws : wb.worksheets().list())
            {
                std::wstring wsName;
                try { wsName = ws.name(); } catch (...) {}
                try { statusBarMsg(L"EGTools++: " + wsName); } catch (...) {}

                bool done = false;
                if (bulkScan)
                {
                    try { done = scanSheetBulk(ws, wsName, plan, counts, targets); }
                    catch (...) { done = false; }
                }
                if (!done)
                {
                    try { scanSheetPerCell(ws, wsName, plan, counts, targets); }
                    catch (...) {}
                }
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
                try { refers = getBStr(nm, L"RefersTo"); } catch (...) {}
                if (refers.empty()) continue;

                bool fr = false;
                std::map<std::wstring, int> local;
                // Application.Evaluate as the IMAGE-tracing context — a name
                // has no home sheet; unqualified refs resolve against the
                // active sheet and any failure just keeps EG.IMAGE.
                IDispatch* appEval = asDisp(thisApp());
                Releaser ae{ appEval };
                std::wstring nf = transform(refers, plan, local, fr, appEval);
                if (nf == refers && !fr) continue;
                for (auto& kv : local) counts[kv.first] += kv.second;
                Target t{ Target::Kind::Name, L"", L"", ix, nf };
                t.nameText    = getBStr(nm, L"Name");        // scope-qualified
                t.nameComment = getBStr(nm, L"Comment");
                t.nameVisible = getBoolProp(nm, L"Visible");
                if (!t.nameText.empty()) targets.push_back(std::move(t));
            }
        }

        size_t applyTargets(ExcelWorkbook& wb, const std::vector<Target>& targets)
        {
            size_t applied = 0;
            IDispatch* wbDisp = asDisp(wb);
            if (!wbDisp) return 0;
            Releaser wbr{ wbDisp };
            IDispatch* sheets = getObject(wbDisp, L"Worksheets");
            Releaser sr{ sheets };
            IDispatch* names = getObject(wbDisp, L"Names");
            Releaser nr{ names };

            // Targets are grouped by sheet in scan order — cache the current one.
            IDispatch* ws = nullptr;
            std::wstring wsName;
            auto releaseWs = [&] { if (ws) { ws->Release(); ws = nullptr; } };

            for (const auto& t : targets)
            {
                if (t.kind == Target::Kind::Name) continue;    // done above
                try
                {
                    if (!sheets) continue;
                    if (!ws || wsName != t.sheet)
                    {
                        releaseWs();
                        VARIANT nv; VariantInit(&nv); nv.vt = VT_BSTR;
                        nv.bstrVal = SysAllocString(t.sheet.c_str());
                        ws = getObject(sheets, L"Item", &nv, 1);
                        VariantClear(&nv);
                        wsName = t.sheet;
                    }
                    if (!ws) continue;

                    VARIANT av; VariantInit(&av); av.vt = VT_BSTR;
                    av.bstrVal = SysAllocString(t.address.c_str());
                    IDispatch* rng = getObject(ws, L"Range", &av, 1);
                    VariantClear(&av);
                    if (!rng) continue;
                    Releaser rr{ rng };

                    if (t.kind == Target::Kind::Array)
                    {
                        try { setArrayFormula(rng, t.formula); ++applied; } catch (...) {}
                    }
                    else
                    {
                        // Same host rule as xlOil setFormula(ExcelObj): Formula2
                        // on dynamic-array hosts (no implicit-intersection '@'),
                        // classic Formula elsewhere (Formula2 put fails there).
                        const bool ok = core::supportsDynamicArrays()
                            ? (putBStr(rng, L"Formula2", t.formula) ||
                               putBStr(rng, L"Formula", t.formula))
                            : putBStr(rng, L"Formula", t.formula);
                        if (ok) ++applied;
                    }
                }
                catch (...) {}
            }
            releaseWs();

            // NAMES LAST — delete and re-create from the scan-time snapshot
            // (사용자 결정 2026-08-03). A defined name that calls an
            // unregistered add-in UDF (e.g. `xSORT` on a modern host) DEGRADES
            // in memory to "=#NAME?" during the run — measured: it is already
            // degraded by apply time even before any cell write (the summary-
            // dialog idle recalc / SaveCopyAs backup triggers it) — and a
            // degraded name refuses both RefersTo writes and Names.Add
            // overwrite. So: try the plain RefersTo put first (healthy names,
            // keeps every property), else DELETE the broken entry and re-Add
            // it from the captured name/formula/visibility/comment.
            for (const auto& t : targets)
            {
                if (t.kind != Target::Kind::Name) continue;
                try
                {
                    if (!names) continue;
                    VARIANT key; VariantInit(&key); key.vt = VT_BSTR;
                    key.bstrVal = SysAllocString(t.nameText.c_str());
                    IDispatch* nm = getObject(names, L"Item", &key, 1);
                    VariantClear(&key);

                    bool ok = false;
                    if (nm)
                    {
                        Releaser mr{ nm };
                        ok = putBStr(nm, L"RefersTo", t.formula);
                        if (!ok)
                            invokeRaw(nm, L"Delete", DISPATCH_METHOD, nullptr, nullptr, 0);
                    }
                    if (!ok)
                    {
                        VARIANT av[3];
                        for (auto& v : av) VariantInit(&v);
                        av[0].vt = VT_BSTR; av[0].bstrVal = SysAllocString(t.nameText.c_str());
                        av[1].vt = VT_BSTR; av[1].bstrVal = SysAllocString(t.formula.c_str());
                        av[2].vt = VT_BOOL; av[2].boolVal = t.nameVisible ? VARIANT_TRUE : VARIANT_FALSE;
                        VARIANT r; VariantInit(&r);
                        ok = invokeRaw(names, L"Add", DISPATCH_METHOD, &r, av, 3);
                        if (ok && !t.nameComment.empty() && r.vt == VT_DISPATCH && r.pdispVal)
                            putBStr(r.pdispVal, L"Comment", t.nameComment);
                        VariantClear(&r);
                        for (auto& v : av) VariantClear(&v);
                    }
                    if (ok) ++applied;
                }
                catch (...) {}
            }
            return applied;
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
                : L"EGTools 함수명을 네이티브 내장 함수로 되돌립니다"
                  L"(IMAGE와 이 Excel에 아직 없는 함수는 EG.* 유지).\n";
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

            // PASS 1 — scan once (read-only): transform in memory and build the
            // work list of changed locations with their rewritten formulas.
            std::map<std::wstring, int> counts;
            std::vector<Target> targets;
            try { collectTargets(wb, toCompat, counts, targets); }
            catch (...) {}
            try { statusBarMsg(L""); } catch (...) {}

            const size_t locations = targets.size();
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

            // PASS 2 — apply the work list only (no re-scan).
            size_t applied = 0;
            {
                Application& app = thisApp();
                PauseExcel pause(app);
                try { applied = applyTargets(wb, targets); }
                catch (...) {}
            }
            try { statusBarMsg(L""); } catch (...) {}
            // Dirty-only Calculate(): the formula writes above already dirtied
            // the converted cells (and their dependents), and leaving Manual
            // mode in ~PauseExcel recalculates them for auto-calc users — this
            // call only covers users who keep Manual calculation. CalculateFull
            // here re-computed EVERY formula in ALL open workbooks (~12s on
            // sample_200k vs 0.03s dirty-only) — the apply phase felt slow.
            try { thisApp().calculate(); } catch (...) {}

            std::wstring done = std::to_wstring(applied)
                + L"곳을 변환했습니다.\n\n백업 파일:\n" + backupPath
                + L"\n\n원복하려면 반대편 버튼을 실행하거나 위 백업 파일을 여세요.";
            MessageBoxW(hwnd, done.c_str(), title, MB_OK | MB_ICONINFORMATION);
        }
    }

    void convertToCompat() { runConversion(/*toCompat*/ true); }
    void convertToNative() { runConversion(/*toCompat*/ false); }
}
