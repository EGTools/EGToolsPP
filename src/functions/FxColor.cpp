// FxColor.cpp — format/visibility reference functions (plan/22 B3), ported from
// VB FX4_Color + FX1 VISIBLEAGGR/ILOOKUP + FX9 ISVISIBLE.
//
//   DISPLAYCOLOR(color_range, [font], [conditional])              — fill/font color array
//   COUNTIFCOLOR(search_range, color_cell, [font], [conditional]) — count cells matching a cell's color
//   SUMIFCOLOR(search_range, color_cell, [font], [conditional])   — sum cells matching a cell's color
//   ISVISIBLE(range)                               — row/column visibility array
//   VISIBLEAGGR(range, [func])                     — visible cells only, listed or aggregated
//   ILOOKUP(find, lookup_array, image_range, [if_not_found], [match_mode], [search_mode])
//                                                  — copy the picture anchored at the matched
//                                                    cell into the calling cell
//
// All are MACRO-TYPE with reference arguments (RangeArg) and use late-bound COM
// (LateCom.h) on Excel's calc main thread. Inherent limits (documented in the
// manual): changing only a format or hidden state does NOT trigger recalc — F9
// is needed. Colors are the cell's own format by default; conditional=TRUE
// reads DisplayFormat (conditional formatting applied, unreliable when the
// recalc runs while another sheet is active).

// windows.h FIRST with full GDI (BITMAPFILEHEADER for the clipboard DIB save):
// xlOil's WindowsSlim.h pre-includes windows.h with NOGDI otherwise.
#include <windows.h>

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "../core/Apply.h"
#include "../core/LateCom.h"
#include "../core/Aggregate.h"
#include "ImageInsert.h"
#include "RegexMatchMode.h"

#include <xlOil/xlOil.h>
#include <xlOil/Caller.h>
#include <xlOil/ExcelRef.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/AppObjects.h>
#include <xlOil/ExcelTypeLib.h>
#include <algorithm>
#include <cwctype>
#include <optional>
#include <regex>
#include <string>
#include <vector>

using namespace xloil;
using namespace egtools::latecom;

namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::functions
{
    namespace
    {
        constexpr size_t kMaxCells = 100000;   // per-cell COM walk guard

        // Range IDispatch from an XLL reference (via its workbook address).
        IDispatch* rangeDisp(const RangeArg& ref)
        {
            try
            {
                xloil::COM::connectCom();
                xloil::ExcelRange rng(ref.address());
                IDispatch* d = nullptr;
                if (FAILED(rng.com().QueryInterface(IID_IDispatch, (void**)&d))) return nullptr;
                return d;
            }
            catch (...) { return nullptr; }
        }

        // Cells(i,j) of a Range (1-based), late-bound. Caller releases.
        IDispatch* cellAt(IDispatch* range, long i, long j)
        {
            VARIANT a[2]; VariantInit(&a[0]); VariantInit(&a[1]);
            a[0].vt = VT_I4; a[0].lVal = i;
            a[1].vt = VT_I4; a[1].lVal = j;
            IDispatch* c = getObject(range, L"Cells", a, 2);
            VariantClear(&a[0]); VariantClear(&a[1]);
            return c;
        }

        // Sub-range rows r1..r2 × cols c1..c2 (1-based, inclusive) of a Range,
        // late-bound. Caller releases.
        IDispatch* blockAt(IDispatch* range, long r1, long r2, long c1, long c2)
        {
            IDispatch* tl = cellAt(range, r1, c1);
            if (!tl) return nullptr;
            Releaser rt(tl);
            VARIANT a[2]; VariantInit(&a[0]); VariantInit(&a[1]);
            a[0].vt = VT_I4; a[0].lVal = r2 - r1 + 1;
            a[1].vt = VT_I4; a[1].lVal = c2 - c1 + 1;
            return getObject(tl, L"Resize", a, 2);
        }

        // {Interior|Font}.Color of one cell; -1 on failure. display=true reads
        // DisplayFormat (conditional formatting applied) — Excel documents it as
        // unreliable inside UDFs (wrong/failed reads when the volatile recalc runs
        // with another sheet active), so the formatted color is the default.
        double cellColor(IDispatch* cell, bool font, bool display)
        {
            IDispatch* src = cell;
            IDispatch* df = nullptr;
            if (display)
            {
                df = getObject(cell, L"DisplayFormat");
                if (!df) return -1.0;
                src = df;
            }
            Releaser r1(df);
            IDispatch* part = getObject(src, font ? L"Font" : L"Interior");
            if (!part) return -1.0;
            Releaser r2(part);
            const double c = getDouble(part, L"Color", -1.0);
            // A multi-cell range of mixed colors reports Color = 0 (same as black)
            // with ColorIndex = Null — report mixed as -1 so the caller splits.
            if (c == 0.0)
            {
                VARIANT ci; VariantInit(&ci);
                const bool mixed = invokeRaw(part, L"ColorIndex", DISPATCH_PROPERTYGET, &ci, nullptr, 0) &&
                                   ci.vt == VT_NULL;
                VariantClear(&ci);
                if (mixed) return -1.0;
            }
            return c;
        }

        bool rowHidden(IDispatch* range, long i)
        {
            IDispatch* row = getObjectIdx(range, L"Rows", i);
            if (!row) return false;
            Releaser r(row);
            return getBool(row, L"Hidden", false);
        }

        bool colHidden(IDispatch* range, long j)
        {
            IDispatch* col = getObjectIdx(range, L"Columns", j);
            if (!col) return false;
            Releaser r(col);
            return getBool(col, L"Hidden", false);
        }

        // Clip the reference loop to the sheet's used range (VB Intersect intent):
        // returns 1-based inclusive row/col bounds INSIDE the reference, or false
        // when the reference does not overlap the used range.
        bool usedClip(const RangeArg& ref, IDispatch* range,
                      long& i0, long& i1, long& j0, long& j1)
        {
            auto [rwFirst, colFirst, rwLast, colLast] = ref.bounds();   // 0-based
            i0 = 1; i1 = (long)(rwLast - rwFirst + 1);
            j0 = 1; j1 = (long)(colLast - colFirst + 1);

            IDispatch* ws = getObject(range, L"Worksheet");
            if (!ws) return true;
            Releaser rw(ws);
            IDispatch* ur = getObject(ws, L"UsedRange");
            if (!ur) return true;
            Releaser ru(ur);

            const long ur0 = getLong(ur, L"Row", 1);                    // 1-based sheet row
            const long uc0 = getLong(ur, L"Column", 1);
            long urN = 1, ucN = 1;
            if (IDispatch* rr = getObject(ur, L"Rows")) { Releaser t(rr); urN = getLong(rr, L"Count", 1); }
            if (IDispatch* cc = getObject(ur, L"Columns")) { Releaser t(cc); ucN = getLong(cc, L"Count", 1); }

            const long refRow0 = (long)rwFirst + 1, refCol0 = (long)colFirst + 1;
            i0 = std::max(i0, ur0 - refRow0 + 1);
            i1 = std::min(i1, ur0 + urN - 1 - refRow0 + 1);
            j0 = std::max(j0, uc0 - refCol0 + 1);
            j1 = std::min(j1, uc0 + ucN - 1 - refCol0 + 1);
            return i0 <= i1 && j0 <= j1;
        }

        // ---- ILOOKUP matching helpers --------------------------------------
        std::wstring upperW(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)towupper(c);
            return s;
        }

        bool wildMatch(const std::wstring& pat, const std::wstring& s, size_t p = 0, size_t i = 0)
        {
            while (p < pat.size())
            {
                if (pat[p] == L'*')
                {
                    for (size_t k = i; k <= s.size(); ++k)
                        if (wildMatch(pat, s, p + 1, k)) return true;
                    return false;
                }
                if (i >= s.size()) return false;
                if (pat[p] != L'?' && towupper(pat[p]) != towupper(s[i])) return false;
                ++p; ++i;
            }
            return i == s.size();
        }
    }

    void registerColor()
    {
        namespace core = egtools::core;

        // ── DISPLAYCOLOR(color_range, [font], [conditional]) ─────────────────
        core::registerFn(L"DISPLAYCOLOR",
            [](const RangeArg& ref, const ExcelObj& fontA, const ExcelObj& condA) -> ExcelObj*
            {
                const bool font = fontA.get<bool>(false);
                const bool cond = condA.get<bool>(false);
                const auto nR = (long)ref.nRows(), nC = (long)ref.nCols();
                if ((size_t)nR * nC > kMaxCells) return returnValue(CellError::Value);
                IDispatch* range = rangeDisp(ref);
                if (!range) return returnValue(CellError::Value);
                Releaser rr(range);

                // Block bisection (see COUNTIFCOLOR): a block of one color fills
                // its part of the output at once; only mixed blocks are split.
                std::vector<ExcelObj> flat((size_t)nR * nC);
                auto fill = [&](long r1, long r2, long c1, long c2, const ExcelObj& v)
                {
                    for (long i = r1; i <= r2; ++i)
                        for (long j = c1; j <= c2; ++j)
                            flat[(size_t)(i - 1) * nC + (j - 1)] = v;
                };
                auto solve = [&](auto& self, long r1, long r2, long c1, long c2) -> void
                {
                    const bool single = r1 == r2 && c1 == c2;
                    IDispatch* blk = single ? cellAt(range, r1, c1) : blockAt(range, r1, r2, c1, c2);
                    double col = -1.0;
                    if (blk)
                    {
                        Releaser rb(blk);
                        col = cellColor(blk, font, cond);   // -1 = mixed or failed
                    }
                    if (col >= 0) { fill(r1, r2, c1, c2, ExcelObj(col)); return; }
                    if (single) { flat[(size_t)(r1 - 1) * nC + (c1 - 1)] = ExcelObj(CellError::Value); return; }
                    if (r2 - r1 >= c2 - c1)
                    {
                        const long m = r1 + (r2 - r1) / 2;
                        self(self, r1, m, c1, c2); self(self, m + 1, r2, c1, c2);
                    }
                    else
                    {
                        const long m = c1 + (c2 - c1) / 2;
                        self(self, r1, r2, c1, m); self(self, r1, r2, m + 1, c2);
                    }
                };
                solve(solve, 1, nR, 1, nC);
                return core::output(core::makeArray(
                    (ExcelArrayBuilder::row_t)nR, (ExcelArrayBuilder::col_t)nC, flat));
            },
            /*macro*/ true, /*threadsafe*/ false);

        // ── COUNTIFCOLOR / SUMIFCOLOR ────────────────────────────────────────
        auto colorAggr = [](const RangeArg& ref, const RangeArg& colorRef,
                            const ExcelObj& fontA, const ExcelObj& condA, bool doSum) -> ExcelObj*
        {
            const bool font = fontA.get<bool>(false);
            const bool cond = condA.get<bool>(false);
            IDispatch* range = rangeDisp(ref);
            if (!range) return returnValue(CellError::Value);
            Releaser rr(range);
            IDispatch* colorRange = rangeDisp(colorRef);
            if (!colorRange) return returnValue(CellError::Value);
            Releaser rcr(colorRange);

            double target = -1.0;
            {
                IDispatch* c0 = cellAt(colorRange, 1, 1);
                if (!c0) return returnValue(CellError::Value);
                Releaser rc(c0);
                target = cellColor(c0, font, cond);
                if (target < 0) return returnValue(CellError::Value);
            }

            long i0, i1, j0, j1;
            if (!usedClip(ref, range, i0, i1, j0, j1))
                return returnValue(ExcelObj(0.0));
            if ((size_t)(i1 - i0 + 1) * (j1 - j0 + 1) > kMaxCells)
                return returnValue(CellError::Value);

            // A merged area counts once: only its first cell inside the scanned
            // block (normally the top-left) is visited — Excel copies the fill to
            // every cell of the area. Range.MergeCells is FALSE when no cell is
            // merged (TRUE/Null otherwise), which skips the per-cell probe.
            bool anyMerged = true;
            {
                VARIANT m; VariantInit(&m);
                if (invokeRaw(range, L"MergeCells", DISPATCH_PROPERTYGET, &m, nullptr, 0) &&
                    m.vt == VT_BOOL && m.boolVal == VARIANT_FALSE)
                    anyMerged = false;
                VariantClear(&m);
            }
            auto [rwFirst, colFirst, rwLast, colLast] = ref.bounds();   // 0-based
            (void)rwLast; (void)colLast;
            const long minRow = (long)rwFirst + i0, minCol = (long)colFirst + j0;   // 1-based sheet
            auto mergedTail = [&](IDispatch* cell, long i, long j) -> bool
            {
                if (!anyMerged || !getBool(cell, L"MergeCells", false)) return false;
                IDispatch* ma = getObject(cell, L"MergeArea");
                if (!ma) return false;
                Releaser rm(ma);
                const long anchorRow = std::max(getLong(ma, L"Row", 0), minRow);
                const long anchorCol = std::max(getLong(ma, L"Column", 0), minCol);
                return (long)rwFirst + i != anchorRow || (long)colFirst + j != anchorCol;
            };

            // values via the fast XLL coerce; colors via COM
            const ExcelObj values = ref.value();
            ExcelArray varr(values, /*trim*/ false);
            const bool isArr = values.isType(ExcelType::Multi);

            double sum = 0.0; long count = 0;
            std::optional<CellError> err;
            auto take = [&](long i, long j)
            {
                ExcelObj v = isArr
                    ? ExcelObj(varr.at((ExcelArray::row_t)(i - 1), (ExcelArray::col_t)(j - 1)))
                    : ExcelObj(values);
                if (doSum)
                {
                    if (v.isType(ExcelType::Err)) { if (!err) err = v.get<CellError>(); return; }
                    if (v.isType(ExcelType::Num)) sum += v.get<double>(0.0);
                }
                else
                {
                    if (font)   // VB: font-color match counts only non-empty cells
                    {
                        const auto t = v.type();
                        if (t == ExcelType::Nil || t == ExcelType::Missing ||
                            (t == ExcelType::Str && v.stringLength() == 0)) return;
                    }
                    ++count;
                }
            };

            // Block bisection instead of a per-cell walk (several COM round trips
            // per cell made large ranges take seconds, and the function is
            // volatile). A mixed block reports -1 (see cellColor): a block
            // of one non-target color is skipped whole, a block of the target
            // color with no merged cell is taken whole from the value array, and
            // only mixed blocks are split. Single cells keep the merge rule.
            auto solve = [&](auto& self, long r1, long r2, long c1, long c2) -> void
            {
                if (err) return;
                if (r1 == r2 && c1 == c2)
                {
                    IDispatch* cell = cellAt(range, r1, c1);
                    if (!cell) return;
                    Releaser rc(cell);
                    // color first: the merge probe is the costlier one and a
                    // merged area's cells all share its color anyway
                    if (cellColor(cell, font, cond) == target && !mergedTail(cell, r1, c1))
                        take(r1, c1);
                    return;
                }
                if (IDispatch* blk = blockAt(range, r1, r2, c1, c2))
                {
                    Releaser rb(blk);
                    const double col = cellColor(blk, font, cond);   // -1 = mixed
                    if (col >= 0 && col != target) return;
                    if (col == target)
                    {
                        VARIANT m; VariantInit(&m);
                        const bool noMerge =
                            !anyMerged ||
                            (invokeRaw(blk, L"MergeCells", DISPATCH_PROPERTYGET, &m, nullptr, 0) &&
                             m.vt == VT_BOOL && m.boolVal == VARIANT_FALSE);
                        VariantClear(&m);
                        if (noMerge)
                        {
                            for (long i = r1; i <= r2; ++i)
                                for (long j = c1; j <= c2; ++j) take(i, j);
                            return;
                        }
                    }
                }
                if (r2 - r1 >= c2 - c1)
                {
                    const long m = r1 + (r2 - r1) / 2;
                    self(self, r1, m, c1, c2); self(self, m + 1, r2, c1, c2);
                }
                else
                {
                    const long m = c1 + (c2 - c1) / 2;
                    self(self, r1, r2, c1, m); self(self, r1, r2, m + 1, c2);
                }
            };
            solve(solve, i0, i1, j0, j1);
            if (err) return returnValue(*err);
            return returnValue(doSum ? ExcelObj(sum) : ExcelObj((double)count));
        };

        core::registerFn(L"COUNTIFCOLOR",
            [colorAggr](const RangeArg& ref, const RangeArg& colorRef,
                        const ExcelObj& fontA, const ExcelObj& condA) -> ExcelObj*
            { return colorAggr(ref, colorRef, fontA, condA, false); },
            /*macro*/ true, /*threadsafe*/ false);

        core::registerFn(L"SUMIFCOLOR",
            [colorAggr](const RangeArg& ref, const RangeArg& colorRef,
                        const ExcelObj& fontA, const ExcelObj& condA) -> ExcelObj*
            { return colorAggr(ref, colorRef, fontA, condA, true); },
            /*macro*/ true, /*threadsafe*/ false);

        // ── ISVISIBLE(range) ─────────────────────────────────────────────────
        core::registerFn(L"ISVISIBLE",
            [](const RangeArg& ref) -> ExcelObj*
            {
                const auto nR = (long)ref.nRows(), nC = (long)ref.nCols();
                if ((size_t)nR * nC > kMaxCells) return returnValue(CellError::Value);
                IDispatch* range = rangeDisp(ref);
                if (!range) return returnValue(CellError::Value);
                Releaser rr(range);

                std::vector<char> rHid((size_t)nR), cHid((size_t)nC);
                for (long i = 1; i <= nR; ++i) rHid[(size_t)i - 1] = rowHidden(range, i);
                for (long j = 1; j <= nC; ++j) cHid[(size_t)j - 1] = colHidden(range, j);

                std::vector<ExcelObj> flat;
                flat.reserve((size_t)nR * nC);
                for (long i = 0; i < nR; ++i)
                    for (long j = 0; j < nC; ++j)
                        flat.emplace_back(!(rHid[(size_t)i] || cHid[(size_t)j]));
                return core::output(core::makeArray(
                    (ExcelArrayBuilder::row_t)nR, (ExcelArrayBuilder::col_t)nC, flat));
            },
            /*macro*/ true, /*threadsafe*/ false);

        // ── VISIBLEAGGR(range, [func], [option]) ─────────────────────────────
        // VB 동등 집계자 세트(core/Aggregate.h). option: TEXTJOIN 구분자,
        // LARGE/SMALL k, PERCENTILE*/QUARTILE* 기준값.
        core::registerFn(L"VISIBLEAGGR",
            [](const RangeArg& ref, const ExcelObj& funcA, const ExcelObj& optA) -> ExcelObj*
            {
                const auto nR = (long)ref.nRows(), nC = (long)ref.nCols();
                if ((size_t)nR * nC > kMaxCells) return returnValue(CellError::Value);
                IDispatch* range = rangeDisp(ref);
                if (!range) return returnValue(CellError::Value);
                Releaser rr(range);

                std::vector<char> rHid((size_t)nR), cHid((size_t)nC);
                for (long i = 1; i <= nR; ++i) rHid[(size_t)i - 1] = rowHidden(range, i);
                for (long j = 1; j <= nC; ++j) cHid[(size_t)j - 1] = colHidden(range, j);

                const ExcelObj values = ref.value();
                const bool isArr = values.isType(ExcelType::Multi);
                ExcelArray varr(values, /*trim*/ false);
                auto valueAt = [&](long i, long j) -> ExcelObj {
                    return isArr
                        ? ExcelObj(varr.at((ExcelArray::row_t)i, (ExcelArray::col_t)j))
                        : ExcelObj(values);
                };

                std::wstring fn = funcA.isType(ExcelType::Multi)
                    ? std::wstring() : upperW(funcA.toString());
                if (!funcA.isType(ExcelType::Multi) && fn.empty())
                {
                    // no aggregator: spill the visible cells (visible rows × visible cols)
                    std::vector<ExcelObj> flat;
                    long outR = 0, outC = 0;
                    for (long j = 0; j < nC; ++j) if (!cHid[(size_t)j]) ++outC;
                    for (long i = 0; i < nR; ++i)
                    {
                        if (rHid[(size_t)i]) continue;
                        ++outR;
                        for (long j = 0; j < nC; ++j)
                            if (!cHid[(size_t)j]) flat.emplace_back(valueAt(i, j));
                    }
                    if (outR == 0 || outC == 0) return returnValue(CellError::NA);
                    return core::output(core::makeArray(
                        (ExcelArrayBuilder::row_t)outR, (ExcelArrayBuilder::col_t)outC, flat));
                }

                // collect visible cells; VB rule — a visible error is returned
                // as-is unless the aggregator is COUNT/COUNTA
                std::vector<ExcelObj> owned;
                owned.reserve((size_t)nR * nC);
                std::optional<ExcelObj> firstErr;
                for (long i = 0; i < nR; ++i)
                {
                    if (rHid[(size_t)i]) continue;
                    for (long j = 0; j < nC; ++j)
                    {
                        if (cHid[(size_t)j]) continue;
                        ExcelObj v = valueAt(i, j);
                        if (v.type() == ExcelType::Err && !firstErr) firstErr = v;
                        owned.emplace_back(std::move(v));
                    }
                }
                std::vector<const ExcelObj*> vals;
                vals.reserve(owned.size());
                for (auto& o : owned) vals.push_back(&o);

                // 집계자 이름·옵션은 원소별 리프팅(집계자 목록 → 집계값 배열).
                return core::mapLift(
                    [&](const ExcelObj& fe, const ExcelObj& oe) -> ExcelObj
                    {
                        const std::wstring f = upperW(fe.toString());
                        // GROUPBY 전용(PERCENTOF)은 여기서 지원하지 않는다.
                        if (f.empty() || f == L"PERCENTOF")
                            return ExcelObj(CellError::Value);
                        if (firstErr && f.find(L"COUNT") == std::wstring::npos)
                            return ExcelObj(*firstErr);
                        return core::aggregateObjs(f, vals,
                                                   oe.isMissing() ? nullptr : &oe);
                    },
                    funcA, optA);
            },
            /*macro*/ true, /*threadsafe*/ false);

        // ── ILOOKUP(find, lookup_array, image_range, [if_not_found], [match_mode], [search_mode]) ──
        core::registerFn(L"ILOOKUP",
            [](const ExcelObj& findA, const ExcelObj& lookupA, const RangeArg& imgRef,
               const ExcelObj& notFoundA, const ExcelObj& matchA, const ExcelObj& searchA) -> ExcelObj*
            {
                // 그림 함수는 셀당 1장 — 배열 찾을값은 거부(plan/22 그룹 C).
                if (findA.isType(ExcelType::Multi))
                    return returnValue(CellError::Value);
                if (findA.isMissing() || !lookupA.isType(ExcelType::Multi))
                    return returnValue(CellError::Value);
                ExcelArray lookup(lookupA, /*trim*/ false);
                const size_t n = (size_t)lookup.nRows() * lookup.nCols();
                if (n == 0 || n != (size_t)imgRef.nRows() * imgRef.nCols())
                    return returnValue(CellError::Ref);

                const int matchMode = matchA.get<int>(0);
                const int searchMode = searchA.get<int>(1);

                // linear match over the flattened lookup array
                const std::wstring findS = upperW(findA.toString());
                const bool findIsNum = findA.isType(ExcelType::Num);
                const double findN = findA.get<double>(0.0);

                long found = -1;                    // 0-based flat index
                double bestDiff = 0; long bestIdx = -1;
                const long total = (long)n;
                for (long k = 0; k < total; ++k)
                {
                    const long idx = (searchMode < 0) ? total - 1 - k : k;
                    ExcelObj v(lookup.at((ExcelArray::row_t)(idx / lookup.nCols()),
                                         (ExcelArray::col_t)(idx % lookup.nCols())));
                    bool hit = false;
                    if (matchMode == 2) hit = wildMatch(findA.toString(), v.toString());
                    else if (matchMode == 3)
                    {
                        try
                        {
                            const auto re = regexForLookup(findA.toString());
                            hit = std::regex_search(v.toString(), re);
                        }
                        catch (...) { return returnValue(CellError::Value); }
                    }
                    else if (findIsNum && v.isType(ExcelType::Num))
                    {
                        const double d = v.get<double>(0.0);
                        if (d == findN) hit = true;
                        else if (matchMode == -1 && d < findN &&
                                 (bestIdx < 0 || findN - d < bestDiff)) { bestDiff = findN - d; bestIdx = idx; }
                        else if (matchMode == 1 && d > findN &&
                                 (bestIdx < 0 || d - findN < bestDiff)) { bestDiff = d - findN; bestIdx = idx; }
                    }
                    else hit = (upperW(v.toString()) == findS);
                    if (hit) { found = idx; break; }
                }
                if (found < 0 && (matchMode == -1 || matchMode == 1)) found = bestIdx;
                if (found < 0)
                {
                    if (notFoundA.isMissing()) return returnValue(CellError::NA);
                    return returnValue(ExcelObj(notFoundA));
                }

                // matched cell inside the image range (row-major flat index)
                const long mi = found / (long)imgRef.nCols() + 1;
                const long mj = found % (long)imgRef.nCols() + 1;

                IDispatch* range = rangeDisp(imgRef);
                if (!range) return returnValue(CellError::Value);
                Releaser rr(range);
                IDispatch* srcCell = cellAt(range, mi, mj);
                if (!srcCell) return returnValue(CellError::Value);
                Releaser rc(srcCell);
                const std::wstring srcAddr = getBStr(srcCell, L"Address");

                // find a picture shape anchored at the matched cell
                IDispatch* ws = getObject(range, L"Worksheet");
                if (!ws) return returnValue(CellError::Value);
                Releaser rw(ws);
                IDispatch* shapes = getObject(ws, L"Shapes");
                if (!shapes) return returnValue(CellError::Value);
                Releaser rs(shapes);

                IDispatch* match = nullptr;
                const long cntS = getLong(shapes, L"Count", 0);
                for (long s = 1; s <= cntS && !match; ++s)
                {
                    IDispatch* shp = getObjectIdx(shapes, L"Item", s);
                    if (!shp) continue;
                    Releaser rsh(shp);
                    const long type = getLong(shp, L"Type", -1);
                    if (type != 13 && type != 11) continue;        // msoPicture / msoLinkedPicture
                    IDispatch* tlc = getObject(shp, L"TopLeftCell");
                    if (!tlc) continue;
                    Releaser rt(tlc);
                    if (getBStr(tlc, L"Address") == srcAddr) { match = shp; rsh.p = nullptr; }
                }
                if (!match)
                {
                    if (notFoundA.isMissing()) return returnValue(CellError::NA);
                    return returnValue(ExcelObj(notFoundA));
                }
                Releaser rm(match);

                // Shape.Copy → clipboard CF_DIB → temp BMP → deferred insert
                invokeRaw(match, L"Copy", DISPATCH_METHOD, nullptr, nullptr, 0);

                std::wstring path;
                for (int attempt = 0; attempt < 10 && path.empty(); ++attempt)
                {
                    if (OpenClipboard(nullptr))
                    {
                        if (HANDLE hDib = GetClipboardData(CF_DIB))
                        {
                            const size_t size = GlobalSize(hDib);
                            if (const void* src = GlobalLock(hDib))
                            {
                                wchar_t tmp[MAX_PATH]{};
                                GetTempPathW(MAX_PATH, tmp);
                                std::wstring p = std::wstring(tmp) + L"egilk_" +
                                    std::to_wstring(GetTickCount64()) + L".bmp";
                                FILE* f = nullptr;
                                if (_wfopen_s(&f, p.c_str(), L"wb") == 0 && f)
                                {
                                    BITMAPFILEHEADER bfh{};
                                    bfh.bfType = 0x4D42;
                                    const BITMAPINFOHEADER* bih = (const BITMAPINFOHEADER*)src;
                                    DWORD colorTable = bih->biClrUsed
                                        ? bih->biClrUsed * 4
                                        : (bih->biBitCount <= 8 ? (4u << bih->biBitCount) : 0);
                                    if (bih->biCompression == BI_BITFIELDS) colorTable += 12;
                                    bfh.bfOffBits = 14 + bih->biSize + colorTable;
                                    bfh.bfSize = (DWORD)(14 + size);
                                    fwrite(&bfh, 1, 14, f);
                                    fwrite(src, 1, size, f);
                                    fclose(f);
                                    path = p;
                                }
                                GlobalUnlock(hDib);
                            }
                        }
                        EmptyClipboard();
                        CloseClipboard();
                    }
                    if (path.empty()) Sleep(50);
                }
                if (path.empty()) return returnValue(CellError::Value);

                const std::wstring fullAddr = xloil::CallerInfo().address();
                const size_t bang = fullAddr.rfind(L'!');
                const std::wstring cell = bang == std::wstring::npos
                                              ? fullAddr : fullAddr.substr(bang + 1);
                queueInsertPicture(fullAddr, cell, path, 1, /*deleteLocalAfter*/ true);
                return returnValue(ExcelObj(std::wstring_view(L"")));
            },
            /*macro*/ true, /*threadsafe*/ false);
    }
}
