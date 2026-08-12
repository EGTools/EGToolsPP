// FxArrayShape.cpp — 2024/365 array-shaping functions
// (VSTACK, HSTACK, TAKE, DROP, TOROW, TOCOL, CHOOSEROWS, CHOOSECOLS, EXPAND).

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "../core/Apply.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <vector>
#include <algorithm>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        using Grid = std::vector<std::vector<ExcelObj>>;  // rows of cells

        Grid readGrid(const ExcelObj& o)
        {
            Grid g;
            if (o.isMissing()) return g;
            if (o.isType(ExcelType::Multi))
            {
                ExcelArray a(o);
                for (ExcelArray::row_t r = 0; r < a.nRows(); ++r)
                {
                    std::vector<ExcelObj> row;
                    row.reserve(a.nCols());
                    for (ExcelArray::col_t c = 0; c < a.nCols(); ++c) row.emplace_back(a.at(r, c));
                    g.push_back(std::move(row));
                }
            }
            else g.push_back({ ExcelObj(o) });
            return g;
        }

        size_t width(const Grid& g) { return g.empty() ? 0 : g[0].size(); }

        // Grid → array ExcelObj (스칼라 코어용 — mapLift가 스필/강등을 결정).
        ExcelObj gridObj(const Grid& g)
        {
            if (g.empty() || g[0].empty()) return ExcelObj(CellError::Value);
            const auto rows = (ExcelArrayBuilder::row_t)g.size();
            const auto cols = (ExcelArrayBuilder::col_t)g[0].size();
            std::vector<ExcelObj> flat;
            flat.reserve((size_t)rows * cols);
            for (auto& row : g)
                for (auto& cell : row) flat.emplace_back(cell);
            return egtools::core::makeArray(rows, cols, flat);
        }

        ExcelObj* emit(const Grid& g)
        {
            if (g.empty() || g[0].empty()) return returnValue(CellError::Value);
            ExcelObj o = gridObj(g);
            return egtools::core::output(std::move(o));
        }

        bool isBlank(const ExcelObj& v)
        {
            auto t = v.type();
            return t == ExcelType::Missing || t == ExcelType::Nil
                || (t == ExcelType::Str && v.stringLength() == 0);
        }
        bool isErr(const ExcelObj& v) { return v.type() == ExcelType::Err; }

        // Resolve a 1-based (possibly negative) index against a dimension size.
        bool resolveIndex(int idx, size_t n, size_t& out)
        {
            if (idx > 0 && (size_t)idx <= n) { out = (size_t)(idx - 1); return true; }
            if (idx < 0 && (size_t)(-idx) <= n) { out = n - (size_t)(-idx); return true; }
            return false;
        }
    }

    void registerArrayShape()
    {
        // VSTACK(array1, …) — stack vertically; widen to widest, pad with #N/A.
        egtools::core::registerRawFn(L"VSTACK",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                std::vector<Grid> grids; size_t w = 0;
                for (size_t i = 0; i < info.numArgs(); ++i)
                    { Grid g = readGrid(*args[i]); if (!g.empty()) { w = std::max(w, width(g)); grids.push_back(std::move(g)); } }
                if (grids.empty()) return returnValue(CellError::Value);
                Grid out;
                for (auto& g : grids)
                    for (auto& row : g)
                    {
                        std::vector<ExcelObj> r = row;
                        while (r.size() < w) r.emplace_back(CellError::NA);
                        out.push_back(std::move(r));
                    }
                return emit(out);
            });

        // HSTACK(array1, …) — stack horizontally; heighten to tallest, pad with #N/A.
        egtools::core::registerRawFn(L"HSTACK",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                std::vector<Grid> grids; size_t h = 0;
                for (size_t i = 0; i < info.numArgs(); ++i)
                    { Grid g = readGrid(*args[i]); if (!g.empty()) { h = std::max(h, g.size()); grids.push_back(std::move(g)); } }
                if (grids.empty()) return returnValue(CellError::Value);
                Grid out(h);
                for (auto& g : grids)
                {
                    const size_t gw = width(g);
                    for (size_t i = 0; i < h; ++i)
                        for (size_t j = 0; j < gw; ++j)
                            out[i].emplace_back(i < g.size() ? g[i][j] : ExcelObj(CellError::NA));
                }
                return emit(out);
            });

        // TAKE(array, rows, [cols]) — keep first/last |rows| rows and |cols| cols.
        egtools::core::registerFn(L"TAKE",
            [](const ExcelObj& array, const ExcelObj& rowsA, const ExcelObj& colsA) -> ExcelObj*
            {
                Grid g = readGrid(array);
                if (g.empty()) return returnValue(CellError::Value);
                const size_t R = g.size(), C = width(g);
                // rows/cols 배열 → 강등 리프팅(네이티브 정합, plan/22 U01).
                return egtools::core::mapLift(
                    [&](const ExcelObj& re, const ExcelObj& ce) -> ExcelObj
                {
                    auto slice = [](size_t n, const ExcelObj& a) -> std::pair<size_t, size_t> {
                        if (a.isMissing()) return { 0, n };
                        int k = a.get<int>(0);
                        if (k >= 0) return { 0, std::min((size_t)k, n) };
                        size_t cnt = std::min((size_t)(-k), n);
                        return { n - cnt, cnt };
                    };
                    auto [r0, rn] = slice(R, re);
                    auto [c0, cn] = slice(C, ce);
                    if (rn == 0 || cn == 0) return ExcelObj(CellError::Value);
                    Grid out;
                    for (size_t i = 0; i < rn; ++i)
                    {
                        std::vector<ExcelObj> row;
                        for (size_t j = 0; j < cn; ++j) row.emplace_back(g[r0 + i][c0 + j]);
                        out.push_back(std::move(row));
                    }
                    return gridObj(out);
                }, rowsA, colsA);
            });

        // DROP(array, rows, [cols]) — drop first/last |rows| rows and |cols| cols.
        egtools::core::registerFn(L"DROP",
            [](const ExcelObj& array, const ExcelObj& rowsA, const ExcelObj& colsA) -> ExcelObj*
            {
                Grid g = readGrid(array);
                if (g.empty()) return returnValue(CellError::Value);
                const size_t R = g.size(), C = width(g);
                return egtools::core::mapLift(
                    [&](const ExcelObj& re, const ExcelObj& ce) -> ExcelObj
                {
                    auto keep = [](size_t n, const ExcelObj& a) -> std::pair<size_t, size_t> {
                        if (a.isMissing()) return { 0, n };
                        int k = a.get<int>(0);
                        size_t d = std::min((size_t)std::abs(k), n);
                        if (k >= 0) return { d, n - d };          // drop from start
                        return { 0, n - d };                       // drop from end
                    };
                    auto [r0, rn] = keep(R, re);
                    auto [c0, cn] = keep(C, ce);
                    if (rn == 0 || cn == 0) return ExcelObj(CellError::Value);
                    Grid out;
                    for (size_t i = 0; i < rn; ++i)
                    {
                        std::vector<ExcelObj> row;
                        for (size_t j = 0; j < cn; ++j) row.emplace_back(g[r0 + i][c0 + j]);
                        out.push_back(std::move(row));
                    }
                    return gridObj(out);
                }, rowsA, colsA);
            });

        // TOROW(array, [ignore], [scan_by_col]) — flatten into one row.
        // TOCOL(array, [ignore], [scan_by_col]) — flatten into one column.
        auto flatten = [](const ExcelObj& array, const ExcelObj& ignoreA, const ExcelObj& byColA,
                          bool asRow) -> ExcelObj*
        {
            Grid g = readGrid(array);
            if (g.empty()) return returnValue(CellError::Value);
            const size_t R = g.size(), C = width(g);
            // 옵션 배열 → 강등 리프팅.
            return egtools::core::mapLift(
                [&, asRow](const ExcelObj& ie, const ExcelObj& be) -> ExcelObj
            {
                const int ignore = ie.isMissing() ? 0 : ie.get<int>(0);
                const bool byCol = be.isMissing() ? false : (be.get<double>(0.0) != 0.0);
                std::vector<ExcelObj> vals;
                auto consider = [&](const ExcelObj& v) {
                    if ((ignore == 1 || ignore == 3) && isBlank(v)) return;
                    if ((ignore == 2 || ignore == 3) && isErr(v)) return;
                    vals.emplace_back(v);
                };
                if (byCol) for (size_t j = 0; j < C; ++j) for (size_t i = 0; i < R; ++i) consider(g[i][j]);
                else       for (size_t i = 0; i < R; ++i) for (size_t j = 0; j < C; ++j) consider(g[i][j]);
                if (vals.empty()) return ExcelObj(CellError::Value);
                return asRow
                    ? egtools::core::makeArray(1, (ExcelArrayBuilder::col_t)vals.size(), vals)
                    : egtools::core::makeArray((ExcelArrayBuilder::row_t)vals.size(), 1, vals);
            }, ignoreA, byColA);
        };
        egtools::core::registerFn(L"TOROW",
            [flatten](const ExcelObj& a, const ExcelObj& ig, const ExcelObj& bc) -> ExcelObj*
            { return flatten(a, ig, bc, true); });
        egtools::core::registerFn(L"TOCOL",
            [flatten](const ExcelObj& a, const ExcelObj& ig, const ExcelObj& bc) -> ExcelObj*
            { return flatten(a, ig, bc, false); });

        // CHOOSEROWS(array, row1, …) — select rows by 1-based (neg = from end) index.
        egtools::core::registerRawFn(L"CHOOSEROWS",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                if (info.numArgs() < 1) return returnValue(CellError::Value);
                Grid g = readGrid(*args[0]);
                if (g.empty()) return returnValue(CellError::Value);
                const size_t R = g.size();
                Grid out;
                // 인덱스 인수는 배열도 허용(네이티브: {1,3}, SEQUENCE(n) 등).
                for (size_t k = 1; k < info.numArgs(); ++k)
                {
                    if (args[k]->isMissing()) continue;
                    if (args[k]->isType(ExcelType::Multi))
                    {
                        ExcelArray a(*args[k]);
                        const size_t n = (size_t)a.nRows() * a.nCols();
                        for (size_t i = 0; i < n; ++i)
                        {
                            size_t r;
                            if (!resolveIndex(a.at(i).get<int>(0), R, r))
                                return returnValue(CellError::Value);
                            out.push_back(g[r]);
                        }
                        continue;
                    }
                    size_t r;
                    if (!resolveIndex(args[k]->get<int>(0), R, r)) return returnValue(CellError::Value);
                    out.push_back(g[r]);
                }
                if (out.empty()) return returnValue(CellError::Value);
                return emit(out);
            });

        // CHOOSECOLS(array, col1, …) — select columns by 1-based index.
        egtools::core::registerRawFn(L"CHOOSECOLS",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                if (info.numArgs() < 1) return returnValue(CellError::Value);
                Grid g = readGrid(*args[0]);
                if (g.empty()) return returnValue(CellError::Value);
                const size_t C = width(g);
                std::vector<size_t> cols;
                // 인덱스 인수는 배열도 허용(네이티브 정합).
                for (size_t k = 1; k < info.numArgs(); ++k)
                {
                    if (args[k]->isMissing()) continue;
                    if (args[k]->isType(ExcelType::Multi))
                    {
                        ExcelArray a(*args[k]);
                        const size_t n = (size_t)a.nRows() * a.nCols();
                        for (size_t i = 0; i < n; ++i)
                        {
                            size_t c;
                            if (!resolveIndex(a.at(i).get<int>(0), C, c))
                                return returnValue(CellError::Value);
                            cols.push_back(c);
                        }
                        continue;
                    }
                    size_t c;
                    if (!resolveIndex(args[k]->get<int>(0), C, c)) return returnValue(CellError::Value);
                    cols.push_back(c);
                }
                if (cols.empty()) return returnValue(CellError::Value);
                Grid out;
                for (auto& row : g)
                {
                    std::vector<ExcelObj> r;
                    for (auto c : cols) r.emplace_back(row[c]);
                    out.push_back(std::move(r));
                }
                return emit(out);
            });

        // EXPAND(array, rows, [cols], [pad_with]) — grow to rows×cols, padding.
        egtools::core::registerFn(L"EXPAND",
            [](const ExcelObj& array, const ExcelObj& rowsA, const ExcelObj& colsA,
               const ExcelObj& padA) -> ExcelObj*
            {
                Grid g = readGrid(array);
                if (g.empty()) return returnValue(CellError::Value);
                const size_t R = g.size(), C = width(g);
                return egtools::core::mapLift(
                    [&](const ExcelObj& re, const ExcelObj& ce, const ExcelObj& pe) -> ExcelObj
                {
                    const size_t nr = re.isMissing() ? R : (size_t)re.get<int>((int)R);
                    const size_t nc = ce.isMissing() ? C : (size_t)ce.get<int>((int)C);
                    if (nr < R || nc < C) return ExcelObj(CellError::Value);
                    const ExcelObj pad = pe.isMissing() ? ExcelObj(CellError::NA) : ExcelObj(pe);
                    Grid out;
                    for (size_t i = 0; i < nr; ++i)
                    {
                        std::vector<ExcelObj> row;
                        for (size_t j = 0; j < nc; ++j)
                            row.emplace_back((i < R && j < C) ? g[i][j] : pad);
                        out.push_back(std::move(row));
                    }
                    return gridObj(out);
                }, rowsA, colsA, padA);
            });

        // Flatten a grid into a single row-major vector (for WRAP*).
        auto flat1d = [](const Grid& g) {
            std::vector<ExcelObj> v;
            for (auto& row : g) for (auto& c : row) v.emplace_back(c);
            return v;
        };

        // WRAPROWS(vector, wrap_count, [pad_with]) — wrap a vector into rows.
        egtools::core::registerFn(L"WRAPROWS",
            [flat1d](const ExcelObj& vec, const ExcelObj& wcA, const ExcelObj& padA) -> ExcelObj*
            {
                std::vector<ExcelObj> v = flat1d(readGrid(vec));
                if (v.empty()) return returnValue(CellError::Value);
                return egtools::core::mapLift(
                    [&](const ExcelObj& we, const ExcelObj& pe) -> ExcelObj
                {
                    const int wc = we.isMissing() ? 0 : we.get<int>(0);
                    if (wc <= 0) return ExcelObj(CellError::Value);
                    const ExcelObj pad = pe.isMissing() ? ExcelObj(CellError::NA) : ExcelObj(pe);
                    const size_t cols = (size_t)wc;
                    const size_t rows = (v.size() + cols - 1) / cols;
                    Grid out;
                    for (size_t i = 0; i < rows; ++i)
                    {
                        std::vector<ExcelObj> row;
                        for (size_t j = 0; j < cols; ++j)
                        {
                            const size_t k = i * cols + j;
                            row.emplace_back(k < v.size() ? v[k] : pad);
                        }
                        out.push_back(std::move(row));
                    }
                    return gridObj(out);
                }, wcA, padA);
            });

        // WRAPCOLS(vector, wrap_count, [pad_with]) — wrap a vector into columns.
        egtools::core::registerFn(L"WRAPCOLS",
            [flat1d](const ExcelObj& vec, const ExcelObj& wcA, const ExcelObj& padA) -> ExcelObj*
            {
                std::vector<ExcelObj> v = flat1d(readGrid(vec));
                if (v.empty()) return returnValue(CellError::Value);
                return egtools::core::mapLift(
                    [&](const ExcelObj& we, const ExcelObj& pe) -> ExcelObj
                {
                    const int wc = we.isMissing() ? 0 : we.get<int>(0);
                    if (wc <= 0) return ExcelObj(CellError::Value);
                    const ExcelObj pad = pe.isMissing() ? ExcelObj(CellError::NA) : ExcelObj(pe);
                    const size_t rows = (size_t)wc;
                    const size_t cols = (v.size() + rows - 1) / rows;
                    Grid out(rows, std::vector<ExcelObj>(cols, pad));
                    for (size_t k = 0; k < v.size(); ++k)
                        out[k % rows][k / rows] = v[k];   // fill column-major
                    return gridObj(out);
                }, wcA, padA);
            });

        // TRIMRANGE(array, [trim_rows], [trim_cols]) — drop blank edge rows/cols.
        // Mode per axis: 0 none, 1 leading, 2 trailing, 3 both (default).
        // NB: reads with trim=false — the default ExcelArray ctor silently drops
        // trailing blank rows/cols, which would pre-trim exactly what this
        // function is supposed to control.
        egtools::core::registerFn(L"TRIMRANGE",
            [](const ExcelObj& array, const ExcelObj& rowsA, const ExcelObj& colsA) -> ExcelObj*
            {
                Grid g;
                if (array.isType(ExcelType::Multi))
                {
                    ExcelArray a(array, /*trim*/ false);
                    for (ExcelArray::row_t r = 0; r < a.nRows(); ++r)
                    {
                        std::vector<ExcelObj> row;
                        row.reserve(a.nCols());
                        for (ExcelArray::col_t c = 0; c < a.nCols(); ++c) row.emplace_back(a.at(r, c));
                        g.push_back(std::move(row));
                    }
                }
                else if (!array.isMissing()) g.push_back({ ExcelObj(array) });
                if (g.empty()) return returnValue(CellError::Value);
                const size_t R = g.size(), C = width(g);
                return egtools::core::mapLift(
                    [&](const ExcelObj& re, const ExcelObj& ce) -> ExcelObj
                {
                    const int mr = re.isMissing() ? 3 : re.get<int>(3);
                    const int mc = ce.isMissing() ? 3 : ce.get<int>(3);
                    if (mr < 0 || mr > 3 || mc < 0 || mc > 3) return ExcelObj(CellError::Value);

                    auto blankRow = [&](size_t r) {
                        for (size_t j = 0; j < C; ++j) if (!isBlank(g[r][j])) return false;
                        return true;
                    };
                    auto blankCol = [&](size_t c) {
                        for (size_t i = 0; i < R; ++i) if (!isBlank(g[i][c])) return false;
                        return true;
                    };
                    size_t r0 = 0, r1 = R, c0 = 0, c1 = C;   // half-open [r0,r1) × [c0,c1)
                    if (mr & 1) while (r0 < r1 && blankRow(r0)) ++r0;
                    if (mr & 2) while (r1 > r0 && blankRow(r1 - 1)) --r1;
                    if (mc & 1) while (c0 < c1 && blankCol(c0)) ++c0;
                    if (mc & 2) while (c1 > c0 && blankCol(c1 - 1)) --c1;
                    if (r0 >= r1 || c0 >= c1) return ExcelObj(CellError::Value);

                    Grid out;
                    for (size_t i = r0; i < r1; ++i)
                    {
                        std::vector<ExcelObj> row;
                        for (size_t j = c0; j < c1; ++j) row.emplace_back(g[i][j]);
                        out.push_back(std::move(row));
                    }
                    return gridObj(out);
                }, rowsA, colsA);
            });
    }
}
