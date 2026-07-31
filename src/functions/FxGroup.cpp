// FxGroup.cpp — GROUPBY / PIVOTBY (text-aggregator variant).
//
// Native GROUPBY/PIVOTBY take an eta-reduced function (e.g. SUM) which a plain
// UDF cannot receive. This variant takes the aggregator as TEXT:
//   "SUM" | "AVERAGE" | "COUNT" | "COUNTA" | "MAX" | "MIN" | "PRODUCT".

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <string>
#include <vector>
#include <algorithm>
#include <limits>
#include <cwctype>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        std::wstring upper(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)std::towupper(c);
            return s;
        }

        bool asNum(const ExcelObj& o, double& d)
        {
            auto t = o.type();
            if (t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool)
                { d = o.get<double>(0.0); return true; }
            return false;
        }

        // grandTotal is only used by PERCENTOF (group sum ÷ overall/column total).
        ExcelObj aggregate(const std::wstring& fn, const std::vector<const ExcelObj*>& vals,
                           double grandTotal = 0.0)
        {
            const std::wstring f = upper(fn);
            if (f == L"PERCENTOF")
            {
                double s = 0, d;
                for (auto* o : vals) if (asNum(*o, d)) s += d;
                if (grandTotal == 0.0) return ExcelObj(CellError::Div0);
                return ExcelObj(s / grandTotal);
            }
            if (f == L"COUNT")
            {
                int c = 0; double d;
                for (auto* o : vals) if (asNum(*o, d)) ++c;
                return ExcelObj((double)c);
            }
            if (f == L"COUNTA")
            {
                int c = 0;
                for (auto* o : vals)
                {
                    auto t = o->type();
                    if (!(t == ExcelType::Missing || t == ExcelType::Nil
                          || (t == ExcelType::Str && o->stringLength() == 0))) ++c;
                }
                return ExcelObj((double)c);
            }
            if (f == L"SUM")
            {
                double s = 0, d;
                for (auto* o : vals) if (asNum(*o, d)) s += d;
                return ExcelObj(s);
            }
            if (f == L"AVERAGE")
            {
                double s = 0, d; int c = 0;
                for (auto* o : vals) if (asNum(*o, d)) { s += d; ++c; }
                return c ? ExcelObj(s / c) : ExcelObj(CellError::Div0);
            }
            if (f == L"MAX" || f == L"MIN")
            {
                const bool mx = (f == L"MAX");
                double best = mx ? -std::numeric_limits<double>::infinity()
                                 :  std::numeric_limits<double>::infinity();
                bool any = false; double d;
                for (auto* o : vals) if (asNum(*o, d)) { any = true; best = mx ? (d > best ? d : best) : (d < best ? d : best); }
                return ExcelObj(any ? best : 0.0);
            }
            if (f == L"PRODUCT")
            {
                double p = 1, d; bool any = false;
                for (auto* o : vals) if (asNum(*o, d)) { any = true; p *= d; }
                return ExcelObj(any ? p : 0.0);
            }
            return ExcelObj(CellError::Value);
        }

        int sortOrder(const ExcelObj& o) { return o.isMissing() ? 1 : o.get<int>(1); }
    }

    void registerGroup()
    {
        // GROUPBY(row_fields, values, function_text, [sort_order]).
        // row_fields R×K (composite key), values R×V; output groups × (K+V).
        egtools::core::registerFn(L"GROUPBY",
            [](const ExcelObj& rowF, const ExcelObj& valuesA, const ExcelObj& fnA,
               const ExcelObj& sortA) -> ExcelObj*
            {
                try
                {
                    ExcelArray keys(rowF), vals(valuesA);
                    const auto R = keys.nRows();
                    const auto K = keys.nCols();
                    const auto V = vals.nCols();
                    if (vals.nRows() != R || R == 0) return returnValue(CellError::Value);
                    const std::wstring fn = fnA.toString();
                    const int order = sortOrder(sortA);

                    struct Group { std::vector<ExcelObj> key; std::vector<ExcelArray::row_t> rows; };
                    std::vector<Group> groups;
                    for (ExcelArray::row_t r = 0; r < R; ++r)
                    {
                        std::vector<ExcelObj> kv;
                        for (ExcelArray::col_t k = 0; k < K; ++k) kv.emplace_back(keys.at(r, k));
                        auto it = std::find_if(groups.begin(), groups.end(), [&](const Group& g) {
                            for (size_t k = 0; k < kv.size(); ++k)
                                if (ExcelObj::compare(g.key[k], kv[k]) != 0) return false;
                            return true;
                        });
                        if (it == groups.end()) groups.push_back({ std::move(kv), { r } });
                        else it->rows.push_back(r);
                    }

                    std::stable_sort(groups.begin(), groups.end(), [&](const Group& a, const Group& b) {
                        int c = ExcelObj::compare(a.key[0], b.key[0]);
                        return order < 0 ? c > 0 : c < 0;
                    });

                    // PERCENTOF: per value-column grand total (group ÷ column total).
                    std::vector<double> grand(V, 0.0);
                    if (upper(fn) == L"PERCENTOF")
                        for (ExcelArray::row_t r = 0; r < R; ++r)
                            for (ExcelArray::col_t v = 0; v < V; ++v)
                            { double d; if (asNum(vals.at(r, v), d)) grand[v] += d; }

                    std::vector<ExcelObj> out;
                    for (auto& g : groups)
                    {
                        for (auto& kc : g.key) out.emplace_back(kc);
                        for (ExcelArray::col_t v = 0; v < V; ++v)
                        {
                            std::vector<const ExcelObj*> cells;
                            for (auto r : g.rows) cells.push_back(&vals.at(r, v));
                            out.emplace_back(aggregate(fn, cells, grand[v]));
                        }
                    }
                    return egtools::core::output(egtools::core::makeArray(
                        (ExcelArrayBuilder::row_t)groups.size(),
                        (ExcelArrayBuilder::col_t)(K + V), out));
                }
                catch (...) { return returnValue(CellError::Value); }
            });

        // PIVOTBY(row_fields, col_fields, values, function_text, [sort_order]).
        // Single-column keys; output (1+rows)×(1+cols) with header row/col.
        egtools::core::registerFn(L"PIVOTBY",
            [](const ExcelObj& rowF, const ExcelObj& colF, const ExcelObj& valuesA,
               const ExcelObj& fnA, const ExcelObj& sortA) -> ExcelObj*
            {
                try
                {
                    ExcelArray rk(rowF), ck(colF), vv(valuesA);
                    const auto R = rk.nRows();
                    if (ck.nRows() != R || vv.nRows() != R || R == 0) return returnValue(CellError::Value);
                    const std::wstring fn = fnA.toString();
                    const int order = sortOrder(sortA);

                    auto uniqueKeys = [&](ExcelArray& a) {
                        std::vector<ExcelObj> u;
                        for (ExcelArray::row_t r = 0; r < R; ++r)
                        {
                            const ExcelObj& v = a.at(r, 0);
                            if (std::none_of(u.begin(), u.end(), [&](const ExcelObj& e) { return ExcelObj::compare(e, v) == 0; }))
                                u.emplace_back(v);
                        }
                        std::stable_sort(u.begin(), u.end(), [&](const ExcelObj& x, const ExcelObj& y) {
                            int c = ExcelObj::compare(x, y);
                            return order < 0 ? c > 0 : c < 0;
                        });
                        return u;
                    };
                    std::vector<ExcelObj> rows = uniqueKeys(rk);
                    std::vector<ExcelObj> cols = uniqueKeys(ck);

                    auto idx = [](const std::vector<ExcelObj>& u, const ExcelObj& v) -> size_t {
                        for (size_t i = 0; i < u.size(); ++i) if (ExcelObj::compare(u[i], v) == 0) return i;
                        return (size_t)-1;
                    };

                    const size_t nR = rows.size(), nC = cols.size();
                    std::vector<std::vector<std::vector<const ExcelObj*>>> bucket(
                        nR, std::vector<std::vector<const ExcelObj*>>(nC));
                    double grand = 0.0;  // PERCENTOF: overall total
                    for (ExcelArray::row_t r = 0; r < R; ++r)
                    {
                        bucket[idx(rows, rk.at(r, 0))][idx(cols, ck.at(r, 0))].push_back(&vv.at(r, 0));
                        double d; if (asNum(vv.at(r, 0), d)) grand += d;
                    }

                    std::vector<ExcelObj> out;
                    out.emplace_back(std::wstring_view(L""));          // top-left
                    for (auto& c : cols) out.emplace_back(c);          // header row
                    for (size_t i = 0; i < nR; ++i)
                    {
                        out.emplace_back(rows[i]);                     // row header
                        for (size_t j = 0; j < nC; ++j)
                            out.emplace_back(aggregate(fn, bucket[i][j], grand));
                    }
                    return egtools::core::output(egtools::core::makeArray(
                        (ExcelArrayBuilder::row_t)(nR + 1),
                        (ExcelArrayBuilder::col_t)(nC + 1), out));
                }
                catch (...) { return returnValue(CellError::Value); }
            });
    }
}
