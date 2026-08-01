// Aggregate.h — ExcelObj adapter over AggCore.h, shared by GROUPBY/PIVOTBY
// (native 16-aggregator set) and VISIBLEAGGR (VB-parity set incl. TEXTJOIN/
// LARGE/SMALL/PERCENTILE/QUARTILE with an option argument).
#pragma once
#include "AggCore.h"

#include <xlOil/ExcelObj.h>

#include <optional>
#include <string>
#include <vector>

namespace egtools::core
{
    namespace detail_agg
    {
        inline bool aggAsNum(const xloil::ExcelObj& o, double& d)
        {
            auto t = o.type();
            if (t == xloil::ExcelType::Num || t == xloil::ExcelType::Int ||
                t == xloil::ExcelType::Bool)
                { d = o.get<double>(0.0); return true; }
            return false;
        }

        inline bool aggIsEmpty(const xloil::ExcelObj& o)
        {
            auto t = o.type();
            return t == xloil::ExcelType::Missing || t == xloil::ExcelType::Nil
                || (t == xloil::ExcelType::Str && o.stringLength() == 0);
        }

        inline std::wstring aggToText(const xloil::ExcelObj& o)
        {
            double d;
            if (o.type() == xloil::ExcelType::Bool)
                return o.get<bool>(false) ? L"TRUE" : L"FALSE";
            if (aggAsNum(o, d)) return egtools::agg::numToText(d);
            if (aggIsEmpty(o)) return L"";
            return o.toString();
        }

        inline xloil::CellError aggErr(egtools::agg::Err e)
        {
            using E = egtools::agg::Err;
            switch (e)
            {
            case E::Num:  return xloil::CellError::Num;
            case E::Div0: return xloil::CellError::Div0;
            case E::NA:   return xloil::CellError::NA;
            default:      return xloil::CellError::Value;
            }
        }
    }

    // The aggregator set the NATIVE GROUPBY/PIVOTBY function argument accepts.
    inline bool isGroupByAggregator(const std::wstring& fnUpper)
    {
        static const wchar_t* k[] = {
            L"SUM", L"AVERAGE", L"COUNT", L"COUNTA", L"MAX", L"MIN", L"PRODUCT",
            L"MEDIAN", L"MODE.SNGL", L"STDEV.S", L"STDEV.P", L"VAR.S", L"VAR.P",
            L"CONCAT", L"ARRAYTOTEXT", L"PERCENTOF" };
        for (auto* s : k) if (fnUpper == s) return true;
        return false;
    }

    // Aggregators that take their parameter from relative_to (TEXTJOIN delimiter,
    // LARGE/SMALL k, PERCENTILE*/QUARTILE* p/quart) — EGTools extension in
    // GROUPBY/PIVOTBY (the native functions reject these names).
    inline bool isParamAggregator(const std::wstring& fnUpper)
    {
        static const wchar_t* k[] = {
            L"TEXTJOIN", L"LARGE", L"SMALL",
            L"PERCENTILE", L"PERCENTILE.INC", L"PERCENTILE.EXC",
            L"QUARTILE", L"QUARTILE.INC", L"QUARTILE.EXC" };
        for (auto* s : k) if (fnUpper == s) return true;
        return false;
    }

    // Full GROUPBY/PIVOTBY set: native 16 + parameterised extensions.
    inline bool isGroupByAggregatorEx(const std::wstring& fnUpper)
    {
        return isGroupByAggregator(fnUpper) || isParamAggregator(fnUpper);
    }

    // Aggregate `vals` with fnUpper. `param` feeds TEXTJOIN's delimiter and
    // LARGE/SMALL/PERCENTILE*/QUARTILE*'s k/p/quart (VISIBLEAGGR only).
    // grandTotal is the PERCENTOF denominator (GROUPBY/PIVOTBY only).
    // Unknown names return #VALUE!.
    inline xloil::ExcelObj aggregateObjs(const std::wstring& fnUpper,
                                         const std::vector<const xloil::ExcelObj*>& vals,
                                         const xloil::ExcelObj* param = nullptr,
                                         double grandTotal = 0.0)
    {
        using namespace detail_agg;
        using xloil::ExcelObj;
        using xloil::CellError;
        const std::wstring& f = fnUpper;
        double d;

        if (f == L"PERCENTOF")
        {
            double s = 0;
            for (auto* o : vals) if (aggAsNum(*o, d)) s += d;
            if (grandTotal == 0.0) return ExcelObj(CellError::Div0);
            return ExcelObj(s / grandTotal);
        }
        if (f == L"COUNT")
        {
            int c = 0;
            for (auto* o : vals) if (aggAsNum(*o, d)) ++c;
            return ExcelObj((double)c);
        }
        if (f == L"COUNTA")
        {
            int c = 0;
            for (auto* o : vals) if (!aggIsEmpty(*o)) ++c;
            return ExcelObj((double)c);
        }
        if (f == L"CONCAT" || f == L"ARRAYTOTEXT" || f == L"TEXTJOIN")
        {
            std::wstring sep;
            if (f == L"ARRAYTOTEXT") sep = L", ";
            else if (f == L"TEXTJOIN") sep = param ? param->toString() : L"";
            std::wstring out;
            bool first = true;
            for (auto* o : vals)
            {
                if (!first) out += sep;
                out += aggToText(*o);
                first = false;
            }
            return ExcelObj(std::wstring_view(out));
        }

        std::vector<double> nums;
        nums.reserve(vals.size());
        for (auto* o : vals) if (aggAsNum(*o, d)) nums.push_back(d);

        std::optional<double> p;
        if (param && aggAsNum(*param, d)) p = d;

        // legacy MIN/MAX/PRODUCT on an all-text group return 0 (VB/native quirk)
        const auto r = egtools::agg::aggNumeric(f, nums, p);
        if (r.err != egtools::agg::Err::None) return ExcelObj(aggErr(r.err));
        return ExcelObj(r.value);
    }
}
