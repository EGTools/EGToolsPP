// Apply.h — element-wise application helpers (array-in → array-out principle).
//
// Lifts a scalar ExcelObj→ExcelObj function so that, given an array, it returns
// an array that spills (via core::output). Scalar in → scalar out.
//
// Native lifting semantics (measured, plan/22 §네이티브 리프팅 실측):
//   * element-wise over broadcast inputs — a 1-sized dimension repeats, an R×1
//     column × 1×C row makes an R×C outer product, and any element outside a
//     non-1 dimension is #N/A;
//   * if an element's result is itself an array it DEMOTES to its top-left
//     value (TEXTSPLIT/SEQUENCE/XLOOKUP all behave this way natively).

#pragma once
#include <algorithm>
#include <optional>
#include <vector>

#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ArrayBuilder.h>
#include <xlOil/StaticRegister.h>   // returnValue

#include "ArrayUtil.h"
#include "Spill.h"

namespace egtools::core
{
    // Read an optional scalar argument the way users expect: Missing/Nil
    // (omitted or empty cell) → default; array or unconvertible → nullopt so
    // the caller returns #VALUE! (or throws into the registry's common catch).
    // Replaces the get<T>(default) idiom, whose default applies to Missing ONLY
    // and which throws on arrays/strings (plan/22 §M2).
    template <class T>
    std::optional<T> optScalar(const xloil::ExcelObj& o, T def)
    {
        using namespace xloil;
        if (o.isMissing() || o.isType(ExcelType::Nil))
            return def;
        if (o.isType(ExcelType::Multi))
            return std::nullopt;
        return o.getIf<T>();
    }

    // 원소별 리프팅의 오류 정책(plan/22 실측): 네이티브는 리프팅되는 인수의
    // 원소가 오류면 그 원소의 결과가 그 오류다(전파 — XLOOKUP 키의 #DIV/0!,
    // CEILING.MATH의 #N/A 등). 오류를 직접 다루는 함수만 PassThrough:
    // IFNA(오류 소비), IFS/SWITCH(사용된 인수만 내부 규칙으로 전파),
    // MAXIFS 조건값(오류=불일치, 네이티브 결과 0).
    enum class LiftErrors { Propagate, PassThrough };

    // An argument view for lifting: scalar acts as a 1×1 array; broadcast rules
    // as above. at() returns nullptr for an element outside a non-1 dimension
    // (the caller emits #N/A for the whole result cell).
    class LiftArg
    {
    public:
        explicit LiftArg(const xloil::ExcelObj& o)
        {
            if (o.isType(xloil::ExcelType::Multi))
            {
                _arr.emplace(o);
                _rows = _arr->nRows();
                _cols = _arr->nCols();
            }
            else
            {
                _obj = &o;
                _rows = 1;
                _cols = 1;
            }
        }
        bool isArray() const { return _arr.has_value(); }
        xloil::ExcelArray::row_t nRows() const { return _rows; }
        xloil::ExcelArray::col_t nCols() const { return _cols; }
        const xloil::ExcelObj* at(xloil::ExcelArray::row_t i,
                                  xloil::ExcelArray::col_t j) const
        {
            if (!_arr) return _obj;
            const auto r = (_rows == 1) ? 0 : i;
            const auto c = (_cols == 1) ? 0 : j;
            if (r >= _rows || c >= _cols) return nullptr;
            return &_arr->at(r, c);
        }
    private:
        const xloil::ExcelObj* _obj = nullptr;
        std::optional<xloil::ExcelArray> _arr;
        xloil::ExcelArray::row_t _rows;
        xloil::ExcelArray::col_t _cols;
    };

    // Demote a nested-array element result to its top-left value (native rule).
    inline xloil::ExcelObj demote(xloil::ExcelObj&& v)
    {
        using namespace xloil;
        if (!v.isType(ExcelType::Multi))
            return std::move(v);
        ExcelArray a(v);
        return a.size() == 0 ? ExcelObj(CellError::Value) : ExcelObj(a.at(0, 0));
    }

    namespace detail
    {
        template <class Fn, size_t... I>
        xloil::ExcelObj callLift(Fn&& fn, const xloil::ExcelObj* const* elems,
                                 std::index_sequence<I...>)
        {
            return demote(fn(*elems[I]...));
        }
    }

    // Lift `fn` (a scalar core taking N `const ExcelObj&` and returning
    // ExcelObj) over the broadcast of the given arguments. If no argument is an
    // array this is a plain scalar call (an array RESULT still spills via
    // output()); otherwise returns the broadcast-shaped array with per-element
    // demotion. fn may throw — the registry's common catch yields #VALUE!.
    template <class Fn, class... TArgs>
    xloil::ExcelObj* mapLiftEx(LiftErrors errs, Fn&& fn, const TArgs&... argObjs)
    {
        using namespace xloil;
        static_assert(sizeof...(TArgs) >= 1, "mapLift needs at least one argument");

        const LiftArg views[] = { LiftArg(argObjs)... };
        constexpr size_t N = sizeof...(TArgs);

        bool anyArray = false;
        ExcelArray::row_t R = 1;
        ExcelArray::col_t C = 1;
        for (const auto& v : views)
        {
            anyArray = anyArray || v.isArray();
            R = std::max(R, v.nRows());
            C = std::max(C, v.nCols());
        }

        if (!anyArray)
        {
            if (errs == LiftErrors::Propagate)
            {
                const ExcelObj* scalars[] = { &argObjs... };
                for (size_t k = 0; k < N; ++k)
                    if (scalars[k]->type() == ExcelType::Err)
                        return returnValue(ExcelObj(*scalars[k]));
            }
            ExcelObj r = fn(argObjs...);
            if (r.isType(ExcelType::Multi))
                return output(std::move(r));
            return returnValue(std::move(r));
        }

        std::vector<ExcelObj> vals;
        vals.reserve((size_t)R * C);
        for (ExcelArray::row_t i = 0; i < R; ++i)
            for (ExcelArray::col_t j = 0; j < C; ++j)
            {
                const ExcelObj* elems[N];
                bool ok = true;
                const ExcelObj* err = nullptr;
                for (size_t k = 0; k < N; ++k)
                {
                    elems[k] = views[k].at(i, j);
                    ok = ok && (elems[k] != nullptr);
                    if (!err && elems[k] && elems[k]->type() == ExcelType::Err)
                        err = elems[k];
                }
                if (!ok) { vals.emplace_back(CellError::NA); continue; }
                if (errs == LiftErrors::Propagate && err)
                { vals.emplace_back(*err); continue; }
                vals.emplace_back(
                    detail::callLift(fn, elems, std::make_index_sequence<N>{}));
            }
        return output(makeArray(R, C, vals));
    }

    template <class Fn, class... TArgs>
    xloil::ExcelObj* mapLift(Fn&& fn, const TArgs&... argObjs)
    {
        return mapLiftEx(LiftErrors::Propagate, std::forward<Fn>(fn), argObjs...);
    }

    // Runtime-N variant of mapLift for variadic UDFs (registerRawFn).
    // fn: (const ExcelObj* const* elems, size_t n) -> ExcelObj. n == 0 is a
    // plain scalar call. Same broadcast + demotion rules as mapLift.
    template <class Fn>
    xloil::ExcelObj* mapLiftN(const xloil::ExcelObj* const* args, size_t n, Fn&& fn,
                              LiftErrors errs = LiftErrors::Propagate)
    {
        using namespace xloil;

        std::vector<LiftArg> views;
        views.reserve(n);
        bool anyArray = false;
        ExcelArray::row_t R = 1;
        ExcelArray::col_t C = 1;
        for (size_t k = 0; k < n; ++k)
        {
            views.emplace_back(*args[k]);
            anyArray = anyArray || views.back().isArray();
            R = std::max(R, views.back().nRows());
            C = std::max(C, views.back().nCols());
        }

        if (!anyArray)
        {
            if (errs == LiftErrors::Propagate)
                for (size_t k = 0; k < n; ++k)
                    if (args[k]->type() == ExcelType::Err)
                        return returnValue(ExcelObj(*args[k]));
            ExcelObj r = fn(args, n);
            if (r.isType(ExcelType::Multi))
                return output(std::move(r));
            return returnValue(std::move(r));
        }

        std::vector<const ExcelObj*> elems(n);
        std::vector<ExcelObj> vals;
        vals.reserve((size_t)R * C);
        for (ExcelArray::row_t i = 0; i < R; ++i)
            for (ExcelArray::col_t j = 0; j < C; ++j)
            {
                bool ok = true;
                const ExcelObj* err = nullptr;
                for (size_t k = 0; k < n; ++k)
                {
                    elems[k] = views[k].at(i, j);
                    ok = ok && (elems[k] != nullptr);
                    if (!err && elems[k] && elems[k]->type() == ExcelType::Err)
                        err = elems[k];
                }
                if (!ok) { vals.emplace_back(CellError::NA); continue; }
                if (errs == LiftErrors::Propagate && err)
                { vals.emplace_back(*err); continue; }
                vals.emplace_back(demote(fn(elems.data(), n)));
            }
        return output(makeArray(R, C, vals));
    }

    template <class Fn>
    xloil::ExcelObj* mapUnary(const xloil::ExcelObj& arg, Fn&& fn,
                              LiftErrors errs = LiftErrors::Propagate)
    {
        using namespace xloil;
        if (arg.isType(ExcelType::Multi))
        {
            ExcelArray a(arg);
            const auto R = a.nRows();
            const auto C = a.nCols();

            std::vector<ExcelObj> tmp;
            tmp.reserve((size_t)R * C);
            size_t strLen = 0;
            for (ExcelArray::row_t i = 0; i < R; ++i)
                for (ExcelArray::col_t j = 0; j < C; ++j)
                {
                    const ExcelObj& e = a.at(i, j);
                    if (errs == LiftErrors::Propagate && e.type() == ExcelType::Err)
                        tmp.emplace_back(e);
                    else
                        tmp.emplace_back(fn(e));
                    strLen += tmp.back().stringLength();
                }

            ExcelArrayBuilder b(R, C, strLen);
            size_t k = 0;
            for (ExcelArray::row_t i = 0; i < R; ++i)
                for (ExcelArray::col_t j = 0; j < C; ++j)
                    b(i, j) = tmp[k++];

            return output(b.toExcelObj());
        }
        if (errs == LiftErrors::Propagate && arg.type() == ExcelType::Err)
            return returnValue(ExcelObj(arg));
        return returnValue(fn(arg));
    }
}
