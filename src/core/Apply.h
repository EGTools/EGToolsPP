// Apply.h — element-wise application helpers (array-in → array-out principle).
//
// Lifts a scalar ExcelObj→ExcelObj function so that, given an array, it returns
// an array that spills (via core::output). Scalar in → scalar out.

#pragma once
#include <vector>

#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ArrayBuilder.h>
#include <xlOil/StaticRegister.h>   // returnValue

#include "Spill.h"

namespace egtools::core
{
    template <class Fn>
    xloil::ExcelObj* mapUnary(const xloil::ExcelObj& arg, Fn&& fn)
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
                    tmp.emplace_back(fn(a.at(i, j)));
                    strLen += tmp.back().stringLength();
                }

            ExcelArrayBuilder b(R, C, strLen);
            size_t k = 0;
            for (ExcelArray::row_t i = 0; i < R; ++i)
                for (ExcelArray::col_t j = 0; j < C; ++j)
                    b(i, j) = tmp[k++];

            return output(b.toExcelObj());
        }
        return returnValue(fn(arg));
    }
}
