// ArrayUtil.h — helpers for building array results (array-out functions).
#pragma once
#include <vector>

#include <xlOil/ExcelObj.h>
#include <xlOil/ArrayBuilder.h>

namespace egtools::core
{
    // Build a rows×cols array ExcelObj from row-major values (copies strings).
    // Missing trailing cells are filled with #N/A.
    inline xloil::ExcelObj makeArray(
        xloil::ExcelArrayBuilder::row_t rows,
        xloil::ExcelArrayBuilder::col_t cols,
        const std::vector<xloil::ExcelObj>& vals)
    {
        using namespace xloil;
        size_t strLen = 0;
        for (const auto& v : vals) strLen += v.stringLength();

        ExcelArrayBuilder b(rows, cols, strLen);
        size_t k = 0;
        for (ExcelArrayBuilder::row_t i = 0; i < rows; ++i)
            for (ExcelArrayBuilder::col_t j = 0; j < cols; ++j, ++k)
                b(i, j) = (k < vals.size()) ? vals[k] : ExcelObj(CellError::NA);
        return b.toExcelObj();
    }
}
