// FxLogic.cpp — logical / error-handling functions.

#include "../core/Registry.h"
#include "../core/Apply.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // XOR over variadic args (each scalar or range): TRUE if an odd number of
        // logical-TRUE values. Errors propagate; text/blank ignored.
        ExcelObj* xorImpl(const ExcelObj* const args[], size_t n)
        {
            long trueCount = 0;
            bool any = false;
            const ExcelObj* firstErr = nullptr;

            auto scalar = [&](const ExcelObj& v)
            {
                switch (v.type())
                {
                case ExcelType::Err:  if (!firstErr) firstErr = &v; break;
                case ExcelType::Bool:
                case ExcelType::Num:
                case ExcelType::Int:  if (v.get<double>(0.0) != 0.0) ++trueCount; any = true; break;
                default: break;  // Str / Missing / Nil ignored
                }
            };
            for (size_t k = 0; k < n; ++k)
            {
                const ExcelObj& v = *args[k];
                if (v.isMissing()) continue;
                if (v.isType(ExcelType::Multi))
                {
                    ExcelArray a(v);
                    for (ExcelArray::row_t i = 0; i < a.nRows(); ++i)
                        for (ExcelArray::col_t j = 0; j < a.nCols(); ++j)
                            scalar(a.at(i, j));
                }
                else scalar(v);
            }
            if (firstErr) return returnValue(ExcelObj(*firstErr));
            if (!any)     return returnValue(CellError::Value);
            return returnValue(ExcelObj((trueCount % 2) != 0));
        }
    }

    void registerLogic()
    {
        // IFNA(value, value_if_na): replace #N/A with value_if_na (element-wise).
        egtools::core::registerFn(L"IFNA",
            [](const ExcelObj& value, const ExcelObj& valueIfNa) -> ExcelObj*
            {
                return egtools::core::mapUnary(value, [&](const ExcelObj& e)
                {
                    return (e == CellError::NA) ? ExcelObj(valueIfNa) : ExcelObj(e);
                });
            });

        // XOR(logical1, [logical2], …) — 네이티브와 동일한 가변 인수(최대 254).
        egtools::core::registerRawFn(L"XOR",
            [](const xloil::FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                return xorImpl(args, info.numArgs());
            });
    }
}
