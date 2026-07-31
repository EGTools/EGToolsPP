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
        // XOR over up to 8 args (each scalar or range): TRUE if an odd number of
        // logical-TRUE values. Errors propagate; text/blank ignored.
        ExcelObj* xorImpl(const ExcelObj* args[], int n)
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
            for (int k = 0; k < n; ++k)
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

        // XOR(logical1, [logical2], …) — up to 8 args.
        egtools::core::registerFn(L"XOR",
            [](const ExcelObj& a1, const ExcelObj& a2, const ExcelObj& a3, const ExcelObj& a4,
               const ExcelObj& a5, const ExcelObj& a6, const ExcelObj& a7, const ExcelObj& a8) -> ExcelObj*
            {
                const ExcelObj* args[] = { &a1, &a2, &a3, &a4, &a5, &a6, &a7, &a8 };
                return xorImpl(args, 8);
            });
    }
}
