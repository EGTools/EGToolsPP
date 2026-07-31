// FxGen.cpp — array-generating functions (SEQUENCE, RANDARRAY).

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"

#include <xlOil/xlOil.h>
#include <vector>
#include <random>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        int argInt(const ExcelObj& v, int def) { return v.isMissing() ? def : v.get<int>(def); }
        double argDbl(const ExcelObj& v, double def) { return v.isMissing() ? def : v.get<double>(def); }

        std::mt19937& rng()
        {
            static std::random_device rd;
            static std::mt19937 g(rd());
            return g;
        }
    }

    void registerGen()
    {
        // SEQUENCE(rows, [cols], [start], [step]) — sequential numbers.
        egtools::core::registerFn(L"SEQUENCE",
            [](const ExcelObj& rowsA, const ExcelObj& colsA,
               const ExcelObj& startA, const ExcelObj& stepA) -> ExcelObj*
            {
                const int rows = argInt(rowsA, 1);
                const int cols = argInt(colsA, 1);
                if (rows <= 0 || cols <= 0) return returnValue(CellError::Value);
                const double start = argDbl(startA, 1.0);
                const double step = argDbl(stepA, 1.0);

                std::vector<ExcelObj> vals;
                vals.reserve((size_t)rows * cols);
                double v = start;
                for (int i = 0; i < rows * cols; ++i) { vals.emplace_back(v); v += step; }
                return egtools::core::output(
                    egtools::core::makeArray((ExcelArrayBuilder::row_t)rows,
                                             (ExcelArrayBuilder::col_t)cols, vals));
            });

        // RANDARRAY([rows], [cols], [min], [max], [integer]).
        egtools::core::registerFn(L"RANDARRAY",
            [](const ExcelObj& rowsA, const ExcelObj& colsA, const ExcelObj& minA,
               const ExcelObj& maxA, const ExcelObj& intA) -> ExcelObj*
            {
                const int rows = argInt(rowsA, 1);
                const int cols = argInt(colsA, 1);
                if (rows <= 0 || cols <= 0) return returnValue(CellError::Value);
                const double lo = argDbl(minA, 0.0);
                const double hi = argDbl(maxA, 1.0);
                const bool asInt = intA.isMissing() ? false : (intA.get<double>(0.0) != 0.0);
                if (hi < lo) return returnValue(CellError::Value);

                std::vector<ExcelObj> vals;
                vals.reserve((size_t)rows * cols);
                if (asInt)
                {
                    std::uniform_int_distribution<long long> d((long long)lo, (long long)hi);
                    for (int i = 0; i < rows * cols; ++i) vals.emplace_back((double)d(rng()));
                }
                else
                {
                    std::uniform_real_distribution<double> d(lo, hi);
                    for (int i = 0; i < rows * cols; ++i) vals.emplace_back(d(rng()));
                }
                return egtools::core::output(
                    egtools::core::makeArray((ExcelArrayBuilder::row_t)rows,
                                             (ExcelArrayBuilder::col_t)cols, vals));
            });
    }
}
