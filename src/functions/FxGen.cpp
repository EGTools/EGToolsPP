// FxGen.cpp — array-generating functions (SEQUENCE, RANDARRAY).
//
// 인수 배열 → 강등 리프팅(네이티브 정합, plan/22 T4·U16): 원소별 결과의
// 첫 값으로 강등된 입력 모양 배열.

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "../core/Apply.h"

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
            // MTR(멀티스레드 재계산) 참여를 위해 스레드별 생성기 사용.
            thread_local std::mt19937 g{ std::random_device{}() };
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
                return egtools::core::mapLift(
                    [](const ExcelObj& re, const ExcelObj& ce,
                       const ExcelObj& se, const ExcelObj& pe) -> ExcelObj
                {
                    const int rows = argInt(re, 1);
                    const int cols = argInt(ce, 1);
                    if (rows <= 0 || cols <= 0) return ExcelObj(CellError::Value);
                    const double start = argDbl(se, 1.0);
                    const double step = argDbl(pe, 1.0);

                    std::vector<ExcelObj> vals;
                    vals.reserve((size_t)rows * cols);
                    double v = start;
                    for (int i = 0; i < rows * cols; ++i) { vals.emplace_back(v); v += step; }
                    return egtools::core::makeArray((ExcelArrayBuilder::row_t)rows,
                                                    (ExcelArrayBuilder::col_t)cols, vals);
                }, rowsA, colsA, startA, stepA);
            });

        // RANDARRAY([rows], [cols], [min], [max], [integer]).
        egtools::core::registerFn(L"RANDARRAY",
            [](const ExcelObj& rowsA, const ExcelObj& colsA, const ExcelObj& minA,
               const ExcelObj& maxA, const ExcelObj& intA) -> ExcelObj*
            {
                return egtools::core::mapLift(
                    [](const ExcelObj& re, const ExcelObj& ce, const ExcelObj& le,
                       const ExcelObj& he, const ExcelObj& ie) -> ExcelObj
                {
                    const int rows = argInt(re, 1);
                    const int cols = argInt(ce, 1);
                    if (rows <= 0 || cols <= 0) return ExcelObj(CellError::Value);
                    const double lo = argDbl(le, 0.0);
                    const double hi = argDbl(he, 1.0);
                    const bool asInt = ie.isMissing() ? false : (ie.get<double>(0.0) != 0.0);
                    if (hi < lo) return ExcelObj(CellError::Value);

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
                    return egtools::core::makeArray((ExcelArrayBuilder::row_t)rows,
                                                    (ExcelArrayBuilder::col_t)cols, vals);
                }, rowsA, colsA, minA, maxA, intA);
            });
    }
}
