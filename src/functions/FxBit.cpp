// FxBit.cpp — bitwise functions (Excel 2013): BITAND/OR/XOR/LSHIFT/RSHIFT.
//
// Excel limits operands to integers in [0, 2^48-1] and results to < 2^53.

#include "../core/Registry.h"

#include <xlOil/xlOil.h>
#include <cmath>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        constexpr double kMaxOperand = 281474976710655.0;   // 2^48 - 1
        constexpr double kMaxResult  = 9007199254740992.0;  // 2^53

        bool asULL(const ExcelObj& o, unsigned long long& out)
        {
            auto t = o.type();
            if (!(t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool))
                return false;
            double d = o.get<double>(-1.0);
            if (d < 0 || std::floor(d) != d || d > kMaxOperand) return false;
            out = (unsigned long long)d;
            return true;
        }
    }

    void registerBit()
    {
        auto binOp = [](const wchar_t* name, auto op)
        {
            egtools::core::registerFn(name,
                [op](const ExcelObj& aA, const ExcelObj& bA) -> ExcelObj*
                {
                    unsigned long long a, b;
                    if (!asULL(aA, a) || !asULL(bA, b)) return returnValue(CellError::Num);
                    return returnValue(ExcelObj((double)op(a, b)));
                });
        };
        binOp(L"BITAND", [](unsigned long long a, unsigned long long b) { return a & b; });
        binOp(L"BITOR",  [](unsigned long long a, unsigned long long b) { return a | b; });
        binOp(L"BITXOR", [](unsigned long long a, unsigned long long b) { return a ^ b; });

        // BITLSHIFT(number, shift_amount): shift left; negative shift → right.
        egtools::core::registerFn(L"BITLSHIFT",
            [](const ExcelObj& numA, const ExcelObj& shiftA) -> ExcelObj*
            {
                unsigned long long n; bool so;
                if (!asULL(numA, n)) return returnValue(CellError::Num);
                double sd = shiftA.get<double>(0.0); (void)so;
                if (std::floor(sd) != sd || std::fabs(sd) > 53) return returnValue(CellError::Num);
                int s = (int)sd;
                double r = (s >= 0) ? (double)n * std::pow(2.0, s)
                                    : std::floor((double)n / std::pow(2.0, -s));
                if (r > kMaxResult) return returnValue(CellError::Num);
                return returnValue(ExcelObj(r));
            });

        // BITRSHIFT(number, shift_amount): shift right; negative shift → left.
        egtools::core::registerFn(L"BITRSHIFT",
            [](const ExcelObj& numA, const ExcelObj& shiftA) -> ExcelObj*
            {
                unsigned long long n;
                if (!asULL(numA, n)) return returnValue(CellError::Num);
                double sd = shiftA.get<double>(0.0);
                if (std::floor(sd) != sd || std::fabs(sd) > 53) return returnValue(CellError::Num);
                int s = (int)sd;
                double r = (s >= 0) ? std::floor((double)n / std::pow(2.0, s))
                                    : (double)n * std::pow(2.0, -s);
                if (r > kMaxResult) return returnValue(CellError::Num);
                return returnValue(ExcelObj(r));
            });
    }
}
