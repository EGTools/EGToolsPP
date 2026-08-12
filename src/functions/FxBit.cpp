// FxBit.cpp — bitwise functions (Excel 2013): BITAND/OR/XOR/LSHIFT/RSHIFT.
//
// Excel limits operands to integers in [0, 2^48-1] and results to < 2^53.
// 전 인수 원소별 리프팅(외적 브로드캐스트, 네이티브 정합 — plan/22 U07·U08).

#include "../core/Registry.h"
#include "../core/Apply.h"

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
                    return egtools::core::mapLift(
                        [op](const ExcelObj& ae, const ExcelObj& be) -> ExcelObj
                        {
                            unsigned long long a, b;
                            if (!asULL(ae, a) || !asULL(be, b)) return ExcelObj(CellError::Num);
                            return ExcelObj((double)op(a, b));
                        }, aA, bA);
                });
        };
        binOp(L"BITAND", [](unsigned long long a, unsigned long long b) { return a & b; });
        binOp(L"BITOR",  [](unsigned long long a, unsigned long long b) { return a | b; });
        binOp(L"BITXOR", [](unsigned long long a, unsigned long long b) { return a ^ b; });

        auto shiftOp = [](const wchar_t* name, bool leftIsPositive)
        {
            egtools::core::registerFn(name,
                [leftIsPositive](const ExcelObj& numA, const ExcelObj& shiftA) -> ExcelObj*
                {
                    return egtools::core::mapLift(
                        [leftIsPositive](const ExcelObj& ne, const ExcelObj& se) -> ExcelObj
                        {
                            unsigned long long n;
                            if (!asULL(ne, n)) return ExcelObj(CellError::Num);
                            auto t = se.type();
                            if (!se.isMissing() &&
                                !(t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool))
                                return ExcelObj(CellError::Num);
                            const double sd = se.isMissing() ? 0.0 : se.get<double>(0.0);
                            if (std::floor(sd) != sd || std::fabs(sd) > 53)
                                return ExcelObj(CellError::Num);
                            const int s = (int)sd;
                            const bool left = leftIsPositive ? (s >= 0) : (s < 0);
                            const int mag = std::abs(s);
                            const double r = left
                                ? (double)n * std::pow(2.0, mag)
                                : std::floor((double)n / std::pow(2.0, mag));
                            if (r > kMaxResult) return ExcelObj(CellError::Num);
                            return ExcelObj(r);
                        }, numA, shiftA);
                });
        };
        // BITLSHIFT(number, shift): shift left; negative shift → right (반대는 RSHIFT).
        shiftOp(L"BITLSHIFT", true);
        shiftOp(L"BITRSHIFT", false);
    }
}
