// FxAgg.cpp — conditional aggregation / selection (IFS, SWITCH, MAXIFS, MINIFS).
//
// IFS의 조건·값, SWITCH의 식·비교값·결과, MAXIFS/MINIFS의 조건값은 원소별
// 리프팅(브로드캐스트) 대상이다(네이티브 정합, plan/22 U09~U11). 범위 인수
// (MAXIFS의 최대범위/조건범위)는 집계 데이터라 리프팅하지 않는다.

#include "../core/Registry.h"
#include "../core/Apply.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <string>
#include <vector>
#include <cwctype>
#include <cwchar>
#include <limits>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        bool truthy(const ExcelObj& v, bool& isErr)
        {
            switch (v.type())
            {
            case ExcelType::Err:  isErr = true; return false;
            case ExcelType::Bool:
            case ExcelType::Num:
            case ExcelType::Int:  return v.get<double>(0.0) != 0.0;
            default:              return false;
            }
        }

        // Case-insensitive wildcard match (* = any run, ? = any one char).
        bool wildcard(const std::wstring& text, const std::wstring& pat)
        {
            size_t t = 0, p = 0, star = std::wstring::npos, mark = 0;
            auto low = [](wchar_t c) { return (wchar_t)std::towlower(c); };
            while (t < text.size())
            {
                if (p < pat.size() && (pat[p] == L'?' || low(pat[p]) == low(text[t])))
                    { ++t; ++p; }
                else if (p < pat.size() && pat[p] == L'*')
                    { star = p++; mark = t; }
                else if (star != std::wstring::npos)
                    { p = star + 1; t = ++mark; }
                else return false;
            }
            while (p < pat.size() && pat[p] == L'*') ++p;
            return p == pat.size();
        }

        bool asNumber(const ExcelObj& v, double& out)
        {
            auto t = v.type();
            if (t == ExcelType::Num || t == ExcelType::Int || t == ExcelType::Bool)
                { out = v.get<double>(0.0); return true; }
            return false;
        }

        // Excel-style criteria match: ">5", "<=3", "<>x", "ab*", or a plain value.
        bool criteriaMatch(const ExcelObj& cell, const ExcelObj& crit)
        {
            double cn;
            if (asNumber(crit, cn))   // numeric criterion → numeric equality
            {
                double v;
                return asNumber(cell, v) && v == cn;
            }
            std::wstring s = crit.toString();
            std::wstring op = L"=";
            if (s.rfind(L">=", 0) == 0) { op = L">="; s = s.substr(2); }
            else if (s.rfind(L"<=", 0) == 0) { op = L"<="; s = s.substr(2); }
            else if (s.rfind(L"<>", 0) == 0) { op = L"<>"; s = s.substr(2); }
            else if (s.rfind(L">", 0) == 0) { op = L">"; s = s.substr(1); }
            else if (s.rfind(L"<", 0) == 0) { op = L"<"; s = s.substr(1); }
            else if (s.rfind(L"=", 0) == 0) { op = L"="; s = s.substr(1); }

            wchar_t* end = nullptr;
            double operand = std::wcstod(s.c_str(), &end);
            const bool numericOperand = (!s.empty() && end && *end == L'\0');
            double cv;
            if (numericOperand && asNumber(cell, cv))
            {
                if (op == L"=")  return cv == operand;
                if (op == L"<>") return cv != operand;
                if (op == L">")  return cv > operand;
                if (op == L"<")  return cv < operand;
                if (op == L">=") return cv >= operand;
                if (op == L"<=") return cv <= operand;
            }
            // text comparison
            const std::wstring cs = cell.toString();
            if (op == L"=")  return wildcard(cs, s);
            if (op == L"<>") return !wildcard(cs, s);
            int cmp = _wcsicmp(cs.c_str(), s.c_str());
            if (op == L">")  return cmp > 0;
            if (op == L"<")  return cmp < 0;
            if (op == L">=") return cmp >= 0;
            if (op == L"<=") return cmp <= 0;
            return false;
        }

        // Scalar core: aggregate with the given (already non-missing) criteria.
        ExcelObj maxminIfs(bool isMax, const ExcelObj& vrange,
                           const ExcelObj* const critRange[],
                           const ExcelObj* const crit[], int pairs)
        {
            try
            {
                ExcelArray vr(vrange);
                const size_t N = (size_t)vr.nRows() * vr.nCols();
                std::vector<ExcelArray> crs;
                std::vector<const ExcelObj*> cvs;
                for (int k = 0; k < pairs; ++k)
                {
                    ExcelArray ca(*critRange[k]);
                    if ((size_t)ca.nRows() * ca.nCols() != N) return ExcelObj(CellError::Value);
                    crs.push_back(ca);
                    cvs.push_back(crit[k]);
                }
                double best = isMax ? -std::numeric_limits<double>::infinity()
                                    :  std::numeric_limits<double>::infinity();
                bool found = false;
                for (size_t i = 0; i < N; ++i)
                {
                    bool all = true;
                    for (size_t c = 0; c < crs.size(); ++c)
                        if (!criteriaMatch(crs[c].at(i), *cvs[c])) { all = false; break; }
                    if (!all) continue;
                    // 조건을 통과한 행의 값이 오류면 그 오류 전파(네이티브 실측).
                    if (vr.at(i).type() == ExcelType::Err)
                        return ExcelObj(vr.at(i));
                    double v;
                    if (asNumber(vr.at(i), v))
                    {
                        found = true;
                        best = isMax ? (v > best ? v : best) : (v < best ? v : best);
                    }
                }
                return ExcelObj(found ? best : 0.0);
            }
            catch (...) { return ExcelObj(CellError::Value); }
        }
    }

    void registerAgg()
    {
        // IFS(cond1, val1, cond2, val2, …) — variadic up to 254 args (127 pairs).
        // 조건·값 전 인수가 원소별 리프팅된다: =IFS(A1:A5>3,"큰",TRUE,"작은") →
        // 5개 배열(네이티브 정합, plan/22 U09).
        egtools::core::registerRawFn(L"IFS",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                // 사용되지 않은 인수의 오류는 전파하지 않음(네이티브) → PassThrough
                // + 내부 규칙: 평가된 조건이 오류면 그 오류를 반환.
                return egtools::core::mapLiftN(args, info.numArgs(),
                    [](const ExcelObj* const* e, size_t n) -> ExcelObj
                    {
                        for (size_t i = 0; i + 1 < n; i += 2)
                        {
                            if (e[i]->isMissing()) break;
                            bool err = false;
                            if (truthy(*e[i], err)) return ExcelObj(*e[i + 1]);
                            if (err) return ExcelObj(*e[i]);
                        }
                        return ExcelObj(CellError::NA);
                    }, egtools::core::LiftErrors::PassThrough);
            });

        // SWITCH(expr, val1, res1, …, [default]) — variadic. 식/비교값/결과가
        // 원소별 리프팅된다(plan/22 U10).
        egtools::core::registerRawFn(L"SWITCH",
            [](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                const size_t total = info.numArgs();
                if (total == 0) return returnValue(CellError::NA);
                // 사용되지 않은 인수의 오류는 전파하지 않음(네이티브) → PassThrough
                // + 내부 규칙: 식/평가된 비교값이 오류면 그 오류를 반환.
                return egtools::core::mapLiftN(args, total,
                    [](const ExcelObj* const* e, size_t tot) -> ExcelObj
                    {
                        const ExcelObj& expr = *e[0];
                        if (expr.type() == ExcelType::Err) return ExcelObj(expr);
                        // Length of the value/result list a[] = e[1..].
                        size_t aN = 0; while (1 + aN < tot && !e[1 + aN]->isMissing()) ++aN;
                        size_t j = 0;
                        for (; j + 1 < aN; j += 2)
                        {
                            if (e[1 + j]->type() == ExcelType::Err)
                                return ExcelObj(*e[1 + j]);
                            if (ExcelObj::compare(expr, *e[1 + j]) == 0)
                                return ExcelObj(*e[1 + j + 1]);
                        }
                        if (j < aN) return ExcelObj(*e[1 + j]);  // trailing default
                        return ExcelObj(CellError::NA);
                    }, egtools::core::LiftErrors::PassThrough);
            });

        // MAXIFS / MINIFS(range, crit_range1, crit1, [crit_range2, crit2, …]) —
        // variadic. 조건값(crit_i)만 리프팅한다: =MAXIFS(A:A,B:B,{">1",">2"}) →
        // 1×2 배열(plan/22 U11, 범위 인수는 집계 데이터).
        auto maxminFn = [](bool isMax) {
            return [isMax](const FuncInfo& info, const ExcelObj** args) -> ExcelObj*
            {
                const size_t n = info.numArgs();
                if (n < 1) return returnValue(CellError::Value);
                std::vector<const ExcelObj*> crs, cvs;
                for (size_t i = 1; i + 1 < n; i += 2)
                {
                    if (args[i]->isMissing() || args[i + 1]->isMissing()) continue;
                    crs.push_back(args[i]);
                    cvs.push_back(args[i + 1]);
                }
                const ExcelObj& vrange = *args[0];
                // 조건값 오류 = 불일치(네이티브 결과 0) → PassThrough.
                return egtools::core::mapLiftN(cvs.data(), cvs.size(),
                    [&](const ExcelObj* const* ce, size_t pairs) -> ExcelObj
                    {
                        return maxminIfs(isMax, vrange, crs.data(), ce, (int)pairs);
                    }, egtools::core::LiftErrors::PassThrough);
            };
        };
        egtools::core::registerRawFn(L"MAXIFS", maxminFn(true));
        egtools::core::registerRawFn(L"MINIFS", maxminFn(false));
    }
}
