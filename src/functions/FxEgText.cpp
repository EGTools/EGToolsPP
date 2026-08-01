// FxEgText.cpp — EGTools 전용 텍스트 함수 (EGToolsVB FX2_Text.vb 포팅, plan/20 E1).
//   TEXTNUMSORT — 텍스트·숫자 혼합 자연 정렬(natural sort).
//   TRIMENDS    — 양끝 공백류 제거(제어문자/nbsp/보이지 않는 유니코드 옵션).
//   STREXT      — 문자 종류(숫자/영문/한글/일본어/한자)별 추출·제거, 정규식 허용.
//   TEXTREPLACE — StartKey~EndKey 구간을 대체.
//   TEXTBETWEEN — StartKey~EndKey 사이 문자열 추출(연결 또는 N번째).
//   HANTONUMBER — 한글(한자) 금액 표기 → 숫자.
// 전부 순수 계산(비매크로형).

#include "../core/Registry.h"
#include "../core/Apply.h"
#include "../core/ArrayUtil.h"
#include "../core/Spill.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>

#include <algorithm>
#include <cwctype>
#include <numeric>
#include <regex>
#include <string>
#include <vector>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        std::wstring lowerW(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)std::towlower(c);
            return s;
        }

        std::wstring trimWs(const std::wstring& s)
        {
            const wchar_t* ws = L" \t\r\n";
            const size_t b = s.find_first_not_of(ws);
            if (b == std::wstring::npos) return {};
            return s.substr(b, s.find_last_not_of(ws) - b + 1);
        }

        // 대소문자 무시 접두/검색 (BMP 단순 소문자화 기준).
        bool startsWithNoCase(const std::wstring& s, size_t at, const std::wstring& key)
        {
            if (at + key.size() > s.size()) return false;
            for (size_t i = 0; i < key.size(); ++i)
                if (std::towlower(s[at + i]) != std::towlower(key[i])) return false;
            return true;
        }
        size_t findNoCaseFrom(const std::wstring& s, const std::wstring& key, size_t from)
        {
            if (key.empty()) return std::wstring::npos;
            for (size_t i = from; i + key.size() <= s.size(); ++i)
                if (startsWithNoCase(s, i, key)) return i;
            return std::wstring::npos;
        }

        // ---- TEXTNUMSORT ---------------------------------------------------

        struct Seg { bool isNum = false; double num = 0; std::wstring txt; };

        // "abc12.5x" → ["abc", 12.5, "x"]. '.'/','는 뒤가 숫자면 숫자의 일부.
        std::vector<Seg> splitTextNum(const std::wstring& s)
        {
            std::vector<Seg> out;
            if (s.empty()) { out.push_back({}); return out; }

            auto flush = [&](bool isNum, const std::wstring& cur)
            {
                Seg seg;
                seg.isNum = isNum;
                if (isNum)
                {
                    std::wstring t;
                    for (wchar_t c : cur) if (c != L',') t.push_back(c);
                    seg.num = wcstod(t.c_str(), nullptr);
                }
                else seg.txt = cur;
                out.push_back(std::move(seg));
            };

            std::wstring cur;
            bool numMode = iswdigit(s[0]) != 0;
            for (size_t i = 0; i < s.size(); ++i)
            {
                const wchar_t c = s[i];
                const bool digitish = iswdigit(c) ||
                    ((c == L'.' || c == L',') && i + 1 < s.size() && iswdigit(s[i + 1]) && numMode);
                if (digitish)
                {
                    if (!numMode) { flush(false, cur); cur.clear(); numMode = true; }
                    cur.push_back(c);
                }
                else
                {
                    if (numMode) { flush(true, cur); cur.clear(); numMode = false; }
                    cur.push_back(c);
                }
            }
            flush(numMode, cur);
            return out;
        }

        int compareSegs(const std::vector<Seg>& a, const std::vector<Seg>& b, bool ignoreText)
        {
            const size_t n = (std::min)(a.size(), b.size());
            for (size_t i = 0; i < n; ++i)
            {
                if (a[i].isNum && b[i].isNum)
                {
                    if (a[i].num < b[i].num) return -1;
                    if (a[i].num > b[i].num) return 1;
                }
                else if (!ignoreText && !a[i].isNum && !b[i].isNum)
                {
                    const int c = a[i].txt.compare(b[i].txt);
                    if (c != 0) return c < 0 ? -1 : 1;
                }
            }
            if (a.size() != b.size()) return a.size() < b.size() ? -1 : 1;
            return 0;
        }

        ExcelObj* textNumSort(const ExcelObj& arrObj, const ExcelObj& colObj,
                              const ExcelObj& orderObj, const ExcelObj& ignoreObj)
        {
            if (arrObj.isMissing()) return returnValue(CellError::Value);
            if (!arrObj.isType(ExcelType::Multi))
                return returnValue(ExcelObj(arrObj));   // 1x1은 그대로

            ExcelArray a(arrObj);
            const size_t nR = a.nRows(), nC = a.nCols();
            if (nR == 0 || nC == 0) return returnValue(CellError::Value);

            int col = colObj.isMissing() ? 1 : colObj.get<int>(1);
            if (col == 0) col = 1;
            if (col < 1 || (size_t)col > nC) return returnValue(CellError::Value);
            int order = orderObj.isMissing() ? 1 : orderObj.get<int>(1);
            if (order == 0) order = 1;
            if (order != 1 && order != -1) return returnValue(CellError::Value);
            const bool ignoreText = ignoreObj.isMissing() ? false
                                                          : (ignoreObj.get<double>(0.0) != 0.0);

            std::vector<std::vector<Seg>> keys(nR);
            for (size_t r = 0; r < nR; ++r)
                keys[r] = splitTextNum(a.at(r * nC + (size_t)col - 1).toString());

            std::vector<size_t> idx(nR);
            std::iota(idx.begin(), idx.end(), (size_t)0);
            std::stable_sort(idx.begin(), idx.end(), [&](size_t x, size_t y)
            {
                const int c = (order == 1) ? compareSegs(keys[x], keys[y], ignoreText)
                                           : compareSegs(keys[y], keys[x], ignoreText);
                return c < 0;
            });

            std::vector<ExcelObj> vals;
            vals.reserve(nR * nC);
            for (size_t r = 0; r < nR; ++r)
                for (size_t c = 0; c < nC; ++c)
                    vals.push_back(ExcelObj(a.at(idx[r] * nC + c)));
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)nR, (ExcelArrayBuilder::col_t)nC, vals));
        }

        // ---- TRIMENDS ------------------------------------------------------

        bool inTrimSet(wchar_t c, int opt)
        {
            if (c == L' ') return true;
            if ((opt & 1) && c < 32) return true;
            if ((opt & 2) && c == 0x00A0) return true;
            if (opt & 4)
            {
                if (c >= 0x2000 && c <= 0x200B) return true;              // 각종 폭 공백·ZWSP
                if (c == 0x2028 || c == 0x2029) return true;              // 줄/문단 구분자
                if (c == 0x202F || c == 0x205F || c == 0x3000) return true; // NNBSP·MMSP·전각 공백
            }
            return false;
        }

        ExcelObj trimEndsOne(const ExcelObj& e, int opt)
        {
            if (e.type() == ExcelType::Missing || e.type() == ExcelType::Nil)
                return ExcelObj(std::wstring_view(L""));
            std::wstring s = trimWs(e.toString());
            size_t b = 0, t = s.size();
            while (b < t && inTrimSet(s[b], opt)) ++b;
            while (t > b && inTrimSet(s[t - 1], opt)) --t;
            return ExcelObj(std::wstring_view(s.data() + b, t - b));
        }

        // ---- STREXT --------------------------------------------------------

        ExcelObj* strExt(const ExcelObj& textObj, const ExcelObj& optObj,
                         const ExcelObj& delimObj)
        {
            if (textObj.isMissing() || optObj.isMissing())
                return returnValue(CellError::Value);
            const std::wstring text = textObj.toString();
            std::wstring opt = trimWs(optObj.toString());
            if (opt.empty()) return returnValue(CellError::Value);
            const std::wstring delim = delimObj.isMissing() ? L"" : delimObj.toString();

            // 고정 옵션(-SANKJH만으로 구성)인지 판별.
            static const std::wstring kFixed = L"-SANKJH";
            bool isFixed = true;
            for (wchar_t c : opt)
                if (kFixed.find((wchar_t)std::towupper(c)) == std::wstring::npos)
                { isFixed = false; break; }

            std::wstring pattern;
            bool remove = false;
            if (isFixed)
            {
                remove = !opt.empty() && opt[0] == L'-';
                if (remove) opt = opt.substr(1);
                for (auto& c : opt) c = (wchar_t)std::towupper(c);
                std::wstring cls;
                if (opt.find(L'S') != std::wstring::npos) cls += L"\\s";
                if (opt.find(L'A') != std::wstring::npos) cls += L"A-Za-z";
                if (opt.find(L'N') != std::wstring::npos) cls += L"0-9.";
                if (opt.find(L'K') != std::wstring::npos) cls += L"\\x{3130}-\\x{318F}\\x{AC00}-\\x{D7AF}";
                if (opt.find(L'J') != std::wstring::npos) cls += L"\\x{3040}-\\x{30FF}\\x{31F0}-\\x{31FF}";
                if (opt.find(L'H') != std::wstring::npos) cls += L"\\x{2E80}-\\x{2EFF}\\x{3400}-\\x{4DBF}\\x{4E00}-\\x{9FFF}\\x{F900}-\\x{FAFF}";
                if (cls.empty()) return returnValue(CellError::Value);
                pattern = L"[" + cls + L"]+";
            }
            else
                pattern = opt;   // 정규식 그대로

            try
            {
                // std::wregex는 \x{...} 확장을 지원하지 않으므로 실제 문자로 치환해 삽입.
                std::wstring rx;
                for (size_t i = 0; i < pattern.size(); )
                {
                    if (pattern.compare(i, 3, L"\\x{") == 0)
                    {
                        const size_t e = pattern.find(L'}', i + 3);
                        if (e != std::wstring::npos)
                        {
                            const unsigned long cp =
                                wcstoul(pattern.substr(i + 3, e - i - 3).c_str(), nullptr, 16);
                            rx.push_back((wchar_t)cp);
                            i = e + 1;
                            continue;
                        }
                    }
                    rx.push_back(pattern[i++]);
                }

                std::wregex re(rx, std::regex_constants::ECMAScript);
                if (remove)
                    return returnValue(ExcelObj(std::wstring_view(
                        std::regex_replace(text, re, L""))));

                std::wstring out;
                bool first = true;
                for (std::wsregex_iterator it(text.begin(), text.end(), re), end;
                     it != end; ++it)
                {
                    if (!first) out += delim;
                    out += it->str();
                    first = false;
                }
                return returnValue(ExcelObj(std::wstring_view(out)));
            }
            catch (...) { return returnValue(CellError::Value); }
        }

        // ---- TEXTREPLACE / TEXTBETWEEN -------------------------------------

        // StartKey~EndKey 구간(키 포함)을 찾아, 구간 밖 텍스트 조각 목록과
        // 구간 안 텍스트 목록을 만든다. 대소문자 무시.
        void splitByKeys(const std::wstring& text, const std::wstring& sk,
                         const std::wstring& ek,
                         std::vector<std::wstring>& outside,
                         std::vector<std::wstring>& inside)
        {
            std::wstring cur;
            size_t i = 0;
            while (i < text.size())
            {
                if (startsWithNoCase(text, i, sk))
                {
                    const size_t e = findNoCaseFrom(text, ek, i + sk.size());
                    if (e != std::wstring::npos)
                    {
                        outside.push_back(cur);
                        cur.clear();
                        inside.push_back(text.substr(i + sk.size(), e - i - sk.size()));
                        i = e + ek.size();
                        continue;
                    }
                    // 끝 키가 없으면: 남은 텍스트(시작 키 제외)를 별도 조각으로 (VB 동작).
                    outside.push_back(cur);
                    cur.clear();
                    outside.push_back(text.substr(i + sk.size()));
                    return;
                }
                cur.push_back(text[i]);
                ++i;
            }
            if (!cur.empty()) outside.push_back(cur);
        }

        ExcelObj* textReplace(const ExcelObj& textObj, const ExcelObj& skObj,
                              const ExcelObj& ekObj, const ExcelObj& repObj,
                              const ExcelObj& incObj)
        {
            const std::wstring text = textObj.isMissing() ? L"" : textObj.toString();
            const std::wstring sk = skObj.isMissing() ? L"" : skObj.toString();
            const std::wstring ek = ekObj.isMissing() ? L"" : ekObj.toString();
            if (text.empty() || sk.empty() || ek.empty())
                return returnValue(ExcelObj(std::wstring_view(text)));
            std::wstring rep = repObj.isMissing() ? L"" : repObj.toString();
            const bool include = incObj.isMissing() ? false : (incObj.get<double>(0.0) != 0.0);
            if (include) rep = sk + rep + ek;

            std::vector<std::wstring> outside, inside;
            splitByKeys(text, sk, ek, outside, inside);

            std::wstring out;
            for (size_t i = 0; i < outside.size(); ++i)
            {
                if (i) out += rep;
                out += outside[i];
            }
            return returnValue(ExcelObj(std::wstring_view(out)));
        }

        ExcelObj* textBetween(const ExcelObj& textObj, const ExcelObj& skObj,
                              const ExcelObj& ekObj, const ExcelObj& delimObj,
                              const ExcelObj& incObj)
        {
            const std::wstring text = textObj.isMissing() ? L"" : textObj.toString();
            const std::wstring sk = skObj.isMissing() ? L"" : skObj.toString();
            const std::wstring ek = ekObj.isMissing() ? L"" : ekObj.toString();
            if (text.empty() || sk.empty() || ek.empty())
                return returnValue(CellError::Value);
            const bool include = incObj.isMissing() ? false : (incObj.get<double>(0.0) != 0.0);

            std::vector<std::wstring> outside, inside;
            splitByKeys(text, sk, ek, outside, inside);
            if (inside.empty()) return returnValue(ExcelObj(std::wstring_view(L"")));
            if (include)
                for (auto& s : inside) s = sk + s + ek;

            // 구분자 인수: 숫자면 N번째 하나만, 텍스트면 연결자.
            if (!delimObj.isMissing() &&
                (delimObj.type() == ExcelType::Num || delimObj.type() == ExcelType::Int))
            {
                const int n = delimObj.get<int>(0);
                if (n < 1 || (size_t)n > inside.size()) return returnValue(CellError::Value);
                return returnValue(ExcelObj(std::wstring_view(inside[(size_t)n - 1])));
            }
            const std::wstring delim = delimObj.isMissing() ? L"," : delimObj.toString();
            std::wstring out;
            for (size_t i = 0; i < inside.size(); ++i)
            {
                if (i) out += delim;
                out += inside[i];
            }
            return returnValue(ExcelObj(std::wstring_view(out)));
        }

        // ---- HANTONUMBER ---------------------------------------------------

        // 한자 이체자 → 한글 단위/숫자 정규화.
        wchar_t hanNormalize(wchar_t c)
        {
            switch (c)
            {
            case L'零': return L'0';
            case L'영': return L'0';
            case L'十': case L'拾': return L'십';
            case L'百': case L'佰': case L'陌': return L'백';
            case L'千': case L'阡': case L'仟': return L'천';
            case L'萬': case L'万': return L'만';
            case L'億': return L'억';
            case L'兆': return L'조';
            case L'京': return L'경';
            case L'垓': return L'해';
            default: return c;
            }
        }

        int hanDigit(wchar_t c)
        {
            static const std::wstring kor = L"일이삼사오육칠팔구";
            static const std::wstring hanja = L"一二三四五六七八九";
            static const std::wstring alt = L"壹貳參肆伍陸柒捌玖";
            size_t p = kor.find(c);
            if (p == std::wstring::npos) p = hanja.find(c);
            if (p == std::wstring::npos) p = alt.find(c);
            return p == std::wstring::npos ? -1 : (int)(p + 1);
        }

        // 만 미만 구간("천이백삼십사" 등) → 수치.
        double hanSection(const std::wstring& s)
        {
            double unit = 1, digit = 0;
            for (size_t i = s.size(); i >= 1; --i)
            {
                const wchar_t d = s[i - 1];
                if (iswdigit(d))
                {
                    digit += (double)(d - L'0') * unit;
                    unit = 1;
                }
                else if (d == L'십') { if (unit > 1) digit += unit; unit = 10; }
                else if (d == L'백') { if (unit > 1) digit += unit; unit = 100; }
                else if (d == L'천') { if (unit > 1) digit += unit; unit = 1000; }
            }
            if (unit > 1) digit += unit;
            return digit;
        }

        std::wstring groupThousands(const std::wstring& intDigits)
        {
            std::wstring out;
            const size_t n = intDigits.size();
            for (size_t i = 0; i < n; ++i)
            {
                if (i && (n - i) % 3 == 0) out += L',';
                out += intDigits[i];
            }
            return out;
        }

        ExcelObj* hanToNumber(const ExcelObj& textObj, const ExcelObj& numOnlyObj)
        {
            if (textObj.isMissing()) return returnValue(CellError::Value);
            std::wstring s = textObj.toString();
            if (s.empty()) return returnValue(CellError::Value);
            const bool numberOnly = numOnlyObj.isMissing() ? true
                                                           : (numOnlyObj.get<double>(1.0) != 0.0);

            // 아라비아 숫자(또는 .)가 2개 이상 연속이면 금액 표기가 아님.
            for (size_t i = 0; i + 1 < s.size(); ++i)
            {
                auto isNumCh = [](wchar_t c) { return iswdigit(c) || c == L'.'; };
                if (isNumCh(s[i]) && isNumCh(s[i + 1])) return returnValue(CellError::Value);
            }

            std::wstring pre, post;
            if (s.find(L"일금") != std::wstring::npos) pre = L"일금";
            else if (s.find(L'금') != std::wstring::npos) pre = L"금";
            if (s.find(L"원정") != std::wstring::npos) post = L"원정";
            else if (s.find(L'원') != std::wstring::npos) post = L"원";

            auto erase = [&](const std::wstring& key)
            {
                size_t p;
                while ((p = s.find(key)) != std::wstring::npos) s.erase(p, key.size());
            };
            if (!pre.empty()) erase(pre);
            if (!post.empty()) erase(post);
            erase(L" ");
            erase(L",");

            for (auto& c : s)
            {
                c = hanNormalize(c);
                const int d = hanDigit(c);
                if (d > 0) c = (wchar_t)(L'0' + d);
            }

            // 큰 단위로 절단: 경/조/억/만 + 소수부.
            auto cutUnit = [&](wchar_t u) -> std::wstring
            {
                const size_t p = s.find(u);
                if (p == std::wstring::npos) return L"";
                std::wstring head = s.substr(0, p);
                s = s.substr(p + 1);
                return head.empty() ? L"1" : head;
            };
            const std::wstring vGyeong = cutUnit(L'경');
            const std::wstring vJo = cutUnit(L'조');
            const std::wstring vEok = cutUnit(L'억');
            const std::wstring vMan = cutUnit(L'만');
            std::wstring vDec;
            {
                const size_t p = s.find(L'.');
                if (p != std::wstring::npos) { vDec = s.substr(p + 1); s = s.substr(0, p); }
            }
            const std::wstring vOne = s;

            double result = 0;
            const std::wstring* parts[5] = { &vOne, &vMan, &vEok, &vJo, &vGyeong };
            for (int x = 0; x < 5; ++x)
                if (!parts[x]->empty())
                    result += hanSection(*parts[x]) * pow(10.0, 4.0 * x);
            if (!vDec.empty())
            {
                const double d = hanSection(vDec);
                wchar_t buf[32];
                swprintf_s(buf, L"%.0f", d);
                result += d * pow(10.0, -(double)wcslen(buf));
            }

            // 15자리 초과 정수부는 double 정밀도 밖 → 텍스트로.
            wchar_t intBuf[64];
            swprintf_s(intBuf, L"%.0f", fabs(result));
            const bool big = wcslen(intBuf) > 15;

            if (numberOnly)
            {
                if (big)
                {
                    wchar_t buf[64];
                    swprintf_s(buf, L"%.0f", result);
                    return returnValue(ExcelObj(std::wstring_view(buf)));
                }
                return returnValue(ExcelObj(result));
            }

            // "금1,234원" 형태.
            wchar_t buf[64];
            swprintf_s(buf, L"%.10g", result);
            std::wstring numTxt = buf;
            const size_t dot = numTxt.find(L'.');
            std::wstring intPart = (dot == std::wstring::npos) ? numTxt : numTxt.substr(0, dot);
            const std::wstring frac = (dot == std::wstring::npos) ? L"" : numTxt.substr(dot);
            const std::wstring formatted = big ? intPart + frac
                                              : groupThousands(intPart) + frac;
            return returnValue(ExcelObj(std::wstring_view(pre + formatted + post)));
        }
    }

    void registerEgText()
    {
        egtools::core::registerFn(L"TEXTNUMSORT",
            [](const ExcelObj& arr, const ExcelObj& col, const ExcelObj& order,
               const ExcelObj& ignore) -> ExcelObj*
            { return textNumSort(arr, col, order, ignore); });

        egtools::core::registerFn(L"TRIMENDS",
            [](const ExcelObj& text, const ExcelObj& optObj) -> ExcelObj*
            {
                const int opt = optObj.isMissing() ? 0 : optObj.get<int>(0);
                if (opt < 0 || opt > 7) return returnValue(CellError::Value);
                return egtools::core::mapUnary(text,
                    [opt](const ExcelObj& e) { return trimEndsOne(e, opt); });
            });

        egtools::core::registerFn(L"STREXT",
            [](const ExcelObj& text, const ExcelObj& opt, const ExcelObj& delim) -> ExcelObj*
            { return strExt(text, opt, delim); });

        egtools::core::registerFn(L"TEXTREPLACE",
            [](const ExcelObj& text, const ExcelObj& sk, const ExcelObj& ek,
               const ExcelObj& rep, const ExcelObj& inc) -> ExcelObj*
            { return textReplace(text, sk, ek, rep, inc); });

        egtools::core::registerFn(L"TEXTBETWEEN",
            [](const ExcelObj& text, const ExcelObj& sk, const ExcelObj& ek,
               const ExcelObj& delim, const ExcelObj& inc) -> ExcelObj*
            { return textBetween(text, sk, ek, delim, inc); });

        egtools::core::registerFn(L"HANTONUMBER",
            [](const ExcelObj& text, const ExcelObj& numOnly) -> ExcelObj*
            {
                try { return hanToNumber(text, numOnly); }
                catch (...) { return returnValue(CellError::Value); }
            });
    }
}
