// FxGoogle.cpp — Google Sheets 호환 함수 (EGToolsVB FXG_Google.vb 포팅, plan/20).
//   GOOGLETRANSLATE / GTRS — translate.google.com/m 페이지 스크래핑 번역.
//     비공식 엔드포인트: 형식 변경/실패는 VB판과 동일하게 "ERROR:…" 문자열로 반환.
//   (후속 배치: IMPORTDATA, IMPORTFEED, IMPORTRANGE, IMPORTHTML, QUERY)
//
// WinHTTP만 사용(Excel API 미접촉) — 매크로형 등록 불필요.

#include "../core/Registry.h"
#include "../core/I18n.h"
#include "../core/ArrayUtil.h"
#include "../core/Spill.h"

#include <windows.h>
#include <winhttp.h>
#include <oleauto.h>
#include <msxml6.h>
#include <mshtml.h>

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelUI.h>      // variantToExcelObj

#include <algorithm>
#include <cstdio>
#include <cwctype>
#include <fstream>
#include <string>
#include <vector>

#pragma comment(lib, "winhttp.lib")

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // ---- 공용 헬퍼 (후속 IMPORT* 배치에서도 사용) ----------------------

        // HTTP(S) GET — 원시 바이트를 돌려준다(문자셋 해석은 호출자 몫).
        // 리다이렉트는 WinHTTP 기본 정책으로 따라간다(https→http 제외).
        bool httpGet(const std::wstring& url, std::string& body)
        {
            URL_COMPONENTS uc{}; uc.dwStructSize = sizeof(uc);
            std::wstring host((size_t)url.size() + 1, L'\0');
            std::wstring path((size_t)url.size() + 1, L'\0');
            uc.lpszHostName = host.data(); uc.dwHostNameLength = (DWORD)url.size();
            uc.lpszUrlPath = path.data(); uc.dwUrlPathLength = (DWORD)url.size();
            if (!WinHttpCrackUrl(url.c_str(), (DWORD)url.size(), 0, &uc))
                return false;
            host.resize(uc.dwHostNameLength);

            HINTERNET hSession = WinHttpOpen(L"Mozilla/5.0 (Windows NT 10.0; Win64; x64) EGTools++",
                WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);
            if (!hSession) return false;
            WinHttpSetTimeouts(hSession, 5000, 5000, 8000, 15000);

            bool ok = false;
            HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), uc.nPort, 0);
            if (hConnect)
            {
                DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
                HINTERNET hReq = WinHttpOpenRequest(hConnect, L"GET", uc.lpszUrlPath, nullptr,
                    WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
                if (hReq)
                {
                    // 일부 사이트(CDN 봇 필터)는 Accept류 헤더가 없으면 거부한다.
                    static const wchar_t* kHeaders =
                        L"Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8\r\n"
                        L"Accept-Language: ko,en;q=0.8\r\n";
                    if (WinHttpSendRequest(hReq, kHeaders, (DWORD)-1,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
                        WinHttpReceiveResponse(hReq, nullptr))
                    {
                        DWORD avail = 0;
                        do
                        {
                            avail = 0;
                            if (!WinHttpQueryDataAvailable(hReq, &avail) || avail == 0) break;
                            std::string chunk((size_t)avail, '\0');
                            DWORD read = 0;
                            if (!WinHttpReadData(hReq, chunk.data(), avail, &read)) break;
                            body.append(chunk.data(), read);
                        } while (avail > 0);
                        ok = true;
                    }
                    WinHttpCloseHandle(hReq);
                }
                WinHttpCloseHandle(hConnect);
            }
            WinHttpCloseHandle(hSession);
            return ok;
        }

        std::wstring utf8ToWide(const std::string& s)
        {
            if (s.empty()) return {};
            int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
            std::wstring w((size_t)n, L'\0');
            MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), w.data(), n);
            return w;
        }

        // UTF-8 percent-encoding (RFC 3986 unreserved 문자만 그대로).
        std::wstring percentEncode(const std::wstring& w)
        {
            int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, nullptr);
            std::string u8((size_t)n, '\0');
            WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), u8.data(), n, nullptr, nullptr);

            static const char* kHex = "0123456789ABCDEF";
            std::wstring out;
            out.reserve(u8.size() * 3);
            for (unsigned char c : u8)
            {
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
                    (c >= '0' && c <= '9') || c == '-' || c == '_' || c == '.' || c == '~')
                    out.push_back((wchar_t)c);
                else
                {
                    out.push_back(L'%');
                    out.push_back((wchar_t)kHex[c >> 4]);
                    out.push_back((wchar_t)kHex[c & 0x0F]);
                }
            }
            return out;
        }

        std::wstring lowerCopy(std::wstring s)
        {
            for (auto& c : s) c = (wchar_t)std::towlower(c);
            return s;
        }

        std::wstring trimCopy(const std::wstring& s)
        {
            const wchar_t* ws = L" \t\r\n";
            const size_t b = s.find_first_not_of(ws);
            if (b == std::wstring::npos) return {};
            return s.substr(b, s.find_last_not_of(ws) - b + 1);
        }

        // 대소문자 무시 부분 문자열 검색.
        size_t findNoCase(const std::wstring& hay, const std::wstring& needle, size_t from = 0)
        {
            if (needle.empty() || hay.size() < needle.size()) return std::wstring::npos;
            const std::wstring h = lowerCopy(hay), n = lowerCopy(needle);
            return h.find(n, from);
        }

        // HTML 엔티티 디코드 — 이름형(주요 5종 + nbsp)과 수치형(&#nn; / &#xhh;).
        std::wstring decodeEntities(const std::wstring& s)
        {
            std::wstring out;
            out.reserve(s.size());
            for (size_t i = 0; i < s.size(); )
            {
                if (s[i] != L'&') { out.push_back(s[i++]); continue; }
                const size_t semi = s.find(L';', i + 1);
                if (semi == std::wstring::npos || semi - i > 10) { out.push_back(s[i++]); continue; }
                const std::wstring ent = s.substr(i + 1, semi - i - 1);
                if      (ent == L"amp")  out.push_back(L'&');
                else if (ent == L"lt")   out.push_back(L'<');
                else if (ent == L"gt")   out.push_back(L'>');
                else if (ent == L"quot") out.push_back(L'"');
                else if (ent == L"apos") out.push_back(L'\'');
                else if (ent == L"nbsp") out.push_back(L' ');
                else if (ent.size() > 1 && ent[0] == L'#')
                {
                    const bool hex = ent.size() > 2 && (ent[1] == L'x' || ent[1] == L'X');
                    unsigned long cp = wcstoul(ent.c_str() + (hex ? 2 : 1), nullptr, hex ? 16 : 10);
                    if (cp == 0 || cp > 0x10FFFF) { out.push_back(s[i++]); continue; }
                    if (cp <= 0xFFFF) out.push_back((wchar_t)cp);
                    else
                    {   // surrogate pair
                        cp -= 0x10000;
                        out.push_back((wchar_t)(0xD800 + (cp >> 10)));
                        out.push_back((wchar_t)(0xDC00 + (cp & 0x3FF)));
                    }
                }
                else { out.push_back(s[i++]); continue; }
                i = semi + 1;
            }
            return out;
        }

        // ---- 텍스트 데이터 공용 (IMPORTDATA / IMPORTRANGE) -----------------

        // charset 이름/코드페이지 문자열 → Windows 코드페이지. 0 = UTF-16LE,
        // 1 = UTF-16BE 특수 처리. 알 수 없으면 UTF-8.
        UINT codepageFromName(std::wstring cs, bool& utf16le, bool& utf16be)
        {
            utf16le = utf16be = false;
            cs = lowerCopy(trimCopy(cs));
            if (cs.empty()) return CP_UTF8;
            if (!cs.empty() && std::all_of(cs.begin(), cs.end(), ::iswdigit))
            {
                const unsigned long cp = wcstoul(cs.c_str(), nullptr, 10);
                if (cp == 1200) { utf16le = true; return 0; }
                if (cp == 1201) { utf16be = true; return 0; }
                return (UINT)cp;
            }
            if (cs == L"utf-8" || cs == L"utf8") return CP_UTF8;
            if (cs == L"utf-16" || cs == L"utf-16le") { utf16le = true; return 0; }
            if (cs == L"utf-16be") { utf16be = true; return 0; }
            if (cs == L"us-ascii" || cs == L"ascii") return 20127;
            if (cs == L"iso-8859-1" || cs == L"latin1") return 28591;
            if (cs == L"euc-kr" || cs == L"cp949" || cs == L"ks_c_5601-1987") return 949;
            if (cs == L"shift_jis" || cs == L"shift-jis" || cs == L"sjis") return 932;
            if (cs == L"gb2312" || cs == L"gbk") return 936;
            if (cs == L"big5") return 950;
            if (cs == L"windows-1252") return 1252;
            return CP_UTF8;
        }

        // 바이트 → 문자열. BOM이 있으면 BOM이 우선한다.
        std::wstring decodeBytes(const std::string& b, const std::wstring& charset)
        {
            size_t off = 0;
            bool le = false, be = false;
            UINT cp = codepageFromName(charset, le, be);
            if (b.size() >= 3 && (unsigned char)b[0] == 0xEF && (unsigned char)b[1] == 0xBB &&
                (unsigned char)b[2] == 0xBF) { cp = CP_UTF8; le = be = false; off = 3; }
            else if (b.size() >= 2 && (unsigned char)b[0] == 0xFF && (unsigned char)b[1] == 0xFE)
                { le = true; be = false; off = 2; }
            else if (b.size() >= 2 && (unsigned char)b[0] == 0xFE && (unsigned char)b[1] == 0xFF)
                { be = true; le = false; off = 2; }

            if (le || be)
            {
                std::wstring w;
                w.reserve((b.size() - off) / 2);
                for (size_t i = off; i + 1 < b.size(); i += 2)
                {
                    const unsigned char c0 = (unsigned char)b[i], c1 = (unsigned char)b[i + 1];
                    w.push_back((wchar_t)(le ? (c0 | (c1 << 8)) : ((c0 << 8) | c1)));
                }
                return w;
            }
            const char* p = b.data() + off;
            const int len = (int)(b.size() - off);
            if (len <= 0) return {};
            int n = MultiByteToWideChar(cp, 0, p, len, nullptr, 0);
            if (n <= 0) { cp = CP_UTF8; n = MultiByteToWideChar(cp, 0, p, len, nullptr, 0); }
            std::wstring w((size_t)n, L'\0');
            MultiByteToWideChar(cp, 0, p, len, w.data(), n);
            return w;
        }

        bool readFileBytes(const std::wstring& path, std::string& body)
        {
            std::ifstream f(path.c_str(), std::ios::binary);
            if (!f) return false;
            f.seekg(0, std::ios::end);
            const std::streamoff sz = f.tellg();
            if (sz < 0) return false;
            f.seekg(0);
            body.resize((size_t)sz);
            f.read(body.data(), sz);
            return (bool)f || sz == 0;
        }

        // \r\n / \n 분리. keepEmpty=false면 빈 줄 제거(IMPORTDATA),
        // true면 유지(IMPORTRANGE — 행 위치가 의미를 가짐).
        std::vector<std::wstring> splitLines(const std::wstring& text, bool keepEmpty = false)
        {
            std::vector<std::wstring> lines;
            std::wstring cur;
            for (wchar_t c : text)
            {
                if (c == L'\r') continue;
                if (c == L'\n')
                {
                    if (keepEmpty || !cur.empty()) lines.push_back(std::move(cur));
                    cur.clear();
                }
                else cur.push_back(c);
            }
            if (!cur.empty()) lines.push_back(std::move(cur));
            return lines;
        }

        std::vector<std::wstring> splitTabs(const std::wstring& line)
        {
            std::vector<std::wstring> cells;
            size_t b = 0;
            for (size_t i = 0; i <= line.size(); ++i)
                if (i == line.size() || line[i] == L'\t')
                {
                    cells.push_back(line.substr(b, i - b));
                    b = i + 1;
                }
            return cells;
        }

        // 숫자 전체 일치 파싱. 기본: 천단위 콤마 허용. decimalComma=true(로캘)면
        // 콤마=소수점, 마침표=천단위로 해석.
        bool tryParseNumber(const std::wstring& s, double& out, bool decimalComma = false)
        {
            std::wstring t;
            t.reserve(s.size());
            for (wchar_t c : s)
            {
                if (decimalComma)
                {
                    if (c == L'.') continue;
                    t.push_back(c == L',' ? L'.' : c);
                }
                else
                {
                    if (c == L',') continue;
                    t.push_back(c);
                }
            }
            if (t.empty()) return false;
            wchar_t* end = nullptr;
            const double v = wcstod(t.c_str(), &end);
            if (!end || *end != L'\0' || end == t.c_str()) return false;
            out = v;
            return true;
        }

        // 로캘 문자열("de-DE" 등)이 소수점 콤마 문화권인지.
        bool localeDecimalComma(const std::wstring& locale)
        {
            std::wstring lang;
            for (wchar_t c : locale)
            {
                if (c == L'-' || c == L'_') break;
                lang.push_back((wchar_t)std::towlower(c));
            }
            static const wchar_t* kComma[] = { L"de", L"fr", L"es", L"it", L"pt", L"nl",
                L"tr", L"ru", L"pl", L"sv", L"da", L"fi", L"nb", L"no", L"id", L"vi",
                L"cs", L"el", L"hu", L"ro", L"uk", L"bg", L"hr", L"sk", L"sl", L"sr", L"lt", L"lv", L"et" };
            for (auto* c : kComma)
                if (lang == c) return true;
            return false;
        }

        // yyyy-M-d / yyyy/M/d [H:m[:s]] → Excel 직렬값.
        bool tryParseDate(const std::wstring& s, double& serial)
        {
            int y = 0, mo = 0, d = 0, h = 0, mi = 0, se = 0;
            const wchar_t* p = s.c_str();
            int n = swscanf_s(p, L"%d-%d-%d %d:%d:%d", &y, &mo, &d, &h, &mi, &se);
            if (n < 3)
            {
                h = mi = se = 0;
                n = swscanf_s(p, L"%d/%d/%d %d:%d:%d", &y, &mo, &d, &h, &mi, &se);
            }
            if (n < 3 || y < 1900 || y > 9999 || mo < 1 || mo > 12 || d < 1 || d > 31)
                return false;
            SYSTEMTIME st{};
            st.wYear = (WORD)y; st.wMonth = (WORD)mo; st.wDay = (WORD)d;
            st.wHour = (WORD)(n >= 4 ? h : 0);
            st.wMinute = (WORD)(n >= 5 ? mi : 0);
            st.wSecond = (WORD)(n >= 6 ? se : 0);
            double v = 0;
            if (!SystemTimeToVariantTime(&st, &v)) return false;
            serial = v;
            return true;
        }

        // 셀 텍스트 → ExcelObj (숫자/날짜 자동 변환, 그 외 텍스트).
        ExcelObj cellValue(const std::wstring& raw, bool decimalComma = false)
        {
            const std::wstring s = trimCopy(raw);
            double v = 0;
            if (tryParseNumber(s, v, decimalComma)) return ExcelObj(v);
            if (tryParseDate(s, v)) return ExcelObj(v);
            return ExcelObj(std::wstring_view(s));
        }

        // 셀 그리드 → 배열 출력 (빈 자리는 빈 문자열).
        ExcelObj* gridResult(const std::vector<std::vector<std::wstring>>& grid,
                             bool decimalComma = false)
        {
            if (grid.empty()) return returnValue(CellError::NA);
            size_t cols = 0;
            for (const auto& r : grid) cols = (std::max)(cols, r.size());
            if (cols == 0) return returnValue(CellError::NA);

            std::vector<ExcelObj> vals;
            vals.reserve(grid.size() * cols);
            for (const auto& r : grid)
                for (size_t c = 0; c < cols; ++c)
                    vals.push_back(c < r.size() ? cellValue(r[c], decimalComma)
                                                : ExcelObj(std::wstring_view(L"")));
            return egtools::core::output(egtools::core::makeArray(
                (xloil::ExcelArrayBuilder::row_t)grid.size(),
                (xloil::ExcelArrayBuilder::col_t)cols, vals));
        }

        // ---- XML 공용 (IMPORTFEED) ----------------------------------------

        // 원시 바이트를 MSXML6로 로드(XML 선언의 인코딩을 존중). 실패 시 nullptr.
        // nsDecl 예: L"xmlns:a='http://www.w3.org/2005/Atom'"
        IXMLDOMDocument2* loadXmlBytes(const std::string& bytes, const wchar_t* nsDecl)
        {
            IXMLDOMDocument2* doc = nullptr;
            if (FAILED(CoCreateInstance(__uuidof(DOMDocument60), nullptr,
                    CLSCTX_INPROC_SERVER, __uuidof(IXMLDOMDocument2), (void**)&doc)) || !doc)
                return nullptr;
            doc->put_async(VARIANT_FALSE);
            doc->put_validateOnParse(VARIANT_FALSE);
            if (nsDecl && *nsDecl)
            {
                BSTR prop = SysAllocString(L"SelectionNamespaces");
                VARIANT v; VariantInit(&v);
                v.vt = VT_BSTR; v.bstrVal = SysAllocString(nsDecl);
                doc->setProperty(prop, v);
                VariantClear(&v);
                SysFreeString(prop);
            }

            SAFEARRAY* sa = SafeArrayCreateVector(VT_UI1, 0, (ULONG)bytes.size());
            if (!sa) { doc->Release(); return nullptr; }
            void* pv = nullptr;
            SafeArrayAccessData(sa, &pv);
            memcpy(pv, bytes.data(), bytes.size());
            SafeArrayUnaccessData(sa);

            VARIANT vsa; VariantInit(&vsa);
            vsa.vt = VT_ARRAY | VT_UI1; vsa.parray = sa;
            VARIANT_BOOL ok = VARIANT_FALSE;
            doc->load(vsa, &ok);
            VariantClear(&vsa);   // frees sa
            if (ok != VARIANT_TRUE) { doc->Release(); return nullptr; }
            return doc;
        }

        // ctx 기준 XPath 첫 노드의 텍스트. 없으면 빈 문자열 & false.
        bool nodeText(IXMLDOMNode* ctx, const wchar_t* xpath, std::wstring& out)
        {
            out.clear();
            IXMLDOMNode* node = nullptr;
            BSTR bxp = SysAllocString(xpath);
            const HRESULT hr = ctx->selectSingleNode(bxp, &node);
            SysFreeString(bxp);
            if (FAILED(hr) || !node) return false;
            BSTR t = nullptr;
            node->get_text(&t);
            if (t) { out = t; SysFreeString(t); }
            node->Release();
            return true;
        }

        // 피드 날짜(RFC822 / ISO8601) → Excel 직렬값. 실패 시 false.
        // 시간대 오프셋은 무시하고 표기된 시각 그대로 사용(VB 동작과 동일 수준).
        bool tryParseFeedDate(const std::wstring& s, double& serial)
        {
            if (tryParseDate(s, serial)) return true;   // yyyy-M-d [H:m:s]

            std::wstring t = trimCopy(s);
            // ISO8601: 2025-03-21T16:17:59(+09:00|Z)
            int y, mo, d, h = 0, mi = 0, se = 0;
            if (swscanf_s(t.c_str(), L"%d-%d-%dT%d:%d:%d", &y, &mo, &d, &h, &mi, &se) >= 5)
            {
                SYSTEMTIME st{};
                st.wYear = (WORD)y; st.wMonth = (WORD)mo; st.wDay = (WORD)d;
                st.wHour = (WORD)h; st.wMinute = (WORD)mi; st.wSecond = (WORD)se;
                double v = 0;
                if (SystemTimeToVariantTime(&st, &v)) { serial = v; return true; }
                return false;
            }
            // RFC822: [Mon, ]21 Mar 2025 16:17:59 (+0900|GMT)
            const size_t comma = t.find(L',');
            if (comma != std::wstring::npos) t = trimCopy(t.substr(comma + 1));
            wchar_t mon[8]{};
            if (swscanf_s(t.c_str(), L"%d %7s %d %d:%d:%d", &d, mon, (unsigned)8, &y,
                          &h, &mi, &se) >= 3)
            {
                static const wchar_t* kMon[12] = { L"jan", L"feb", L"mar", L"apr",
                    L"may", L"jun", L"jul", L"aug", L"sep", L"oct", L"nov", L"dec" };
                const std::wstring m3 = lowerCopy(std::wstring(mon).substr(0, 3));
                for (int i = 0; i < 12; ++i)
                    if (m3 == kMon[i])
                    {
                        SYSTEMTIME st{};
                        st.wYear = (WORD)y; st.wMonth = (WORD)(i + 1); st.wDay = (WORD)d;
                        st.wHour = (WORD)h; st.wMinute = (WORD)mi; st.wSecond = (WORD)se;
                        double v = 0;
                        if (SystemTimeToVariantTime(&st, &v)) { serial = v; return true; }
                        return false;
                    }
            }
            return false;
        }

        ExcelObj feedDateObj(const std::wstring& s)
        {
            double v = 0;
            if (tryParseFeedDate(s, v)) return ExcelObj(v);
            return ExcelObj(std::wstring_view(s));
        }

        // ---- IMPORTFEED ---------------------------------------------------

        ExcelObj* importFeed(const ExcelObj& urlObj, const ExcelObj& queryObj,
                             const ExcelObj& headersObj, const ExcelObj& numObj)
        {
            if (urlObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            if (urlObj.isMissing()) return returnValue(CellError::Value);
            const std::wstring url = trimCopy(urlObj.toString());
            if (url.empty()) return returnValue(CellError::Value);

            std::wstring query = queryObj.isMissing() ? L"items"
                                                      : lowerCopy(trimCopy(queryObj.toString()));
            if (query.empty()) query = L"items";
            const bool headers = headersObj.isMissing() ? false
                                                        : (headersObj.get<double>(0.0) != 0.0);
            int numItems = numObj.isMissing() ? 0 : numObj.get<int>(0);

            // 단일 필드 질의: "title" 또는 "items title" 형태의 마지막 토큰.
            std::wstring field;
            const bool isMulti = (query == L"items" || query == L"feed");
            if (!isMulti)
            {
                const size_t sp = query.find_last_of(L' ');
                field = (sp == std::wstring::npos) ? query : query.substr(sp + 1);
                static const wchar_t* kFields[] = { L"title", L"description", L"summary",
                                                    L"author", L"url", L"created" };
                bool okField = false;
                for (auto* f : kFields) if (field == f) { okField = true; break; }
                if (!okField) return returnValue(CellError::Value);
            }

            std::string body;
            if (!httpGet(url, body) || body.empty()) return returnValue(CellError::NA);

            // RSS 여부는 원문에서 <rss 존재로 판별 (VB 동일).
            const std::wstring probe = utf8ToWide(body.substr(0, (std::min)(body.size(), (size_t)2048)));
            const bool isRss = findNoCase(probe, L"<rss") != std::wstring::npos;

            IXMLDOMDocument2* doc = loadXmlBytes(body,
                isRss ? nullptr : L"xmlns:a='http://www.w3.org/2005/Atom'");
            if (!doc) return returnValue(CellError::NA);

            ExcelObj* result = returnValue(CellError::NA);

            if (query == L"feed")
            {
                // 피드 자체 정보 1행.
                std::wstring rootPath = isRss ? L"//channel" : L"//a:feed";
                IXMLDOMNode* root = nullptr;
                BSTR bxp = SysAllocString(rootPath.c_str());
                doc->selectSingleNode(bxp, &root);
                SysFreeString(bxp);
                if (root)
                {
                    const std::vector<const wchar_t*> paths = isRss
                        ? std::vector<const wchar_t*>{ L"title", L"description", L"link" }
                        : std::vector<const wchar_t*>{ L"a:title", L"a:subtitle",
                                                       L"a:link/@href", L"a:author/a:name" };
                    const std::vector<const wchar_t*> heads = isRss
                        ? std::vector<const wchar_t*>{ L"Title", L"Description", L"URL" }
                        : std::vector<const wchar_t*>{ L"Title", L"Summary", L"URL", L"Author" };

                    std::vector<ExcelObj> vals;
                    if (headers)
                        for (auto* h : heads) vals.push_back(ExcelObj(std::wstring_view(h)));
                    for (auto* p : paths)
                    {
                        std::wstring v;
                        nodeText(root, p, v);
                        vals.push_back(ExcelObj(std::wstring_view(v)));
                    }
                    root->Release();
                    result = egtools::core::output(egtools::core::makeArray(
                        headers ? 2 : 1, (xloil::ExcelArrayBuilder::col_t)paths.size(), vals));
                }
            }
            else
            {
                IXMLDOMNodeList* list = nullptr;
                BSTR bxp = SysAllocString(isRss ? L"//item" : L"//a:entry");
                doc->selectNodes(bxp, &list);
                SysFreeString(bxp);
                long total = 0;
                if (list) list->get_length(&total);
                if (list && total > 0)
                {
                    if (numItems <= 0 || numItems > total) numItems = (int)total;

                    struct FieldDef { const wchar_t* name; const wchar_t* rss; const wchar_t* atom; };
                    static const FieldDef kMap[] = {
                        { L"title",       L"title",       L"a:title" },
                        { L"description", L"description", L"a:subtitle" },
                        { L"summary",     L"description", L"a:content" },
                        { L"author",      L"author",      L"a:author/a:name" },
                        { L"url",         L"link",        L"a:link/@href" },
                        { L"created",     L"pubDate",     L"a:published" },
                    };
                    std::vector<FieldDef> sel;
                    if (!field.empty())
                    {
                        for (const auto& fd : kMap) if (field == fd.name) sel.push_back(fd);
                    }
                    else
                        sel = { kMap[0], kMap[2], kMap[4], kMap[3], kMap[5] };   // title,summary,url,author,created

                    std::vector<ExcelObj> vals;
                    vals.reserve(((size_t)numItems + (headers ? 1 : 0)) * sel.size());
                    if (headers)
                        for (const auto& fd : sel)
                        {
                            std::wstring h(fd.name);
                            if (h == L"url") h = L"URL"; else h[0] = (wchar_t)std::towupper(h[0]);
                            vals.push_back(ExcelObj(std::wstring_view(h)));
                        }
                    for (int i = 0; i < numItems; ++i)
                    {
                        IXMLDOMNode* item = nullptr;
                        if (list->get_item(i, &item) != S_OK || !item) break;
                        for (const auto& fd : sel)
                        {
                            std::wstring v;
                            nodeText(item, isRss ? fd.rss : fd.atom, v);
                            if (wcscmp(fd.name, L"created") == 0)
                                vals.push_back(feedDateObj(v));
                            else
                                vals.push_back(ExcelObj(std::wstring_view(decodeEntities(v))));
                        }
                        item->Release();
                    }
                    result = egtools::core::output(egtools::core::makeArray(
                        (xloil::ExcelArrayBuilder::row_t)(numItems + (headers ? 1 : 0)),
                        (xloil::ExcelArrayBuilder::col_t)sel.size(), vals));
                }
                if (list) list->Release();
            }
            doc->Release();
            return result;
        }

        // ---- IMPORTHTML (MSHTML) ------------------------------------------

        // HTML 바이트의 문자셋: BOM → meta charset= → UTF-8.
        std::wstring htmlToWide(const std::string& body)
        {
            // meta 검색은 앞부분 ASCII만 소문자로 훑는다.
            const size_t probeLen = (std::min)(body.size(), (size_t)4096);
            std::string probe(body, 0, probeLen);
            for (auto& c : probe) c = (char)tolower((unsigned char)c);
            std::wstring cs = L"utf-8";
            const size_t p = probe.find("charset=");
            if (p != std::string::npos)
            {
                size_t b = p + 8;
                if (b < probe.size() && (probe[b] == '"' || probe[b] == '\'')) ++b;
                size_t e = b;
                while (e < probe.size() &&
                       (isalnum((unsigned char)probe[e]) || probe[e] == '-' || probe[e] == '_'))
                    ++e;
                if (e > b) cs = std::wstring(probe.begin() + b, probe.begin() + e);
            }
            return decodeBytes(body, cs);
        }

        // <script>…</script> 제거. 외부 스크립트가 있으면 MSHTML 파서가 로드 대기
        // 상태로 멈춰(메시지 펌프 필요) 스크립트 뒤의 요소가 DOM에 생기지 않는다.
        // 정적 콘텐츠만 필요하므로 통째로 들어낸다.
        std::wstring stripScripts(const std::wstring& html)
        {
            const std::wstring low = lowerCopy(html);   // 검색용 1회 변환
            std::wstring out;
            out.reserve(html.size());
            size_t pos = 0;
            while (pos < html.size())
            {
                const size_t s = low.find(L"<script", pos);
                if (s == std::wstring::npos) { out.append(html, pos, html.size() - pos); break; }
                out.append(html, pos, s - pos);
                const size_t e = low.find(L"</script>", s);
                if (e == std::wstring::npos) break;
                pos = e + 9;
            }
            return out;
        }

        IHTMLDocument2* loadHtmlDoc(const std::wstring& html)
        {
            IHTMLDocument2* doc = nullptr;
            if (FAILED(CoCreateInstance(__uuidof(HTMLDocument), nullptr, CLSCTX_INPROC_SERVER,
                    __uuidof(IHTMLDocument2), (void**)&doc)) || !doc)
                return nullptr;
            SAFEARRAY* sa = SafeArrayCreateVector(VT_VARIANT, 0, 1);
            if (!sa) { doc->Release(); return nullptr; }
            VARIANT* pv = nullptr;
            SafeArrayAccessData(sa, (void**)&pv);
            pv[0].vt = VT_BSTR;
            pv[0].bstrVal = SysAllocStringLen(html.data(), (UINT)html.size());
            SafeArrayUnaccessData(sa);
            doc->write(sa);
            doc->close();
            SafeArrayDestroy(sa);   // frees the BSTR
            return doc;
        }

        std::wstring elemText(IHTMLElement* el)
        {
            BSTR t = nullptr;
            std::wstring out;
            if (el && SUCCEEDED(el->get_innerText(&t)) && t) { out = t; SysFreeString(t); }
            return trimCopy(out);
        }

        // 컬렉션 i번째 요소를 iid로 QI (실패 시 nullptr).
        void* collItem(IHTMLElementCollection* col, long i, REFIID iid)
        {
            VARIANT vi; VariantInit(&vi); vi.vt = VT_I4; vi.lVal = i;
            VARIANT vname; VariantInit(&vname); vname.vt = VT_I4; vname.lVal = i;
            IDispatch* disp = nullptr;
            if (FAILED(col->item(vname, vi, &disp)) || !disp) return nullptr;
            void* out = nullptr;
            disp->QueryInterface(iid, &out);
            disp->Release();
            return out;
        }

        // HTML 셀 값: 숫자면 변환, 그 외 텍스트 (VB 동작).
        ExcelObj htmlCellObj(const std::wstring& s, bool decimalComma)
        {
            double v = 0;
            if (tryParseNumber(s, v, decimalComma)) return ExcelObj(v);
            return ExcelObj(std::wstring_view(s));
        }

        ExcelObj* importHtml(const ExcelObj& urlObj, const ExcelObj& searchObj,
                             const ExcelObj& indexObj, const ExcelObj& localeObj)
        {
            if (urlObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            const bool decimalComma = localeObj.isMissing()
                ? false : localeDecimalComma(localeObj.toString());
            if (urlObj.isMissing()) return returnValue(CellError::Value);
            const std::wstring url = trimCopy(urlObj.toString());
            if (url.empty()) return returnValue(CellError::Value);

            std::wstring search = searchObj.isMissing() ? L"table"
                                                        : lowerCopy(trimCopy(searchObj.toString()));
            if (search.empty()) search = L"table";
            if (search != L"table" && search != L"list") return returnValue(CellError::Value);
            int index = indexObj.isMissing() ? 1 : indexObj.get<int>(1);
            if (index == 0) index = 1;
            if (index < 1) return returnValue(CellError::Value);

            std::string body;
            if (!httpGet(url, body) || body.empty()) return returnValue(CellError::NA);

            IHTMLDocument2* doc = loadHtmlDoc(stripScripts(htmlToWide(body)));
            if (!doc) return returnValue(CellError::Value);

            ExcelObj* result = returnValue(CellError::NA);

            if (search == L"table")
            {
                IHTMLDocument3* doc3 = nullptr;
                doc->QueryInterface(__uuidof(IHTMLDocument3), (void**)&doc3);
                IHTMLElementCollection* tables = nullptr;
                if (doc3)
                {
                    BSTR tag = SysAllocString(L"table");
                    doc3->getElementsByTagName(tag, &tables);
                    SysFreeString(tag);
                }
                long nTables = 0;
                if (tables) tables->get_length(&nTables);
                if (tables && index <= nTables)
                {
                    IHTMLTable* table = (IHTMLTable*)collItem(tables, index - 1,
                                                              __uuidof(IHTMLTable));
                    IHTMLElementCollection* rows = nullptr;
                    if (table) table->get_rows(&rows);
                    long nRows = 0;
                    if (rows) rows->get_length(&nRows);
                    if (rows && nRows > 0)
                    {
                        // rowspan/colspan 정규화: 점유 그리드 방식.
                        std::vector<std::vector<std::wstring>> grid((size_t)nRows);
                        std::vector<std::vector<char>> occupied((size_t)nRows);
                        auto ensure = [&](size_t r, size_t c)
                        {
                            if (occupied[r].size() <= c)
                            {
                                occupied[r].resize(c + 1, 0);
                                grid[r].resize(c + 1);
                            }
                        };
                        for (long r = 0; r < nRows; ++r)
                        {
                            IHTMLTableRow* row = (IHTMLTableRow*)collItem(rows, r,
                                                     __uuidof(IHTMLTableRow));
                            if (!row) continue;
                            IHTMLElementCollection* cells = nullptr;
                            row->get_cells(&cells);
                            long nCells = 0;
                            if (cells) cells->get_length(&nCells);
                            size_t c = 0;
                            for (long k = 0; cells && k < nCells; ++k)
                            {
                                IHTMLTableCell* cell = (IHTMLTableCell*)collItem(cells, k,
                                                           __uuidof(IHTMLTableCell));
                                if (!cell) continue;
                                ensure((size_t)r, c);
                                while (occupied[(size_t)r][c]) { ++c; ensure((size_t)r, c); }

                                long rs = 1, cs = 1;
                                cell->get_rowSpan(&rs);
                                cell->get_colSpan(&cs);
                                if (rs < 1) rs = 1;
                                if (cs < 1) cs = 1;

                                IHTMLElement* el = nullptr;
                                cell->QueryInterface(__uuidof(IHTMLElement), (void**)&el);
                                const std::wstring text = elemText(el);
                                if (el) el->Release();

                                for (long rr = 0; rr < rs && r + rr < nRows; ++rr)
                                    for (long cc = 0; cc < cs; ++cc)
                                    {
                                        ensure((size_t)(r + rr), c + (size_t)cc);
                                        occupied[(size_t)(r + rr)][c + (size_t)cc] = 1;
                                    }
                                grid[(size_t)r][c] = text;   // 값은 좌상단 셀에만
                                c += (size_t)cs;
                                cell->Release();
                            }
                            if (cells) cells->Release();
                            row->Release();
                        }
                        size_t maxCols = 0;
                        for (const auto& g : grid) maxCols = (std::max)(maxCols, g.size());
                        if (maxCols > 0)
                        {
                            std::vector<ExcelObj> vals;
                            vals.reserve(grid.size() * maxCols);
                            for (const auto& g : grid)
                                for (size_t cc = 0; cc < maxCols; ++cc)
                                    vals.push_back(cc < g.size() ? htmlCellObj(g[cc], decimalComma)
                                                                 : ExcelObj(std::wstring_view(L"")));
                            result = egtools::core::output(egtools::core::makeArray(
                                (xloil::ExcelArrayBuilder::row_t)grid.size(),
                                (xloil::ExcelArrayBuilder::col_t)maxCols, vals));
                        }
                    }
                    if (rows) rows->Release();
                    if (table) table->Release();
                }
                if (tables) tables->Release();
                if (doc3) doc3->Release();
            }
            else   // list: ul/ol을 문서 순서로 세어 index번째의 li들
            {
                IHTMLElementCollection* all = nullptr;
                doc->get_all(&all);
                long nAll = 0;
                if (all) all->get_length(&nAll);
                IHTMLElement* target = nullptr;
                int seen = 0;
                for (long i = 0; all && i < nAll && !target; ++i)
                {
                    IHTMLElement* el = (IHTMLElement*)collItem(all, i, __uuidof(IHTMLElement));
                    if (!el) continue;
                    BSTR tag = nullptr;
                    el->get_tagName(&tag);
                    if (tag)
                    {
                        if (_wcsicmp(tag, L"UL") == 0 || _wcsicmp(tag, L"OL") == 0)
                            if (++seen == index) { target = el; el = nullptr; }
                        SysFreeString(tag);
                    }
                    if (el) el->Release();
                }
                if (all) all->Release();

                if (target)
                {
                    IHTMLElement2* el2 = nullptr;
                    target->QueryInterface(__uuidof(IHTMLElement2), (void**)&el2);
                    IHTMLElementCollection* items = nullptr;
                    if (el2)
                    {
                        BSTR tag = SysAllocString(L"li");
                        el2->getElementsByTagName(tag, &items);
                        SysFreeString(tag);
                    }
                    long nItems = 0;
                    if (items) items->get_length(&nItems);
                    if (items && nItems > 0)
                    {
                        std::vector<ExcelObj> vals;
                        vals.reserve((size_t)nItems);
                        for (long i = 0; i < nItems; ++i)
                        {
                            IHTMLElement* li = (IHTMLElement*)collItem(items, i,
                                                   __uuidof(IHTMLElement));
                            vals.push_back(htmlCellObj(elemText(li), decimalComma));
                            if (li) li->Release();
                        }
                        result = egtools::core::output(egtools::core::makeArray(
                            (xloil::ExcelArrayBuilder::row_t)nItems, 1, vals));
                    }
                    if (items) items->Release();
                    if (el2) el2->Release();
                    target->Release();
                }
            }
            doc->Release();
            return result;
        }

        // ---- 텍스트 파일 가져오기 공통 코어 --------------------------------
        // IMPORTTEXT(365) / IMPORTCSV(365) / IMPORTDATA(Google)가 공유한다.

        struct ImportOpts
        {
            std::vector<std::wstring> delims;   // 텍스트 구분자(들). 비면 자동감지(탭→콤마)
            std::vector<int> fixedStarts;       // 고정폭 열 시작 위치(1기준, 오름차순)
            int skip = 0;                       // 건너뛸 행 수(음수 = 끝에서)
            bool hasTake = false;
            int take = 0;                       // 가져올 행 수(음수 = 끝에서)
            std::wstring charset = L"utf-8";
            bool decimalComma = false;          // locale: 소수점 콤마 문화권
        };

        // 인용부호("") 인식 + 임의 구분자(들, 긴 것 우선) 한 줄 분리.
        std::vector<std::wstring> parseDelimLine(const std::wstring& line,
                                                 const std::vector<std::wstring>& delims)
        {
            std::vector<std::wstring> cells;
            std::wstring cur;
            bool inQ = false;
            for (size_t i = 0; i < line.size(); )
            {
                const wchar_t c = line[i];
                if (c == L'"')
                {
                    if (inQ && i + 1 < line.size() && line[i + 1] == L'"') { cur.push_back(L'"'); i += 2; }
                    else { inQ = !inQ; ++i; }
                    continue;
                }
                if (!inQ)
                {
                    bool matched = false;
                    for (const auto& d : delims)
                        if (!d.empty() && line.compare(i, d.size(), d) == 0)
                        {
                            cells.push_back(std::move(cur));
                            cur.clear();
                            i += d.size();
                            matched = true;
                            break;
                        }
                    if (matched) continue;
                }
                cur.push_back(c);
                ++i;
            }
            cells.push_back(std::move(cur));
            return cells;
        }

        // 고정폭 분리: starts = 1기준 시작 위치 오름차순 (예 {1,3} → 2개 열).
        std::vector<std::wstring> parseFixedLine(const std::wstring& line,
                                                 const std::vector<int>& starts)
        {
            std::vector<std::wstring> cells;
            for (size_t k = 0; k < starts.size(); ++k)
            {
                const size_t b = (size_t)(starts[k] - 1);
                const size_t e = (k + 1 < starts.size()) ? (size_t)(starts[k + 1] - 1)
                                                         : line.size();
                cells.push_back(b >= line.size() ? L"" : line.substr(b, e - b));
            }
            return cells;
        }

        ExcelObj* importTextCore(const std::wstring& source, ImportOpts opt)
        {
            const std::wstring src = trimCopy(source);
            if (src.empty()) return returnValue(CellError::Value);

            std::string body;
            if (lowerCopy(src.substr(0, 4)) == L"http")
            {
                if (!httpGet(src, body)) return returnValue(CellError::Value);
            }
            else if (!readFileBytes(src, body))
                return returnValue(CellError::Value);

            // 내부 빈 행은 유지, 끝쪽 빈 행만 제거(IMPORTTEXT 행 단위 skip/take 정합).
            std::vector<std::wstring> lines = splitLines(decodeBytes(body, opt.charset),
                                                         /*keepEmpty*/ true);
            while (!lines.empty() && lines.back().empty()) lines.pop_back();
            if (lines.empty()) return returnValue(CellError::NA);

            // skip_rows: 양수=앞에서, 음수=뒤에서 제거.
            if (opt.skip > 0)
                lines.erase(lines.begin(),
                            lines.begin() + (std::min)((size_t)opt.skip, lines.size()));
            else if (opt.skip < 0)
                lines.resize(lines.size() - (std::min)((size_t)(-opt.skip), lines.size()));
            // take_rows: 양수=앞에서, 음수=뒤에서.
            if (opt.hasTake)
            {
                const size_t n = (std::min)((size_t)abs(opt.take), lines.size());
                if (opt.take >= 0) lines.resize(n);
                else lines.erase(lines.begin(), lines.end() - n);
            }
            if (lines.empty()) return returnValue(CellError::NA);

            if (opt.fixedStarts.empty() && opt.delims.empty())
            {
                // 자동감지(IMPORTDATA): 첫 줄에 탭 있으면 탭, 아니면 콤마.
                opt.delims.push_back(
                    lines[0].find(L'\t') != std::wstring::npos ? L"\t" : L",");
            }
            // 긴 구분자 우선 매칭.
            std::sort(opt.delims.begin(), opt.delims.end(),
                      [](const std::wstring& a, const std::wstring& b)
                      { return a.size() > b.size(); });

            std::vector<std::vector<std::wstring>> grid;
            grid.reserve(lines.size());
            for (const auto& ln : lines)
                grid.push_back(opt.fixedStarts.empty()
                                   ? parseDelimLine(ln, opt.delims)
                                   : parseFixedLine(ln, opt.fixedStarts));
            return gridResult(grid, opt.decimalComma);
        }

        // 구분자 인수 해석: 숫자/숫자배열 → 고정폭 시작 위치, 문자열 → 구분자.
        bool readDelimiterArg(const ExcelObj& v, ImportOpts& opt)
        {
            if (v.isMissing()) return true;
            if (v.isType(ExcelType::Multi))
            {
                ExcelArray a(v);
                const size_t n = (size_t)a.nRows() * a.nCols();
                for (size_t k = 0; k < n; ++k)
                {
                    const ExcelObj& e = a.at(k);
                    if (e.type() == ExcelType::Num || e.type() == ExcelType::Int)
                        opt.fixedStarts.push_back(e.get<int>(0));
                    else if (e.type() == ExcelType::Str && !e.toString().empty())
                        opt.delims.push_back(e.toString());
                }
                if (!opt.fixedStarts.empty())
                {
                    for (size_t i = 1; i < opt.fixedStarts.size(); ++i)
                        if (opt.fixedStarts[i] <= opt.fixedStarts[i - 1]) return false;
                    if (opt.fixedStarts[0] < 1) return false;
                    opt.delims.clear();   // 고정폭이 우선
                }
                return true;
            }
            if (v.type() == ExcelType::Num || v.type() == ExcelType::Int)
            {
                const int s = v.get<int>(0);
                if (s < 1) return false;
                opt.fixedStarts.push_back(s);
                return true;
            }
            const std::wstring d = v.toString();
            if (!d.empty()) opt.delims.push_back(d);
            return true;
        }

        int readIntArg(const ExcelObj& v, int def) { return v.isMissing() ? def : v.get<int>(def); }

        // ---- IMPORTTEXT(path,[delimiter],[skip_rows],[take_rows],[encoding],[locale]) ----
        ExcelObj* importText(const ExcelObj& pathObj, const ExcelObj& delimObj,
                             const ExcelObj& skipObj, const ExcelObj& takeObj,
                             const ExcelObj& encObj, const ExcelObj& locObj)
        {
            // 배열 인수 거부(plan/22 그룹 C) — path/take 배열의 무성 오답 차단.
            if (pathObj.isType(ExcelType::Multi) || takeObj.isType(ExcelType::Multi))
                return returnValue(CellError::Value);
            if (pathObj.isMissing()) return returnValue(CellError::Value);
            ImportOpts opt;
            if (!readDelimiterArg(delimObj, opt)) return returnValue(CellError::Value);
            if (opt.fixedStarts.empty() && opt.delims.empty())
                opt.delims.push_back(L"\t");   // 네이티브 기본: 탭
            opt.skip = readIntArg(skipObj, 0);
            if (!takeObj.isMissing()) { opt.hasTake = true; opt.take = takeObj.get<int>(0); }
            if (!encObj.isMissing()) opt.charset = encObj.toString();
            if (!locObj.isMissing()) opt.decimalComma = localeDecimalComma(locObj.toString());
            return importTextCore(pathObj.toString(), opt);
        }

        // ---- IMPORTCSV(path,[skip_rows],[take_rows],[locale]) — 콤마+UTF-8 고정 ----
        ExcelObj* importCsv(const ExcelObj& pathObj, const ExcelObj& skipObj,
                            const ExcelObj& takeObj, const ExcelObj& locObj)
        {
            if (pathObj.isType(ExcelType::Multi) || takeObj.isType(ExcelType::Multi))
                return returnValue(CellError::Value);
            if (pathObj.isMissing()) return returnValue(CellError::Value);
            ImportOpts opt;
            opt.delims.push_back(L",");
            opt.skip = readIntArg(skipObj, 0);
            if (!takeObj.isMissing()) { opt.hasTake = true; opt.take = takeObj.get<int>(0); }
            if (!locObj.isMissing()) opt.decimalComma = localeDecimalComma(locObj.toString());
            return importTextCore(pathObj.toString(), opt);
        }

        // ---- IMPORTDATA(url_or_path,[delimiter],[locale],[charset]) — Google 정합 ----
        // 2번째 인수가 알려진 문자셋 이름이면 구버전 호출(IMPORTDATA(url,"euc-kr"))로
        // 간주해 charset으로 받아준다(과도기 하위 호환).
        ExcelObj* importData(const ExcelObj& urlObj, const ExcelObj& delimObj,
                             const ExcelObj& locObj, const ExcelObj& charsetObj)
        {
            if (urlObj.isType(ExcelType::Multi)) return returnValue(CellError::Value);
            if (urlObj.isMissing()) return returnValue(CellError::Value);
            ImportOpts opt;

            bool legacyCharset = false;
            if (!delimObj.isMissing() && delimObj.type() == ExcelType::Str)
            {
                bool le, be;
                const std::wstring s = delimObj.toString();
                // 이름이 코드페이지로 해석되고 통상적 구분자 길이(1~2)가 아니면 문자셋.
                if (s.size() > 2)
                {
                    codepageFromName(s, le, be);
                    const std::wstring l = lowerCopy(trimCopy(s));
                    legacyCharset = le || be ||
                        l == L"utf-8" || l == L"utf8" || l == L"us-ascii" || l == L"ascii" ||
                        l == L"iso-8859-1" || l == L"latin1" || l == L"euc-kr" || l == L"cp949" ||
                        l == L"ks_c_5601-1987" || l == L"shift_jis" || l == L"shift-jis" ||
                        l == L"sjis" || l == L"gb2312" || l == L"gbk" || l == L"big5" ||
                        l == L"windows-1252" ||
                        (!l.empty() && std::all_of(l.begin(), l.end(), ::iswdigit));
                }
                if (legacyCharset) opt.charset = s;
            }
            if (!legacyCharset && !readDelimiterArg(delimObj, opt))
                return returnValue(CellError::Value);
            if (!locObj.isMissing()) opt.decimalComma = localeDecimalComma(locObj.toString());
            if (!charsetObj.isMissing()) opt.charset = charsetObj.toString();
            return importTextCore(urlObj.toString(), opt);
        }

        // ---- IMPORTRANGE --------------------------------------------------

        // A1 주소 한쪽("B12"/"B"/"12") → 1기준 행·열. 열/행이 없으면 기본값.
        void parseA1Side(const std::wstring& s, bool endSide, long& row, long& col)
        {
            std::wstring letters, digits;
            for (wchar_t c : s)
            {
                if (iswalpha(c)) letters.push_back((wchar_t)std::towupper(c));
                else if (iswdigit(c)) digits.push_back(c);
            }
            col = 0;
            for (wchar_t c : letters) col = col * 26 + (c - L'A' + 1);
            if (col == 0) col = endSide ? LONG_MAX : 1;
            row = digits.empty() ? (endSide ? LONG_MAX : 1)
                                 : wcstol(digits.c_str(), nullptr, 10);
            if (row <= 0) row = 1;
        }

        // "Sheet1!$A$1:B5" 등 → [r1,c1,r2,c2] (1기준). 빈 주소면 false(전체).
        bool parseRangeAddress(std::wstring addr, long& r1, long& c1, long& r2, long& c2)
        {
            addr = trimCopy(addr);
            addr.erase(std::remove(addr.begin(), addr.end(), L'$'), addr.end());
            const size_t bang = addr.find(L'!');
            if (bang != std::wstring::npos) addr = addr.substr(bang + 1);
            if (addr.empty()) return false;

            const size_t colon = addr.find(L':');
            const std::wstring a = (colon == std::wstring::npos) ? addr : addr.substr(0, colon);
            const std::wstring b = (colon == std::wstring::npos) ? addr : addr.substr(colon + 1);
            parseA1Side(a, false, r1, c1);
            parseA1Side(b, true, r2, c2);
            if (r2 < r1) std::swap(r1, r2);
            if (c2 < c1) std::swap(c1, c2);
            return true;
        }

        ExcelObj* importRange(const ExcelObj& urlObj, const ExcelObj& rangeObj)
        {
            // 배열 인수 거부(plan/22 그룹 C). 특히 range_address 자리에 실수로
            // 실제 범위 참조를 넘기면 "[R x C]"가 X24 따위로 파싱돼 오답이 났다.
            if (urlObj.isType(ExcelType::Multi) || rangeObj.isType(ExcelType::Multi))
                return returnValue(CellError::Value);
            if (urlObj.isMissing()) return returnValue(CellError::Value);
            const std::wstring url = trimCopy(urlObj.toString());
            if (url.empty()) return returnValue(CellError::Value);

            // 스프레드시트 ID (…/d/<id>/…)와 gid(#gid=N / ?gid=N) 추출.
            const size_t dPos = url.find(L"/d/");
            if (dPos == std::wstring::npos) return returnValue(CellError::Value);
            const size_t idBeg = dPos + 3;
            size_t idEnd = url.find_first_of(L"/?#", idBeg);
            if (idEnd == std::wstring::npos) idEnd = url.size();
            const std::wstring id = url.substr(idBeg, idEnd - idBeg);
            if (id.empty()) return returnValue(CellError::Value);

            std::wstring gid = L"0";
            const size_t gPos = url.find(L"gid=");
            if (gPos != std::wstring::npos)
            {
                size_t e = gPos + 4;
                while (e < url.size() && iswdigit(url[e])) ++e;
                if (e > gPos + 4) gid = url.substr(gPos + 4, e - (gPos + 4));
            }

            const std::wstring exportUrl =
                L"https://docs.google.com/spreadsheets/d/" + id +
                L"/export?format=tsv&gid=" + gid;

            std::string body;
            if (!httpGet(exportUrl, body)) return returnValue(CellError::NA);
            const std::wstring text = decodeBytes(body, L"utf-8");

            // 공유 미설정 등: TSV 대신 HTML(로그인 페이지)이 온다.
            if (findNoCase(text.substr(0, (std::min)(text.size(), (size_t)300)), L"<html")
                    != std::wstring::npos)
                return returnValue(CellError::NA);

            std::vector<std::wstring> lines = splitLines(text, /*keepEmpty*/true);
            while (!lines.empty() && lines.back().empty()) lines.pop_back();
            if (lines.empty()) return returnValue(CellError::NA);

            std::vector<std::vector<std::wstring>> grid;
            grid.reserve(lines.size());
            for (const auto& ln : lines) grid.push_back(splitTabs(ln));

            // 범위 주소 지정 시 해당 부분만 절단.
            long r1, c1, r2, c2;
            if (!rangeObj.isMissing() &&
                parseRangeAddress(rangeObj.toString(), r1, c1, r2, c2))
            {
                if ((size_t)r1 > grid.size()) return returnValue(CellError::NA);
                const size_t rEnd = (std::min)((size_t)r2, grid.size());
                std::vector<std::vector<std::wstring>> cut;
                cut.reserve(rEnd - (size_t)r1 + 1);
                for (size_t r = (size_t)r1 - 1; r < rEnd; ++r)
                {
                    const auto& row = grid[r];
                    std::vector<std::wstring> outRow;
                    if ((size_t)c1 <= row.size())
                    {
                        const size_t cEnd = (std::min)((size_t)c2, row.size());
                        for (size_t c = (size_t)c1 - 1; c < cEnd; ++c) outRow.push_back(row[c]);
                    }
                    cut.push_back(std::move(outRow));
                }
                grid.swap(cut);
            }
            return gridResult(grid);
        }

        // ---- QUERY (ACE 텍스트 드라이버 + ADO late-bound) -------------------

        // IDispatch 멤버 호출. argsFwd는 선언 순서대로 주면 내부에서 역순으로
        // 뒤집는다(DISPPARAMS 규약). 실패 시 EXCEPINFO 설명을 errDesc에 채움.
        HRESULT invokeDisp(IDispatch* d, const wchar_t* name, WORD flags,
                           std::vector<VARIANT>& argsFwd, VARIANT* out,
                           std::wstring* errDesc = nullptr)
        {
            DISPID id = 0;
            OLECHAR* nm = const_cast<OLECHAR*>(name);
            HRESULT hr = d->GetIDsOfNames(IID_NULL, &nm, 1, LOCALE_USER_DEFAULT, &id);
            if (FAILED(hr)) return hr;
            std::vector<VARIANT> rev(argsFwd.rbegin(), argsFwd.rend());
            DISPPARAMS dp{ rev.empty() ? nullptr : rev.data(), nullptr, (UINT)rev.size(), 0 };
            EXCEPINFO ei{};
            hr = d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &dp, out, &ei, nullptr);
            if (FAILED(hr) && errDesc && ei.bstrDescription) *errDesc = ei.bstrDescription;
            if (ei.bstrDescription) SysFreeString(ei.bstrDescription);
            if (ei.bstrSource) SysFreeString(ei.bstrSource);
            if (ei.bstrHelpFile) SysFreeString(ei.bstrHelpFile);
            return hr;
        }

        VARIANT bstrVar(const std::wstring& s)
        {
            VARIANT v; VariantInit(&v);
            v.vt = VT_BSTR; v.bstrVal = SysAllocString(s.c_str());
            return v;
        }
        VARIANT i4Var(long n)
        {
            VARIANT v; VariantInit(&v);
            v.vt = VT_I4; v.lVal = n;
            return v;
        }

        // 셀 값 → CSV 필드 텍스트(로캘 무관, 문자열은 " 이스케이프 후 인용).
        std::wstring csvField(const ExcelObj& e)
        {
            switch (e.type())
            {
            case ExcelType::Num:
            {
                wchar_t buf[40];
                swprintf_s(buf, L"%.15g", e.get<double>(0.0));
                return buf;
            }
            case ExcelType::Int:
            {
                wchar_t buf[24];
                swprintf_s(buf, L"%d", e.get<int>(0));
                return buf;
            }
            case ExcelType::Bool:
                return e.get<bool>(false) ? L"TRUE" : L"FALSE";
            case ExcelType::Str:
            {
                std::wstring s = e.toString(), out = L"\"";
                for (wchar_t c : s)
                {
                    if (c == L'"') out += L"\"\"";
                    else out.push_back(c);
                }
                out += L"\"";
                return out;
            }
            default:
                return L"";
            }
        }

        ExcelObj* queryData(const ExcelObj& dataObj, const ExcelObj& queryObj,
                            const ExcelObj& headersObj)
        {
            auto errStr = [](const std::wstring& msg) -> ExcelObj*
            {
                return returnValue(ExcelObj(std::wstring_view(L"ERROR: " + msg)));
            };

            if (dataObj.isMissing() || queryObj.isMissing())
                return returnValue(CellError::Value);
            std::wstring query = trimCopy(queryObj.toString());
            if (query.empty()) return returnValue(CellError::Value);
            // headers — Google 의미(헤더 행 수, 숫자): N≥1 = 위 N행이 필드명,
            // 0 = 없음(F1..Fn), -1/생략 = 자동 추정. TRUE/FALSE도 하위 호환(1/0).
            int headerRows = -1;
            if (!headersObj.isMissing())
            {
                if (headersObj.type() == ExcelType::Bool)
                    headerRows = headersObj.get<bool>(false) ? 1 : 0;
                else
                    headerRows = headersObj.get<int>(-1);
            }

            // 1) 데이터 → %TEMP%\EGTOOLS.csv (UTF-8, 항상 HDR=Yes 형태로 기록:
            //    headerRows>0이면 위 N행을 합쳐 필드명, 아니면 F1..Fn을 생성).
            std::wstring csv;
            {
                size_t nR = 1, nC = 1;
                const bool isArr = dataObj.isType(ExcelType::Multi);
                if (isArr)
                {
                    ExcelArray arr(dataObj);
                    nR = arr.nRows(); nC = arr.nCols();
                }
                if (nR == 0 || nC == 0) return returnValue(CellError::Value);

                auto cellAt = [&](size_t r, size_t c) -> ExcelObj
                {
                    if (isArr) { ExcelArray arr(dataObj); return ExcelObj(arr.at(r * nC + c)); }
                    return ExcelObj(dataObj);
                };

                // 자동 추정(-1): 첫 행이 전부 텍스트이고 데이터 행이 더 있으면 1.
                if (headerRows < 0)
                {
                    bool allText = nR > 1;
                    for (size_t c = 0; allText && c < nC; ++c)
                        if (cellAt(0, c).type() != ExcelType::Str) allText = false;
                    headerRows = allText ? 1 : 0;
                }
                if ((size_t)headerRows >= nR && headerRows > 0)
                    return returnValue(CellError::Value);

                // 헤더 줄: 위 N행 결합(공백 연결) 또는 F1..Fn. 빈/중복 이름 방지.
                std::vector<std::wstring> names(nC);
                for (size_t c = 0; c < nC; ++c)
                {
                    if (headerRows > 0)
                    {
                        std::wstring nm;
                        for (int r = 0; r < headerRows; ++r)
                        {
                            const std::wstring part = trimCopy(cellAt((size_t)r, c).toString());
                            if (part.empty()) continue;
                            if (!nm.empty()) nm += L' ';
                            nm += part;
                        }
                        names[c] = nm;
                    }
                    if (names[c].empty()) names[c] = L"F" + std::to_wstring(c + 1);
                    for (size_t p = 0; p < c; ++p)
                        if (_wcsicmp(names[p].c_str(), names[c].c_str()) == 0)
                        { names[c] += L"_" + std::to_wstring(c + 1); break; }
                }
                for (size_t c = 0; c < nC; ++c)
                {
                    if (c) csv += L',';
                    std::wstring quoted = L"\"";
                    for (wchar_t ch : names[c])
                    { if (ch == L'"') quoted += L"\"\""; else quoted.push_back(ch); }
                    csv += quoted + L"\"";
                }
                csv += L"\r\n";

                for (size_t r = (size_t)headerRows; r < nR; ++r)
                {
                    for (size_t c = 0; c < nC; ++c)
                    {
                        if (c) csv += L',';
                        csv += csvField(cellAt(r, c));
                    }
                    csv += L"\r\n";
                }
            }

            wchar_t tmpDir[MAX_PATH]{};
            GetTempPathW(MAX_PATH, tmpDir);
            const std::wstring dir = tmpDir;
            const std::wstring path = dir + L"EGTOOLS.csv";
            {
                int n = WideCharToMultiByte(CP_UTF8, 0, csv.data(), (int)csv.size(),
                                            nullptr, 0, nullptr, nullptr);
                std::string u8((size_t)n, '\0');
                WideCharToMultiByte(CP_UTF8, 0, csv.data(), (int)csv.size(),
                                    u8.data(), n, nullptr, nullptr);
                std::ofstream f(path.c_str(), std::ios::binary | std::ios::trunc);
                if (!f) return errStr(L"cannot write temp CSV");
                f.write(u8.data(), (std::streamsize)u8.size());
            }

            // 2) 쿼리의 테이블명 EGTOOLS → EGTOOLS#csv (텍스트 드라이버 표기).
            {
                std::wstring up = lowerCopy(query);
                std::wstring out;
                size_t pos = 0;
                while (pos < query.size())
                {
                    const size_t p = up.find(L"egtools", pos);
                    if (p == std::wstring::npos) { out.append(query, pos, query.size() - pos); break; }
                    out.append(query, pos, p - pos);
                    out += query.substr(p, 7);
                    if (up.compare(p, 11, L"egtools#csv") != 0) out += L"#csv";
                    else { out += L"#csv"; pos = p + 11; continue; }
                    pos = p + 7;
                }
                query = out;
            }

            // 3) ADODB.Recordset 열기 — ACE 12.0 → 16.0 폴백.
            CLSID clsid;
            if (FAILED(CLSIDFromProgID(L"ADODB.Recordset", &clsid)))
                return errStr(L"ADO not available");
            IDispatch* rs = nullptr;
            if (FAILED(CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER,
                    IID_IDispatch, (void**)&rs)) || !rs)
                return errStr(L"ADO not available");

            std::wstring lastErr;
            bool opened = false;
            for (const wchar_t* ver : { L"12.0", L"16.0" })
            {
                const std::wstring conn =
                    std::wstring(L"Provider=Microsoft.ACE.OLEDB.") + ver +
                    L";Data Source=" + dir +
                    L";Extended Properties='text;HDR=Yes;FMT=Delimited(,);IMEX=1;CharacterSet=65001'";
                std::vector<VARIANT> args;
                args.push_back(bstrVar(query));   // Source
                args.push_back(bstrVar(conn));    // ActiveConnection
                args.push_back(i4Var(3));         // adOpenStatic
                args.push_back(i4Var(1));         // adLockReadOnly
                args.push_back(i4Var(1));         // adCmdText
                std::wstring err;
                const HRESULT hr = invokeDisp(rs, L"Open", DISPATCH_METHOD, args, nullptr, &err);
                for (auto& a : args) VariantClear(&a);
                if (SUCCEEDED(hr)) { opened = true; break; }
                if (!err.empty()) lastErr = err;
            }
            if (!opened)
            {
                rs->Release();
                return errStr(lastErr.empty() ? L"cannot open ACE OLEDB (provider missing?)"
                                              : lastErr);
            }

            ExcelObj* result = returnValue(CellError::NA);

            // EOF면 결과 행 없음 → #N/A.
            VARIANT vEof; VariantInit(&vEof);
            std::vector<VARIANT> noArgs;
            if (SUCCEEDED(invokeDisp(rs, L"EOF", DISPATCH_PROPERTYGET, noArgs, &vEof)) &&
                !(vEof.vt == VT_BOOL && vEof.boolVal == VARIANT_TRUE))
            {
                VARIANT vRows; VariantInit(&vRows);
                if (SUCCEEDED(invokeDisp(rs, L"GetRows", DISPATCH_METHOD, noArgs, &vRows)) &&
                    (vRows.vt & VT_ARRAY) && vRows.parray &&
                    SafeArrayGetDim(vRows.parray) == 2)
                {
                    SAFEARRAY* sa = vRows.parray;
                    LONG f0, f1, r0, r1;
                    SafeArrayGetLBound(sa, 1, &f0); SafeArrayGetUBound(sa, 1, &f1);
                    SafeArrayGetLBound(sa, 2, &r0); SafeArrayGetUBound(sa, 2, &r1);
                    const long nF = f1 - f0 + 1, nRec = r1 - r0 + 1;
                    if (nF > 0 && nRec > 0)
                    {
                        std::vector<ExcelObj> vals;
                        vals.reserve((size_t)nF * nRec);
                        for (long r = 0; r < nRec; ++r)
                            for (long f = 0; f < nF; ++f)
                            {
                                LONG idx[2] = { f0 + f, r0 + r };
                                VARIANT v; VariantInit(&v);
                                SafeArrayGetElement(sa, idx, &v);
                                if (v.vt == VT_NULL || v.vt == VT_EMPTY)
                                    vals.push_back(ExcelObj(std::wstring_view(L"")));
                                else
                                    vals.push_back(xloil::variantToExcelObj(v, false));
                                VariantClear(&v);
                            }
                        result = egtools::core::output(egtools::core::makeArray(
                            (xloil::ExcelArrayBuilder::row_t)nRec,
                            (xloil::ExcelArrayBuilder::col_t)nF, vals));
                    }
                }
                VariantClear(&vRows);
            }
            VariantClear(&vEof);

            invokeDisp(rs, L"Close", DISPATCH_METHOD, noArgs, nullptr);
            rs->Release();
            return result;
        }

        // ---- GOOGLETRANSLATE ----------------------------------------------

        bool isValidLangCode(const std::wstring& lc)
        {
            static const wchar_t* kCodes[] = {
                L"auto", L"af", L"sq", L"am", L"ar", L"hy", L"as", L"ay", L"az", L"bm",
                L"eu", L"be", L"bn", L"bho", L"bs", L"bg", L"ca", L"ceb", L"zh-cn", L"zh",
                L"zh-tw", L"co", L"hr", L"cs", L"da", L"dv", L"doi", L"nl", L"en", L"eo",
                L"et", L"ee", L"fil", L"fi", L"fr", L"fy", L"gl", L"ka", L"de", L"el",
                L"gn", L"gu", L"ht", L"ha", L"haw", L"he", L"iw", L"hi", L"hmn", L"hu",
                L"is", L"ig", L"ilo", L"id", L"ga", L"it", L"ja", L"jv", L"jw", L"kn",
                L"kk", L"km", L"rw", L"gom", L"ko", L"kr", L"kri", L"ku", L"ckb", L"ky",
                L"lo", L"la", L"lv", L"ln", L"lt", L"lg", L"lb", L"mk", L"mai", L"mg",
                L"ms", L"ml", L"mt", L"mi", L"mr", L"mni-mtei", L"lus", L"mn", L"my", L"ne",
                L"no", L"ny", L"or", L"om", L"ps", L"fa", L"pl", L"pt", L"pa", L"qu",
                L"ro", L"ru", L"sm", L"sa", L"gd", L"nso", L"sr", L"st", L"sn", L"sd",
                L"si", L"sk", L"sl", L"so", L"es", L"su", L"sw", L"sv", L"tl", L"tg",
                L"ta", L"tt", L"te", L"th", L"ti", L"ts", L"tr", L"tk", L"ak", L"uk",
                L"ur", L"ug", L"uz", L"vi", L"cy", L"xh", L"yi", L"yo", L"zu" };
            for (auto* c : kCodes)
                if (lc == c) return true;
            return false;
        }

        // 소문자 코드 → Google 표기. (VB판의 부분 문자열 치환 버그 — "zh-CN"이
        // "zh-CN-CN"이 되던 것 — 을 정확일치 매핑으로 대체.)
        std::wstring normalizeLang(const std::wstring& lc)
        {
            if (lc == L"kr") return L"ko";
            if (lc == L"zh" || lc == L"zh-cn") return L"zh-CN";
            if (lc == L"zh-tw") return L"zh-TW";
            return lc;
        }

        // 언어 인수 읽기: 생략/빈칸이면 def, 아니면 소문자 trim.
        std::wstring langArg(const ExcelObj& v, const std::wstring& def)
        {
            if (v.isMissing()) return def;
            const std::wstring s = lowerCopy(trimCopy(v.toString()));
            return s.empty() ? def : s;
        }

        ExcelObj* googleTranslate(const ExcelObj& textObj, const ExcelObj& slObj,
                                  const ExcelObj& tlObj)
        {
            if (textObj.isMissing()) return returnValue(CellError::Value);

            // 범위/배열 입력은 줄바꿈으로 이어붙여 한 번에 번역 (VB 동작).
            std::wstring text;
            if (textObj.isType(ExcelType::Multi))
            {
                ExcelArray a(textObj);
                const size_t n = (size_t)a.nRows() * a.nCols();
                for (size_t k = 0; k < n; ++k)
                {
                    if (!text.empty()) text += L'\n';
                    text += a.at(k).toString();
                }
            }
            else
                text = textObj.toString();

            text = trimCopy(text);
            if (text.empty()) return returnValue(ExcelObj(std::wstring_view(L"")));
            if (text.size() > 5000)
                return returnValue(ExcelObj(std::wstring_view(L"ERROR:Exceed max length")));

            // 대상언어 기본값은 UI 언어 (VB판은 "ko" 고정이었음).
            std::wstring sl = langArg(slObj, L"auto");
            std::wstring tl = langArg(tlObj, lowerCopy(egtools::i18n::current()));

            if (!isValidLangCode(sl))
                return returnValue(ExcelObj(std::wstring_view(L"ERROR:Source language Code invalid")));
            if (!isValidLangCode(tl))
                return returnValue(ExcelObj(std::wstring_view(L"ERROR:Target language Code invalid")));
            sl = normalizeLang(sl);
            tl = normalizeLang(tl);

            const std::wstring url = L"https://translate.google.com/m?hl=en&sl=" + sl +
                                     L"&tl=" + tl + L"&q=" + percentEncode(text);

            std::string body;
            if (!httpGet(url, body) || body.empty())
                return returnValue(CellError::Value);
            const std::wstring html = utf8ToWide(body);

            // Google 오류 페이지: <title>Error NNN …</title>
            size_t pos = findNoCase(html, L"<title>Error");
            if (pos != std::wstring::npos)
            {
                const size_t from = pos + 7;                     // "<title>" 뒤
                const size_t end = findNoCase(html, L"</title>", from);
                const std::wstring msg = (end == std::wstring::npos)
                    ? L"Error" : html.substr(from, end - from);
                return returnValue(ExcelObj(std::wstring_view(L"ERROR:" + msg)));
            }

            static const std::wstring kMarker = L"<div class=\"result-container\">";
            pos = findNoCase(html, kMarker);
            if (pos == std::wstring::npos)
                return returnValue(ExcelObj(std::wstring_view(L"ERROR:No result")));
            const size_t from = pos + kMarker.size();
            const size_t end = findNoCase(html, L"</div>", from);
            if (end == std::wstring::npos)
                return returnValue(ExcelObj(std::wstring_view(L"ERROR:html error")));

            const std::wstring result = decodeEntities(html.substr(from, end - from));
            return returnValue(ExcelObj(std::wstring_view(result)));
        }
    }

    void registerGoogle()
    {
        auto fn = [](const ExcelObj& text, const ExcelObj& sl, const ExcelObj& tl)
            -> ExcelObj* { return googleTranslate(text, sl, tl); };

        // 스칼라 네트워크 함수 — legacy 호스트에서도 MTR 병렬 허용(효과 큼).
        egtools::core::registerFn(L"GOOGLETRANSLATE", fn, /*macro*/ false, /*threadsafe*/ true);
        egtools::core::registerFn(L"GTRS", fn, /*macro*/ false, /*threadsafe*/ true);   // 짧은 별칭

        // 텍스트 가져오기 3종 — 공통 코어(importTextCore) 공유.
        egtools::core::registerFn(L"IMPORTDATA",
            [](const ExcelObj& url, const ExcelObj& delim, const ExcelObj& loc,
               const ExcelObj& cs) -> ExcelObj*
            { return importData(url, delim, loc, cs); });

        egtools::core::registerFn(L"IMPORTTEXT",
            [](const ExcelObj& path, const ExcelObj& delim, const ExcelObj& skip,
               const ExcelObj& take, const ExcelObj& enc, const ExcelObj& loc) -> ExcelObj*
            { return importText(path, delim, skip, take, enc, loc); });

        egtools::core::registerFn(L"IMPORTCSV",
            [](const ExcelObj& path, const ExcelObj& skip, const ExcelObj& take,
               const ExcelObj& loc) -> ExcelObj*
            { return importCsv(path, skip, take, loc); });

        // QUERY/IMPORTHTML/IMPORTFEED는 실행 중 COM(ADO/MSHTML/MSXML)을 만들고
        // QUERY는 공용 임시 CSV도 사용 — MTR 워커 불가.
        egtools::core::registerFn(L"QUERY",
            [](const ExcelObj& data, const ExcelObj& query, const ExcelObj& headers) -> ExcelObj*
            {
                try { return queryData(data, query, headers); }
                catch (...) { return returnValue(CellError::Value); }
            },
            /*macro*/ false, /*threadsafe*/ false);

        egtools::core::registerFn(L"IMPORTHTML",
            [](const ExcelObj& url, const ExcelObj& search, const ExcelObj& index,
               const ExcelObj& locale) -> ExcelObj*
            { return importHtml(url, search, index, locale); },
            /*macro*/ false, /*threadsafe*/ false);

        egtools::core::registerFn(L"IMPORTRANGE",
            [](const ExcelObj& url, const ExcelObj& range) -> ExcelObj*
            { return importRange(url, range); });

        egtools::core::registerFn(L"IMPORTFEED",
            [](const ExcelObj& url, const ExcelObj& query, const ExcelObj& headers,
               const ExcelObj& num) -> ExcelObj*
            { return importFeed(url, query, headers, num); },
            /*macro*/ false, /*threadsafe*/ false);
    }
}
