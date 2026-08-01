// FxPublicApi.cpp — Korean public-API functions (plan/22 B4), ported from VB
// FX7_PublicAPI with a registry-backed API-key store.
//
//   SEARCHADDRESS(text, [info], [row], [api_key])   — juso.go.kr address search (XML)
//   BRNSTATUS(numbers, [api_key])                   — NTS business-registration status (odcloud, JSON POST)
//   GEOSEARCH(keyword, [type], [epsg], [api_key])   — VWorld search (address/place → x,y[,addr])
//   GEOCODER(keyword, [epsg], [api_key])            — VWorld geocoder (road → parcel fallback)
//   GEOADDRESS(x, y, [return_type], [epsg], [api_key]) — VWorld reverse geocoder
//   ADDRESSMAP(address, [scale], [api_key])         — VWorld WMS map picture into the cell
//
// API keys (user decision 2026-08-01): NO key ships with the add-in. The last
// optional argument takes a key; on first use it is stored under
// HKCU\Software\EGTools++\ApiKeys (value: juso / odcloud / vworld) and reused
// when the argument is omitted. When a service rejects the key
// (expired/unregistered), the stored value is deleted and the error tells the
// user where to get a new key. Dev keys live in the repo-root .env
// (gitignored, test-only — production code never reads it).

#include <windows.h>
#include <winhttp.h>
#include <msxml6.h>

#include "../core/Registry.h"
#include "../core/Spill.h"
#include "../core/ArrayUtil.h"
#include "ImageInsert.h"

#include <xlOil/xlOil.h>
#include <xlOil/Caller.h>
#include <xlOil/ExcelArray.h>

#include <nlohmann/json.hpp>

#include <algorithm>
#include <cwctype>
#include <string>
#include <vector>

#pragma comment(lib, "winhttp.lib")

using namespace xloil;
using json = nlohmann::json;

namespace egtools::functions
{
    namespace
    {
        // ── UTF-8 ⇄ wide ────────────────────────────────────────────────────
        std::wstring fromUtf8Pa(const std::string& s)
        {
            if (s.empty()) return L"";
            int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
            std::wstring w((size_t)n, L'\0');
            MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), w.data(), n);
            return w;
        }

        std::string toUtf8Pa(const std::wstring& w)
        {
            if (w.empty()) return "";
            int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, nullptr);
            std::string s((size_t)n, '\0');
            WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), s.data(), n, nullptr, nullptr);
            return s;
        }

        std::wstring urlEncodeW(const std::wstring& w)
        {
            const std::string u8 = toUtf8Pa(w);
            static const wchar_t* kHex = L"0123456789ABCDEF";
            std::wstring out;
            for (unsigned char c : u8)
            {
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
                    (c >= '0' && c <= '9') || c == '-' || c == '_' || c == '.' || c == '~')
                    out.push_back((wchar_t)c);
                else { out.push_back(L'%'); out.push_back(kHex[c >> 4]); out.push_back(kHex[c & 0xF]); }
            }
            return out;
        }

        std::wstring trimWs(std::wstring s)
        {
            while (!s.empty() && iswspace(s.front())) s.erase(0, 1);
            while (!s.empty() && iswspace(s.back())) s.pop_back();
            return s;
        }

        // ── API-key registry store ──────────────────────────────────────────
        constexpr const wchar_t* kKeyPath = L"Software\\EGTools++\\ApiKeys";

        std::wstring regGetKey(const wchar_t* service)
        {
            HKEY h;
            if (RegOpenKeyExW(HKEY_CURRENT_USER, kKeyPath, 0, KEY_READ, &h) != ERROR_SUCCESS)
                return L"";
            wchar_t buf[512]{};
            DWORD cb = sizeof(buf), type = 0;
            std::wstring out;
            if (RegQueryValueExW(h, service, nullptr, &type, (BYTE*)buf, &cb) == ERROR_SUCCESS &&
                type == REG_SZ)
                out = buf;
            RegCloseKey(h);
            return trimWs(out);
        }

        void regSetKey(const wchar_t* service, const std::wstring& key)
        {
            HKEY h;
            if (RegCreateKeyExW(HKEY_CURRENT_USER, kKeyPath, 0, nullptr, 0, KEY_WRITE,
                                nullptr, &h, nullptr) != ERROR_SUCCESS)
                return;
            RegSetValueExW(h, service, 0, REG_SZ, (const BYTE*)key.c_str(),
                           (DWORD)((key.size() + 1) * sizeof(wchar_t)));
            RegCloseKey(h);
        }

        void regDelKey(const wchar_t* service)
        {
            HKEY h;
            if (RegOpenKeyExW(HKEY_CURRENT_USER, kKeyPath, 0, KEY_WRITE, &h) != ERROR_SUCCESS)
                return;
            RegDeleteValueW(h, service);
            RegCloseKey(h);
        }

        // arg key (→ store) > stored key > "" (caller returns guidance).
        std::wstring resolveKey(const ExcelObj& keyA, const wchar_t* service)
        {
            std::wstring k = trimWs(keyA.toString());
            if (!k.empty()) { regSetKey(service, k); return k; }
            return regGetKey(service);
        }

        ExcelObj* keyMissingError(const wchar_t* service, const wchar_t* issuer)
        {
            std::wstring msg = std::wstring(L"ERROR: API 키 필요 — ") + issuer +
                L" 에서 발급 후 마지막 인수로 1회 입력하면 저장됩니다(서비스: " + service + L")";
            return returnValue(ExcelObj(std::wstring_view(msg)));
        }

        ExcelObj* keyRejectedError(const wchar_t* service, const wchar_t* issuer)
        {
            regDelKey(service);
            std::wstring msg = std::wstring(L"ERROR: API 키가 거부되어(만료/미등록) 저장된 키를 삭제했습니다 — ")
                + issuer + L" 에서 재발급 후 마지막 인수로 다시 입력하세요(서비스: " + service + L")";
            return returnValue(ExcelObj(std::wstring_view(msg)));
        }

        // ── HTTP (WinHTTP) ──────────────────────────────────────────────────
        bool httpCall(const std::wstring& url, const wchar_t* method,
                      const std::string& body, const wchar_t* contentType,
                      std::string& out, DWORD* statusOut = nullptr)
        {
            URL_COMPONENTS uc{}; uc.dwStructSize = sizeof(uc);
            std::wstring host((size_t)url.size() + 1, L'\0');
            std::wstring path((size_t)url.size() + 1, L'\0');
            uc.lpszHostName = host.data(); uc.dwHostNameLength = (DWORD)url.size();
            uc.lpszUrlPath = path.data(); uc.dwUrlPathLength = (DWORD)url.size();
            if (!WinHttpCrackUrl(url.c_str(), (DWORD)url.size(), 0, &uc)) return false;
            host.resize(uc.dwHostNameLength);

            HINTERNET hs = WinHttpOpen(L"EGTools++/1.0 (WinHTTP)",
                WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);
            if (!hs) return false;
            WinHttpSetTimeouts(hs, 8000, 8000, 15000, 30000);

            bool ok = false;
            HINTERNET hc = WinHttpConnect(hs, host.c_str(), uc.nPort, 0);
            if (hc)
            {
                const DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
                HINTERNET hr = WinHttpOpenRequest(hc, method, uc.lpszUrlPath, nullptr,
                    WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
                if (hr)
                {
                    std::wstring hdr;
                    if (contentType) hdr = std::wstring(L"Content-Type: ") + contentType + L"\r\n";
                    if (WinHttpSendRequest(hr,
                            hdr.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : hdr.c_str(),
                            hdr.empty() ? 0 : (DWORD)hdr.size(),
                            body.empty() ? WINHTTP_NO_REQUEST_DATA : (LPVOID)body.data(),
                            (DWORD)body.size(), (DWORD)body.size(), 0) &&
                        WinHttpReceiveResponse(hr, nullptr))
                    {
                        if (statusOut)
                        {
                            DWORD sc = 0, cb = sizeof(sc);
                            WinHttpQueryHeaders(hr,
                                WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                WINHTTP_HEADER_NAME_BY_INDEX, &sc, &cb, WINHTTP_NO_HEADER_INDEX);
                            *statusOut = sc;
                        }
                        DWORD avail = 0;
                        out.clear();
                        do
                        {
                            avail = 0;
                            if (!WinHttpQueryDataAvailable(hr, &avail) || avail == 0) break;
                            std::string chunk(avail, '\0');
                            DWORD got = 0;
                            if (!WinHttpReadData(hr, chunk.data(), avail, &got)) break;
                            out.append(chunk.data(), got);
                        } while (avail > 0);
                        ok = true;
                    }
                    WinHttpCloseHandle(hr);
                }
                WinHttpCloseHandle(hc);
            }
            WinHttpCloseHandle(hs);
            return ok;
        }

        // ── minimal MSXML helpers (SEARCHADDRESS) ───────────────────────────
        struct XmlDoc
        {
            IXMLDOMDocument2* doc = nullptr;
            ~XmlDoc() { if (doc) doc->Release(); }
            bool load(const std::wstring& xml)
            {
                if (FAILED(CoCreateInstance(__uuidof(DOMDocument60), nullptr,
                        CLSCTX_INPROC_SERVER, __uuidof(IXMLDOMDocument2), (void**)&doc)) || !doc)
                    return false;
                doc->put_async(VARIANT_FALSE);
                doc->put_validateOnParse(VARIANT_FALSE);
                BSTR b = SysAllocString(xml.c_str());
                VARIANT_BOOL okB = VARIANT_FALSE;
                doc->loadXML(b, &okB);
                SysFreeString(b);
                return okB == VARIANT_TRUE;
            }
            std::wstring text(const wchar_t* xpath)
            {
                if (!doc) return L"";
                IXMLDOMNode* node = nullptr;
                BSTR q = SysAllocString(xpath);
                doc->selectSingleNode(q, &node);
                SysFreeString(q);
                if (!node) return L"";
                BSTR t = nullptr;
                node->get_text(&t);
                std::wstring out = t ? t : L"";
                if (t) SysFreeString(t);
                node->Release();
                return out;
            }
        };

        bool looksLikeKeyError(const std::wstring& msg)
        {
            return msg.find(L"승인") != std::wstring::npos ||
                   msg.find(L"만료") != std::wstring::npos ||
                   msg.find(L"인증") != std::wstring::npos ||
                   msg.find(L"등록되지") != std::wstring::npos ||
                   msg.find(L"KEY") != std::wstring::npos ||
                   msg.find(L"key") != std::wstring::npos;
        }

        // VWorld shared bits
        const long kEpsgList[] = { 4326, 4019, 3857, 900913, 5180, 5185, 5181, 5186,
                                   5182, 5183, 5187, 5184, 5188, 5179 };
        bool epsgOk(long e)
        {
            for (long v : kEpsgList) if (v == e) return true;
            return false;
        }
        constexpr const wchar_t* kVworldIssuer = L"vworld.kr(공간정보 오픈플랫폼)";
        constexpr const wchar_t* kJusoIssuer = L"business.juso.go.kr(주소기반산업지원서비스)";
        constexpr const wchar_t* kDataIssuer = L"data.go.kr(공공데이터포털, 국세청 사업자등록정보 상태조회)";

        // GET + JSON parse for VWorld; returns false + err on transport failure.
        bool vworldGet(const std::wstring& url, json& j, std::wstring& err)
        {
            std::string body;
            DWORD status = 0;
            if (!httpCall(url, L"GET", "", nullptr, body, &status))
            {
                err = L"ERROR: 네트워크 요청 실패";
                return false;
            }
            try { j = json::parse(body); }
            catch (...) { err = L"ERROR: 응답 파싱 실패(HTTP " + std::to_wstring(status) + L")"; return false; }
            return true;
        }

        double jnum(const json& v)
        {
            if (v.is_number()) return v.get<double>();
            if (v.is_string()) return std::stod(v.get<std::string>());
            throw std::runtime_error("not a number");
        }

        std::wstring jstr(const json& j)
        {
            if (j.is_string()) return fromUtf8Pa(j.get<std::string>());
            if (j.is_number()) return std::to_wstring(j.get<double>());
            return L"";
        }
    }

    void registerPublicApi()
    {
        namespace core = egtools::core;

        // ── SEARCHADDRESS(text, [info], [row], [api_key]) ────────────────────
        core::registerFn(L"SEARCHADDRESS",
            [](const ExcelObj& textA, const ExcelObj& infoA, const ExcelObj& rowA,
               const ExcelObj& keyA) -> ExcelObj*
            {
                const std::wstring text = trimWs(textA.toString());
                if (text.empty()) return returnValue(CellError::Value);
                const std::wstring key = resolveKey(keyA, L"juso");
                if (key.empty()) return keyMissingError(L"juso", kJusoIssuer);

                // output column selection (1..27), default {2}=roadAddr
                std::vector<int> info;
                if (infoA.isType(ExcelType::Multi))
                {
                    ExcelArray a(infoA);
                    for (size_t k = 0; k < a.size(); ++k)
                        info.push_back(a.at((ExcelArray::row_t)(k / a.nCols()), (ExcelArray::col_t)(k % a.nCols())).get<int>(0));
                }
                else if (!infoA.isMissing()) info.push_back(infoA.get<int>(2));
                if (info.empty()) info.push_back(2);
                for (int c : info) if (c < 1 || c > 27) return returnValue(CellError::Value);
                const int iRow = rowA.get<int>(0);
                if (iRow < 0) return returnValue(CellError::Value);

                const std::wstring url =
                    L"https://business.juso.go.kr/addrlink/addrLinkApi.do?confmKey=" + key +
                    L"&currentPage=1&countPerPage=100&resultType=xml&hstryYn=N&firstSort=none"
                    L"&addInfoYn=Y&keyword=" + urlEncodeW(text);

                std::string body;
                if (!httpCall(url, L"GET", "", nullptr, body))
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: 네트워크 요청 실패")));

                XmlDoc doc;
                if (!doc.load(fromUtf8Pa(body)))
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: 응답 파싱 실패")));

                const std::wstring errCode = doc.text(L"/results/common/errorCode");
                const std::wstring errMsg = doc.text(L"/results/common/errorMessage");
                if (!errCode.empty() && errCode != L"0")
                {
                    if (looksLikeKeyError(errMsg) || errCode == L"E0001")
                        return keyRejectedError(L"juso", kJusoIssuer);
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: " + errMsg)));
                }
                const std::wstring total = doc.text(L"/results/common/totalCount");
                if (total.empty() || total == L"0") return returnValue(CellError::NA);

                static const wchar_t* kFields[27] = {
                    L"zipNo", L"roadAddr", L"jibunAddr", L"engAddr", L"roadAddrPart1",
                    L"roadAddrPart2", L"admCd", L"rnMgtSn", L"bdMgtSn", L"detBdNmList",
                    L"bdNm", L"bdKdcd", L"siNm", L"sggNm", L"emdNm", L"liNm", L"rn",
                    L"udrtYn", L"buldMnnm", L"buldSlno", L"mtYn", L"lnbrMnnm",
                    L"lnbrSlno", L"emdNo", L"hstryYn", L"relJibun", L"hemdNm" };

                // gather rows via selectNodes
                IXMLDOMNodeList* list = nullptr;
                BSTR q = SysAllocString(L"/results/juso");
                doc.doc->selectNodes(q, &list);
                SysFreeString(q);
                if (!list) return returnValue(CellError::NA);
                long nRows = 0;
                list->get_length(&nRows);
                if (nRows == 0) { list->Release(); return returnValue(CellError::NA); }
                if (iRow > nRows) { list->Release(); return returnValue(CellError::NA); }

                std::vector<std::vector<std::wstring>> table(
                    (size_t)nRows, std::vector<std::wstring>(27));
                for (long r = 0; r < nRows; ++r)
                {
                    IXMLDOMNode* row = nullptr;
                    list->get_item(r, &row);
                    if (!row) continue;
                    IXMLDOMNodeList* kids = nullptr;
                    row->get_childNodes(&kids);
                    if (kids)
                    {
                        long nk = 0; kids->get_length(&nk);
                        for (long k = 0; k < nk; ++k)
                        {
                            IXMLDOMNode* kid = nullptr;
                            kids->get_item(k, &kid);
                            if (!kid) continue;
                            BSTR nm = nullptr; kid->get_nodeName(&nm);
                            BSTR tx = nullptr; kid->get_text(&tx);
                            if (nm)
                            {
                                for (int f = 0; f < 27; ++f)
                                    if (_wcsicmp(nm, kFields[f]) == 0)
                                    { table[(size_t)r][(size_t)f] = tx ? tx : L""; break; }
                            }
                            if (nm) SysFreeString(nm);
                            if (tx) SysFreeString(tx);
                            kid->Release();
                        }
                        kids->Release();
                    }
                    row->Release();
                }
                list->Release();

                const long r0 = iRow == 0 ? 0 : iRow - 1;
                const long r1 = iRow == 0 ? nRows : iRow;
                std::vector<ExcelObj> flat;
                for (long r = r0; r < r1; ++r)
                    for (int c : info)
                        flat.emplace_back(std::wstring_view(table[(size_t)r][(size_t)c - 1]));
                return core::output(core::makeArray(
                    (ExcelArrayBuilder::row_t)(r1 - r0),
                    (ExcelArrayBuilder::col_t)info.size(), flat));
            },
            /*macro*/ false, /*threadsafe*/ false);

        // ── BRNSTATUS(numbers, [api_key]) ────────────────────────────────────
        core::registerFn(L"BRNSTATUS",
            [](const ExcelObj& numsA, const ExcelObj& keyA) -> ExcelObj*
            {
                std::wstring key = resolveKey(keyA, L"odcloud");
                if (key.empty()) return keyMissingError(L"odcloud", kDataIssuer);
                // accept both encoded and raw keys (raw contains '/', '+', '=')
                if (key.find(L'%') == std::wstring::npos &&
                    key.find_first_of(L"+/=") != std::wstring::npos)
                    key = urlEncodeW(key);

                std::vector<std::wstring> nums;
                std::vector<bool> blank;
                if (numsA.isType(ExcelType::Multi))
                {
                    ExcelArray a(numsA);
                    for (ExcelArray::row_t r = 0; r < a.nRows(); ++r)
                        for (ExcelArray::col_t c = 0; c < a.nCols(); ++c)
                        {
                            std::wstring s = trimWs(ExcelObj(a.at(r, c)).toString());
                            std::wstring digits;
                            for (wchar_t ch : s) if (ch != L'-') digits.push_back(ch);
                            blank.push_back(digits.empty());
                            nums.push_back(digits);
                        }
                }
                else
                {
                    std::wstring s = trimWs(numsA.toString());
                    std::wstring digits;
                    for (wchar_t ch : s) if (ch != L'-') digits.push_back(ch);
                    blank.push_back(digits.empty());
                    nums.push_back(digits);
                }
                if (nums.empty()) return returnValue(CellError::Value);

                const std::wstring url =
                    L"https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=" + key;

                std::vector<std::wstring> out(nums.size());
                for (size_t base = 0; base < nums.size(); base += 100)
                {
                    const size_t hi = std::min(base + 100, nums.size());
                    json req;
                    req["b_no"] = json::array();
                    for (size_t k = base; k < hi; ++k)
                        req["b_no"].push_back(toUtf8Pa(nums[k]));

                    std::string body;
                    DWORD status = 0;
                    if (!httpCall(url, L"POST", req.dump(), L"application/json", body, &status))
                        return returnValue(ExcelObj(std::wstring_view(L"ERROR: 네트워크 요청 실패")));
                    if (status == 401 || status == 403)
                        return keyRejectedError(L"odcloud", kDataIssuer);
                    json j;
                    try { j = json::parse(body); }
                    catch (...)
                    {
                        const std::wstring wb = fromUtf8Pa(body);
                        if (looksLikeKeyError(wb)) return keyRejectedError(L"odcloud", kDataIssuer);
                        return returnValue(ExcelObj(std::wstring_view(
                            L"ERROR: 응답 파싱 실패(HTTP " + std::to_wstring(status) + L")")));
                    }
                    if (!j.contains("data") || !j["data"].is_array())
                    {
                        std::wstring m = j.contains("msg") ? jstr(j["msg"]) : L"응답 형식 오류";
                        if (looksLikeKeyError(m)) return keyRejectedError(L"odcloud", kDataIssuer);
                        return returnValue(ExcelObj(std::wstring_view(L"ERROR: " + m)));
                    }
                    const auto& data = j["data"];
                    for (size_t k = 0; k < data.size() && base + k < nums.size(); ++k)
                    {
                        if (blank[base + k]) { out[base + k] = L""; continue; }
                        const auto& item = data[k];
                        out[base + k] = item.contains("tax_type") ? jstr(item["tax_type"]) : L"#N/A";
                    }
                }

                std::vector<ExcelObj> flat;
                for (auto& s : out) flat.emplace_back(std::wstring_view(s));
                // preserve input shape when it was a 2-D array
                if (numsA.isType(ExcelType::Multi))
                {
                    ExcelArray a(numsA);
                    return core::output(core::makeArray(a.nRows(), a.nCols(), flat));
                }
                return returnValue(ExcelObj(std::wstring_view(out[0])));
            },
            /*macro*/ false, /*threadsafe*/ false);

        // ── GEOSEARCH(keyword, [type], [epsg], [api_key]) ────────────────────
        core::registerFn(L"GEOSEARCH",
            [](const ExcelObj& kwA, const ExcelObj& typeA, const ExcelObj& epsgA,
               const ExcelObj& keyA) -> ExcelObj*
            {
                const std::wstring kw = trimWs(kwA.toString());
                if (kw.empty()) return returnValue(CellError::Value);
                int type = typeA.get<int>(1);
                if (type == 0) type = 1;
                if (type < 1 || type > 2) return returnValue(CellError::Value);
                long epsg = epsgA.get<int>(4326);
                if (epsg == 0) epsg = 4326;
                if (!epsgOk(epsg)) return returnValue(CellError::Value);
                const std::wstring key = resolveKey(keyA, L"vworld");
                if (key.empty()) return keyMissingError(L"vworld", kVworldIssuer);

                auto makeUrl = [&](bool parcel) {
                    std::wstring u = L"https://api.vworld.kr/req/search?key=" + key +
                        L"&service=search&version=2.0&request=search&format=json&size=1000"
                        L"&crs=EPSG:" + std::to_wstring(epsg) + L"&query=" + urlEncodeW(kw);
                    u += (type == 1)
                        ? (parcel ? L"&type=address&category=parcel" : L"&type=address&category=road")
                        : L"&type=place";
                    return u;
                };

                json j; std::wstring err;
                if (!vworldGet(makeUrl(false), j, err))
                    return returnValue(ExcelObj(std::wstring_view(err)));
                auto status = [&](const json& x) -> std::wstring {
                    try { return fromUtf8Pa(x.at("response").at("status").get<std::string>()); }
                    catch (...) { return L""; }
                };
                if (type == 1 && status(j) != L"OK")
                    if (!vworldGet(makeUrl(true), j, err))
                        return returnValue(ExcelObj(std::wstring_view(err)));

                const std::wstring st = status(j);
                if (st == L"NOT_FOUND") return returnValue(CellError::NA);
                if (st != L"OK")
                {
                    std::wstring m;
                    try { m = fromUtf8Pa(j["response"]["error"]["text"].get<std::string>()); }
                    catch (...) { m = L"응답 오류"; }
                    if (looksLikeKeyError(m)) return keyRejectedError(L"vworld", kVworldIssuer);
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: " + m)));
                }

                try
                {
                    const auto& items = j["response"]["result"]["items"];
                    const size_t n = items.size();
                    if (n == 0) return returnValue(CellError::NA);
                    const bool withAddr = n > 1;
                    std::vector<ExcelObj> flat;
                    for (const auto& it : items)
                    {
                        flat.emplace_back(jnum(it["point"]["x"]));
                        flat.emplace_back(jnum(it["point"]["y"]));
                        if (withAddr)
                        {
                            std::wstring addr;
                            if (it.contains("address"))
                            {
                                const auto& ad = it["address"];
                                if (ad.contains("road") && ad["road"].is_string())
                                    addr = fromUtf8Pa(ad["road"].get<std::string>());
                                else if (ad.contains("parcel") && ad["parcel"].is_string())
                                    addr = fromUtf8Pa(ad["parcel"].get<std::string>());
                                if (ad.contains("bldnm") && ad["bldnm"].is_string())
                                {
                                    const std::wstring b = fromUtf8Pa(ad["bldnm"].get<std::string>());
                                    if (!b.empty() && addr.find(b) == std::wstring::npos)
                                        addr += L" " + b;
                                }
                            }
                            flat.emplace_back(std::wstring_view(trimWs(addr)));
                        }
                    }
                    return core::output(core::makeArray(
                        (ExcelArrayBuilder::row_t)n,
                        (ExcelArrayBuilder::col_t)(withAddr ? 3 : 2), flat));
                }
                catch (...) { return returnValue(CellError::NA); }
            },
            /*macro*/ false, /*threadsafe*/ false);

        // ── GEOCODER(keyword, [epsg], [api_key]) ─────────────────────────────
        core::registerFn(L"GEOCODER",
            [](const ExcelObj& kwA, const ExcelObj& epsgA, const ExcelObj& keyA) -> ExcelObj*
            {
                const std::wstring kw = trimWs(kwA.toString());
                if (kw.empty()) return returnValue(CellError::Value);
                long epsg = epsgA.get<int>(4326);
                if (epsg == 0) epsg = 4326;
                if (!epsgOk(epsg)) return returnValue(CellError::Value);
                const std::wstring key = resolveKey(keyA, L"vworld");
                if (key.empty()) return keyMissingError(L"vworld", kVworldIssuer);

                auto makeUrl = [&](const wchar_t* type) {
                    return L"https://api.vworld.kr/req/address?key=" + key +
                        L"&service=address&version=2.0&request=GetCoord&format=json&crs=EPSG:" +
                        std::to_wstring(epsg) + L"&address=" + urlEncodeW(kw) + L"&type=" + type;
                };
                json j; std::wstring err;
                if (!vworldGet(makeUrl(L"road"), j, err))
                    return returnValue(ExcelObj(std::wstring_view(err)));
                auto status = [&](const json& x) -> std::wstring {
                    try { return fromUtf8Pa(x.at("response").at("status").get<std::string>()); }
                    catch (...) { return L""; }
                };
                if (status(j) != L"OK")
                    if (!vworldGet(makeUrl(L"parcel"), j, err))
                        return returnValue(ExcelObj(std::wstring_view(err)));

                const std::wstring st = status(j);
                if (st == L"NOT_FOUND") return returnValue(CellError::NA);
                if (st != L"OK")
                {
                    std::wstring m;
                    try { m = fromUtf8Pa(j["response"]["error"]["text"].get<std::string>()); }
                    catch (...) { m = L"응답 오류"; }
                    if (looksLikeKeyError(m)) return keyRejectedError(L"vworld", kVworldIssuer);
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: " + m)));
                }
                try
                {
                    const auto& pt = j["response"]["result"]["point"];
                    std::vector<ExcelObj> flat;
                    flat.emplace_back(jnum(pt["x"]));
                    flat.emplace_back(jnum(pt["y"]));
                    std::wstring refined;
                    try { refined = fromUtf8Pa(j["response"]["refined"]["text"].get<std::string>()); }
                    catch (...) {}
                    flat.emplace_back(std::wstring_view(refined));
                    return core::output(core::makeArray(1, 3, flat));
                }
                catch (...) { return returnValue(CellError::NA); }
            },
            /*macro*/ false, /*threadsafe*/ false);

        // ── GEOADDRESS(x, y, [return_type], [epsg], [api_key]) ───────────────
        core::registerFn(L"GEOADDRESS",
            [](const ExcelObj& xA, const ExcelObj& yA, const ExcelObj& rtA,
               const ExcelObj& epsgA, const ExcelObj& keyA) -> ExcelObj*
            {
                const double x = xA.get<double>(0.0), y = yA.get<double>(0.0);
                int rt = rtA.get<int>(1);
                if (rt == 0) rt = 1;
                if (rt < 1 || rt > 3) return returnValue(CellError::Value);
                long epsg = epsgA.get<int>(4326);
                if (epsg == 0) epsg = 4326;
                if (!epsgOk(epsg)) return returnValue(CellError::Value);
                if (epsg == 4326 && (std::abs(y) > 90 || std::abs(x) > 180))
                    return returnValue(CellError::Value);
                const std::wstring key = resolveKey(keyA, L"vworld");
                if (key.empty()) return keyMissingError(L"vworld", kVworldIssuer);

                std::wstring url = L"https://api.vworld.kr/req/address?key=" + key +
                    L"&service=address&version=2.0&request=GetAddress&format=json&crs=EPSG:" +
                    std::to_wstring(epsg) + L"&point=" + std::to_wstring(x) + L"," + std::to_wstring(y);
                url += (rt == 1) ? L"&type=parcel" : (rt == 2) ? L"&type=road" : L"&type=both";

                json j; std::wstring err;
                if (!vworldGet(url, j, err))
                    return returnValue(ExcelObj(std::wstring_view(err)));
                std::wstring st;
                try { st = fromUtf8Pa(j.at("response").at("status").get<std::string>()); }
                catch (...) {}
                if (st == L"NOT_FOUND") return returnValue(CellError::NA);
                if (st != L"OK")
                {
                    std::wstring m;
                    try { m = fromUtf8Pa(j["response"]["error"]["text"].get<std::string>()); }
                    catch (...) { m = L"응답 오류"; }
                    if (looksLikeKeyError(m)) return keyRejectedError(L"vworld", kVworldIssuer);
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: " + m)));
                }
                try
                {
                    const auto& res = j["response"]["result"];
                    std::vector<ExcelObj> flat;
                    size_t n = 0;
                    for (const auto& item : res)
                    {
                        flat.emplace_back(std::wstring_view(
                            fromUtf8Pa(item["text"].get<std::string>())));
                        if (++n >= (rt == 3 ? 2u : 1u)) break;
                    }
                    if (flat.empty()) return returnValue(CellError::NA);
                    return core::output(core::makeArray(
                        (ExcelArrayBuilder::row_t)flat.size(), 1, flat));
                }
                catch (...) { return returnValue(CellError::NA); }
            },
            /*macro*/ false, /*threadsafe*/ false);

        // ── ADDRESSMAP(address, [scale], [api_key]) — macro (picture insert) ─
        core::registerFn(L"ADDRESSMAP",
            [](const ExcelObj& addrA, const ExcelObj& scaleA, const ExcelObj& keyA) -> ExcelObj*
            {
                const std::wstring addr = trimWs(addrA.toString());
                if (addr.empty()) return returnValue(CellError::Value);
                int scale = scaleA.get<int>(5);
                if (scale == 0) scale = 5;
                if (scale < 1 || scale > 10) return returnValue(CellError::Value);
                const std::wstring key = resolveKey(keyA, L"vworld");
                if (key.empty()) return keyMissingError(L"vworld", kVworldIssuer);

                // geocode (road → parcel), same as GEOCODER
                auto makeUrl = [&](const wchar_t* type) {
                    return L"https://api.vworld.kr/req/address?key=" + key +
                        L"&service=address&version=2.0&request=GetCoord&format=json&crs=EPSG:4326"
                        L"&address=" + urlEncodeW(addr) + L"&type=" + type;
                };
                json j; std::wstring err;
                if (!vworldGet(makeUrl(L"road"), j, err))
                    return returnValue(ExcelObj(std::wstring_view(err)));
                std::wstring st;
                try { st = fromUtf8Pa(j.at("response").at("status").get<std::string>()); }
                catch (...) {}
                if (st != L"OK")
                {
                    if (!vworldGet(makeUrl(L"parcel"), j, err))
                        return returnValue(ExcelObj(std::wstring_view(err)));
                    try { st = fromUtf8Pa(j.at("response").at("status").get<std::string>()); }
                    catch (...) { st.clear(); }
                }
                if (st == L"NOT_FOUND") return returnValue(CellError::NA);
                if (st != L"OK")
                {
                    std::wstring m;
                    try { m = fromUtf8Pa(j["response"]["error"]["text"].get<std::string>()); }
                    catch (...) { m = L"응답 오류"; }
                    if (looksLikeKeyError(m)) return keyRejectedError(L"vworld", kVworldIssuer);
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: " + m)));
                }

                double x = 0, y = 0;
                try
                {
                    const auto& pt = j["response"]["result"]["point"];
                    x = jnum(pt["x"]);
                    y = jnum(pt["y"]);
                }
                catch (...) { return returnValue(CellError::NA); }

                const double dx = (2.0 / scale) / (90 * 4), dy = (2.0 / scale) / (110 * 4);
                auto num = [](double v) {
                    wchar_t b[32];
                    swprintf_s(b, L"%.8f", v);
                    return std::wstring(b);
                };
                std::wstring wms = L"https://api.vworld.kr/req/wms?key=" + key +
                    L"&request=GetMap&version=1.3.0&crs=EPSG:4326&format=image/png&transparent=true"
                    L"&width=1110&height=920&bbox=" + num(y - dy) + L"," + num(x - dx) + L"," +
                    num(y + dy) + L"," + num(x + dx) +
                    L"&layers=lt_c_spbd,lt_l_sprd,lp_pa_cbnd_bonbun,lp_pa_cbnd_bubun"
                    L"&styles=lt_c_spbd,lt_l_sprd,lp_pa_cbnd_bonbun,lp_pa_cbnd_bubun";

                std::string png;
                if (!httpCall(wms, L"GET", "", nullptr, png) || png.size() < 100 ||
                    png.compare(0, 4, "\x89PNG") != 0)
                    return returnValue(ExcelObj(std::wstring_view(L"ERROR: 지도 이미지 수신 실패")));

                wchar_t tmp[MAX_PATH]{};
                GetTempPathW(MAX_PATH, tmp);
                const std::wstring path = std::wstring(tmp) + L"egmap_" +
                    std::to_wstring(GetTickCount64()) + L".png";
                FILE* f = nullptr;
                if (_wfopen_s(&f, path.c_str(), L"wb") != 0 || !f)
                    return returnValue(CellError::Value);
                fwrite(png.data(), 1, png.size(), f);
                fclose(f);

                const std::wstring fullAddr = xloil::CallerInfo().address();
                const size_t bang = fullAddr.rfind(L'!');
                const std::wstring cell = bang == std::wstring::npos
                                              ? fullAddr : fullAddr.substr(bang + 1);
                if (cell.empty()) { DeleteFileW(path.c_str()); return returnValue(CellError::Ref); }
                queueInsertPicture(fullAddr, cell, path, 1, /*deleteLocalAfter*/ true);
                return returnValue(ExcelObj(std::wstring_view(L"")));
            },
            /*macro*/ true, /*threadsafe*/ false);
    }
}
