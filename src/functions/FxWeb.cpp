// FxWeb.cpp — web / text-encoding functions.
//   ENCODEURL  — UTF-8 percent-encoding (no network).
//   FILTERXML  — XPath query over XML, via MSXML6 (COM, no network).
//   WEBSERVICE — HTTP(S) GET, via WinHTTP.

#include "../core/Registry.h"
#include "../core/Apply.h"
#include "../core/Spill.h"

#include <windows.h>
#include <winhttp.h>
#include <msxml6.h>

#include <xlOil/xlOil.h>
#include <xlOil/ArrayBuilder.h>
#include <string>
#include <vector>

#pragma comment(lib, "winhttp.lib")

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        // ---- ENCODEURL ----------------------------------------------------
        ExcelObj encodeUrlOne(const ExcelObj& e)
        {
            const std::wstring w = e.toString();
            int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), nullptr, 0, nullptr, nullptr);
            std::string u8((size_t)n, '\0');
            WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(), u8.data(), n, nullptr, nullptr);

            static const char* kHex = "0123456789ABCDEF";
            std::wstring out;
            out.reserve(u8.size());
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
            return ExcelObj(std::wstring_view(out));
        }

        // ---- FILTERXML ----------------------------------------------------
        ExcelObj* filterXml(const ExcelObj& xmlObj, const ExcelObj& xpathObj)
        {
            const std::wstring xml = xmlObj.toString();
            const std::wstring xpath = xpathObj.toString();
            if (xml.empty() || xpath.empty()) return returnValue(CellError::Value);

            IXMLDOMDocument2* doc = nullptr;
            HRESULT hr = CoCreateInstance(__uuidof(DOMDocument60), nullptr,
                CLSCTX_INPROC_SERVER, __uuidof(IXMLDOMDocument2), (void**)&doc);
            if (FAILED(hr) || !doc) return returnValue(CellError::Value);

            ExcelObj* result = returnValue(CellError::Value);
            doc->put_async(VARIANT_FALSE);
            doc->put_validateOnParse(VARIANT_FALSE);

            VARIANT_BOOL ok = VARIANT_FALSE;
            BSTR bxml = SysAllocString(xml.c_str());
            doc->loadXML(bxml, &ok);
            SysFreeString(bxml);

            if (ok == VARIANT_TRUE)
            {
                IXMLDOMNodeList* list = nullptr;
                BSTR bxp = SysAllocString(xpath.c_str());
                hr = doc->selectNodes(bxp, &list);
                SysFreeString(bxp);
                if (SUCCEEDED(hr) && list)
                {
                    long n = 0; list->get_length(&n);
                    std::vector<std::wstring> vals;
                    for (long i = 0; i < n; ++i)
                    {
                        IXMLDOMNode* node = nullptr;
                        if (list->get_item(i, &node) == S_OK && node)
                        {
                            BSTR t = nullptr;
                            node->get_text(&t);
                            vals.emplace_back(t ? t : L"");
                            if (t) SysFreeString(t);
                            node->Release();
                        }
                    }
                    list->Release();

                    if (vals.empty())
                        result = returnValue(CellError::NA);
                    else if (vals.size() == 1)
                        result = returnValue(ExcelObj(std::wstring_view(vals[0])));
                    else
                    {
                        size_t strLen = 0;
                        for (auto& s : vals) strLen += s.size();
                        ExcelArrayBuilder b((ExcelArrayBuilder::row_t)vals.size(), 1, strLen);
                        for (size_t i = 0; i < vals.size(); ++i)
                            b((ExcelArrayBuilder::row_t)i, 0) = ExcelObj(std::wstring_view(vals[i]));
                        result = egtools::core::output(b.toExcelObj());
                    }
                }
            }
            doc->Release();
            return result;
        }

        // ---- WEBSERVICE ---------------------------------------------------
        ExcelObj webservice(const ExcelObj& urlObj)
        {
            const std::wstring url = urlObj.toString();
            URL_COMPONENTS uc{}; uc.dwStructSize = sizeof(uc);
            wchar_t host[256]{}, path[4096]{};
            uc.lpszHostName = host; uc.dwHostNameLength = 255;
            uc.lpszUrlPath = path; uc.dwUrlPathLength = 4095;
            if (!WinHttpCrackUrl(url.c_str(), (DWORD)url.size(), 0, &uc))
                return ExcelObj(CellError::Value);
            host[uc.dwHostNameLength] = L'\0';

            HINTERNET hSession = WinHttpOpen(L"EGTools++/1.0",
                WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);
            if (!hSession) return ExcelObj(CellError::Value);
            WinHttpSetTimeouts(hSession, 5000, 5000, 8000, 15000);

            ExcelObj out(CellError::Value);
            HINTERNET hConnect = WinHttpConnect(hSession, host, uc.nPort, 0);
            if (hConnect)
            {
                DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
                HINTERNET hReq = WinHttpOpenRequest(hConnect, L"GET", path, nullptr,
                    WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
                if (hReq)
                {
                    if (WinHttpSendRequest(hReq, WINHTTP_NO_ADDITIONAL_HEADERS, 0,
                            WINHTTP_NO_REQUEST_DATA, 0, 0, 0) &&
                        WinHttpReceiveResponse(hReq, nullptr))
                    {
                        std::string body;
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

                        int n = MultiByteToWideChar(CP_UTF8, 0, body.data(), (int)body.size(), nullptr, 0);
                        std::wstring w((size_t)n, L'\0');
                        MultiByteToWideChar(CP_UTF8, 0, body.data(), (int)body.size(), w.data(), n);
                        out = ExcelObj(std::wstring_view(w));
                    }
                    WinHttpCloseHandle(hReq);
                }
                WinHttpCloseHandle(hConnect);
            }
            WinHttpCloseHandle(hSession);
            return out;
        }
    }

    void registerWeb()
    {
        egtools::core::registerFn(L"ENCODEURL",
            [](const ExcelObj& t) -> ExcelObj* { return egtools::core::mapUnary(t, encodeUrlOne); });

        egtools::core::registerFn(L"FILTERXML",
            [](const ExcelObj& xml, const ExcelObj& xpath) -> ExcelObj* { return filterXml(xml, xpath); },
            /*macro*/ false, /*threadsafe*/ false);   // MSXML COM — MTR 워커 불가

        egtools::core::registerFn(L"WEBSERVICE",
            [](const ExcelObj& url) -> ExcelObj* { return returnValue(webservice(url)); },
            /*macro*/ false, /*threadsafe*/ true);    // 스칼라 네트워크 — 병렬 효과 큼
    }
}
