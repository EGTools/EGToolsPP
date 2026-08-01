// FxEgUtil.cpp — EGTools 유틸리티 함수 (EGToolsVB FX9_EGTools + FX7 일부 포팅, plan/20 E5·E8).
//   EVAL       — 수식 문자열 계산 (매크로형, COM Application.Evaluate — FxLet과 동일 경로).
//   DIRFOLDER  — 폴더의 파일/하위폴더 목록 (Win32, 비매크로).
//   SHEETLIST  — 호출 통합문서의 (보이는) 시트 이름 목록 (매크로형, late-bound COM).
//   TOTALPAGES — 현재 인쇄 설정 기준 전체 페이지 수 (매크로형, XLM GET.DOCUMENT).
//   RESTAPI    — 임의 메서드/헤더/본문 HTTP 요청 (WinHTTP, 비매크로).
//   EXRATE     — 지정일 기준환율 조회 (smbs.biz, 최대 10일 소급, 비매크로).
// GEOSEARCH/SEARCHADDRESS/BRNSTATUS 등 API 키가 필요한 함수는 보류(plan/20 §4).

#include "../core/Registry.h"
#include "../core/ArrayUtil.h"
#include "../core/Spill.h"

#include <windows.h>
#include <winhttp.h>
#include <oleauto.h>

#include <xlOil/xlOil.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/AppObjects.h>   // thisApp()
#include <xlOil/ExcelUI.h>      // variantToExcelObj

#include <algorithm>
#include <cwctype>
#include <string>
#include <vector>

#pragma comment(lib, "winhttp.lib")

using namespace xloil;

namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::functions
{
    namespace
    {
        constexpr int kXlfCaller = 89;
        constexpr int kXlfGetDocument = 88;
        constexpr int kXlSheetNm = 0x4005;   // 참조 → "[통합문서]시트" 이름

        std::wstring trimW(const std::wstring& s)
        {
            const wchar_t* ws = L" \t\r\n";
            const size_t b = s.find_first_not_of(ws);
            if (b == std::wstring::npos) return {};
            return s.substr(b, s.find_last_not_of(ws) - b + 1);
        }

        // ---- HTTP (RESTAPI / EXRATE) --------------------------------------

        bool httpRequest(const std::wstring& url, const std::wstring& method,
                         const std::vector<std::pair<std::wstring, std::wstring>>& headers,
                         const std::string& body, std::string& out)
        {
            URL_COMPONENTS uc{}; uc.dwStructSize = sizeof(uc);
            std::wstring host((size_t)url.size() + 1, L'\0');
            std::wstring path((size_t)url.size() + 1, L'\0');
            uc.lpszHostName = host.data(); uc.dwHostNameLength = (DWORD)url.size();
            uc.lpszUrlPath = path.data(); uc.dwUrlPathLength = (DWORD)url.size();
            if (!WinHttpCrackUrl(url.c_str(), (DWORD)url.size(), 0, &uc)) return false;
            host.resize(uc.dwHostNameLength);

            HINTERNET hSession = WinHttpOpen(L"EGTools++/1.0 (WinHTTP)",
                WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME,
                WINHTTP_NO_PROXY_BYPASS, 0);
            if (!hSession) return false;
            WinHttpSetTimeouts(hSession, 8000, 8000, 15000, 30000);

            bool ok = false;
            HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), uc.nPort, 0);
            if (hConnect)
            {
                const DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
                HINTERNET hReq = WinHttpOpenRequest(hConnect, method.c_str(), uc.lpszUrlPath,
                    nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
                if (hReq)
                {
                    std::wstring hdr;
                    for (const auto& [k, v] : headers) hdr += k + L": " + v + L"\r\n";
                    const BOOL sent = WinHttpSendRequest(hReq,
                        hdr.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : hdr.c_str(),
                        hdr.empty() ? 0 : (DWORD)-1,
                        body.empty() ? WINHTTP_NO_REQUEST_DATA : (LPVOID)body.data(),
                        (DWORD)body.size(), (DWORD)body.size(), 0);
                    if (sent && WinHttpReceiveResponse(hReq, nullptr))
                    {
                        DWORD avail = 0;
                        do
                        {
                            avail = 0;
                            if (!WinHttpQueryDataAvailable(hReq, &avail) || avail == 0) break;
                            std::string chunk((size_t)avail, '\0');
                            DWORD read = 0;
                            if (!WinHttpReadData(hReq, chunk.data(), avail, &read)) break;
                            out.append(chunk.data(), read);
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

        std::wstring utf8ToW(const std::string& s)
        {
            if (s.empty()) return {};
            int n = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
            std::wstring w((size_t)n, L'\0');
            MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), w.data(), n);
            return w;
        }

        ExcelObj* restApi(const ExcelObj& urlObj, const ExcelObj& methodObj,
                          const ExcelObj& headersObj, const ExcelObj& bodyObj)
        {
            if (urlObj.isMissing()) return returnValue(CellError::Value);
            const std::wstring url = trimW(urlObj.toString());
            if (url.empty()) return returnValue(CellError::Value);
            std::wstring method = methodObj.isMissing() ? L"GET" : trimW(methodObj.toString());
            if (method.empty()) method = L"GET";
            for (auto& c : method) c = (wchar_t)std::towupper(c);

            std::vector<std::pair<std::wstring, std::wstring>> headers;
            if (!headersObj.isMissing() && headersObj.isType(ExcelType::Multi))
            {
                ExcelArray h(headersObj);
                if (h.nCols() >= 2)
                    for (size_t r = 0; r < h.nRows(); ++r)
                    {
                        const std::wstring k = h.at(r * h.nCols()).toString();
                        const std::wstring v = h.at(r * h.nCols() + 1).toString();
                        if (!k.empty()) headers.emplace_back(k, v);
                    }
            }

            std::string body;
            if (!bodyObj.isMissing())
            {
                const std::wstring w = bodyObj.toString();
                if (!w.empty())
                {
                    int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(),
                                                nullptr, 0, nullptr, nullptr);
                    body.resize((size_t)n);
                    WideCharToMultiByte(CP_UTF8, 0, w.data(), (int)w.size(),
                                        body.data(), n, nullptr, nullptr);
                }
            }

            std::string resp;
            if (!httpRequest(url, method, headers, body, resp))
                return returnValue(CellError::Value);
            return returnValue(ExcelObj(std::wstring_view(utf8ToW(resp))));
        }

        // ---- EXRATE --------------------------------------------------------

        bool validCurrency(const std::wstring& c)
        {
            static const wchar_t* kList =
                L",USD,EUR,JPY,GBP,CHF,CAD,AUD,CNY,HKD,SEK,NZD,KRW,SGD,NOK,MXN,INR,RUB,"
                L"ZAR,TRY,BRL,AED,BHD,BND,CNH,CZK,DKK,IDR,ILS,MYR,QAR,SAR,THB,TWD,CLP,"
                L"COP,EGP,HUF,KWD,OMR,PHP,PLN,PKR,RON,BDT,DZD,ETB,FJD,JOD,KES,KHR,KZT,"
                L"LKR,LYD,MMK,MNT,MOP,NPR,TZS,UZS,VND,";
            return std::wstring(kList).find(L"," + c + L",") != std::wstring::npos;
        }

        bool smbsRate(const std::wstring& curr, int serial, double& rate)
        {
            // 직렬값 → 날짜 문자열: VariantTimeToSystemTime 사용.
            SYSTEMTIME st{};
            if (!VariantTimeToSystemTime((double)serial, &st)) return false;
            wchar_t dateBuf[16];
            swprintf_s(dateBuf, L"%04d-%02d-%02d", st.wYear, st.wMonth, st.wDay);

            const std::wstring url =
                L"http://www.smbs.biz/ExRate/StdExRate_xml.jsp?arr_value=" + curr +
                L"_" + dateBuf + L"_" + dateBuf;
            std::string body;
            if (!httpRequest(url, L"GET", {}, "", body)) return false;
            const std::wstring xml = utf8ToW(body);

            std::wstring low = xml;
            for (auto& c : low) c = (wchar_t)std::towlower(c);
            const size_t setPos = low.find(L"<set");
            if (setPos == std::wstring::npos) return false;
            const size_t valPos = low.find(L"value='", setPos);
            if (valPos == std::wstring::npos) return false;
            const size_t beg = valPos + 7;
            const size_t end = xml.find(L'\'', beg);
            if (end == std::wstring::npos) return false;
            std::wstring num = xml.substr(beg, end - beg);
            num.erase(std::remove(num.begin(), num.end(), L','), num.end());
            wchar_t* e = nullptr;
            rate = wcstod(num.c_str(), &e);
            return e && *e == L'\0' && rate > 0;
        }

        ExcelObj* exRate(const ExcelObj& currObj, const ExcelObj& dateObj)
        {
            std::wstring curr = currObj.isMissing() ? L"USD" : trimW(currObj.toString());
            if (curr.empty()) curr = L"USD";
            for (auto& c : curr) c = (wchar_t)std::towupper(c);
            if (curr == L"CNY") curr = L"CNH";
            if (!validCurrency(curr)) return returnValue(CellError::NA);

            SYSTEMTIME now;
            GetLocalTime(&now);
            double todayV = 0;
            SYSTEMTIME nd{ now.wYear, now.wMonth, 0, now.wDay, 0, 0, 0, 0 };
            SystemTimeToVariantTime(&nd, &todayV);
            const int today = (int)todayV;

            int target = today;
            if (!dateObj.isMissing() &&
                (dateObj.type() == ExcelType::Num || dateObj.type() == ExcelType::Int))
            {
                const int d = dateObj.get<int>(0);
                if (d > 2) target = d;
            }
            if (target > today) return returnValue(CellError::Value);

            for (int t = target; t >= target - 10; --t)
            {
                double rate;
                if (smbsRate(curr, t, rate)) return returnValue(ExcelObj(rate));
            }
            return returnValue(CellError::NA);
        }

        // ---- DIRFOLDER -----------------------------------------------------

        bool matchAnyFilter(const std::wstring& name,
                            const std::vector<std::wstring>& filters)
        {
            auto wildMatch = [](const std::wstring& text, const std::wstring& pat)
            {
                size_t t = 0, p = 0, star = std::wstring::npos, mark = 0;
                while (t < text.size())
                {
                    if (p < pat.size() && (pat[p] == L'?' ||
                        std::towupper(pat[p]) == std::towupper(text[t]))) { ++t; ++p; }
                    else if (p < pat.size() && pat[p] == L'*') { star = p++; mark = t; }
                    else if (star != std::wstring::npos) { p = star + 1; t = ++mark; }
                    else return false;
                }
                while (p < pat.size() && pat[p] == L'*') ++p;
                return p == pat.size();
            };
            for (const auto& f : filters)
                if (wildMatch(name, f)) return true;
            return false;
        }

        void listDir(const std::wstring& path, const std::vector<std::wstring>& filters,
                     bool incDir, bool incSub, bool incHidden,
                     std::vector<std::pair<int, std::wstring>>& out, int level)
        {
            WIN32_FIND_DATAW fd;
            std::vector<std::wstring> dirs;
            HANDLE h = FindFirstFileW((path + L"*").c_str(), &fd);
            if (h == INVALID_HANDLE_VALUE) return;
            do
            {
                const std::wstring name = fd.cFileName;
                if (name == L"." || name == L"..") continue;
                const bool hidden = (fd.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) != 0;
                if (hidden && !incHidden) continue;
                if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
                    dirs.push_back(name);
                else if (matchAnyFilter(name, filters))
                    out.emplace_back(level, name);
            } while (FindNextFileW(h, &fd));
            FindClose(h);

            if (incSub)
            {
                std::sort(dirs.begin(), dirs.end(),
                          [](const std::wstring& a, const std::wstring& b)
                          { return _wcsicmp(a.c_str(), b.c_str()) < 0; });
                for (const auto& d : dirs)
                {
                    if (incDir) out.emplace_back(level, d);
                    listDir(path + d + L"\\", filters, incDir, incSub, incHidden,
                            out, level + 1);
                }
            }
        }

        ExcelObj* dirFolder(const ExcelObj& pathObj, const ExcelObj& filterObj,
                            const ExcelObj& dirObj, const ExcelObj& subObj,
                            const ExcelObj& hiddenObj)
        {
            if (pathObj.isMissing()) return returnValue(CellError::Value);
            std::wstring path = trimW(pathObj.toString());
            if (path.empty()) return returnValue(CellError::Value);
            if (path.back() != L'\\' && path.back() != L'/') path += L'\\';

            std::wstring filter = filterObj.isMissing() ? L"*.*" : filterObj.toString();
            if (filter.empty()) filter = L"*.*";
            std::vector<std::wstring> filters;
            {
                std::wstring cur;
                for (wchar_t c : filter + L";")
                {
                    if (c == L';') { if (!cur.empty()) filters.push_back(cur); cur.clear(); }
                    else cur.push_back(c);
                }
            }
            const bool incDir = dirObj.isMissing() ? false : (dirObj.get<double>(0.0) != 0.0);
            const bool incSub = subObj.isMissing() ? false : (subObj.get<double>(0.0) != 0.0);
            const bool incHid = hiddenObj.isMissing() ? false : (hiddenObj.get<double>(0.0) != 0.0);

            std::vector<std::pair<int, std::wstring>> items;
            listDir(path, filters, incDir, incSub, incHid, items, 0);
            if (items.empty()) return returnValue(CellError::NA);

            int maxLevel = 0;
            for (const auto& [lvl, name] : items) maxLevel = (std::max)(maxLevel, lvl);

            std::vector<ExcelObj> vals;
            vals.reserve(items.size() * ((size_t)maxLevel + 1));
            for (const auto& [lvl, name] : items)
                for (int c = 0; c <= maxLevel; ++c)
                    vals.push_back(c == lvl ? ExcelObj(std::wstring_view(name))
                                            : ExcelObj(std::wstring_view(L"")));
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)items.size(),
                (ExcelArrayBuilder::col_t)(maxLevel + 1), vals));
        }

        // ---- EVAL (매크로형, COM Evaluate) ---------------------------------

        bool comEvaluate(const std::wstring& formula, ExcelObj& out)
        {
            IUnknown* unk = nullptr;
            try { xloil::COM::connectCom(); unk = xloil::thisApp().ptr(); }
            catch (...) { return false; }
            if (!unk) return false;
            IDispatch* disp = nullptr;
            if (FAILED(unk->QueryInterface(IID_IDispatch, (void**)&disp)) || !disp) return false;

            bool ok = false;
            OLECHAR* nm = (OLECHAR*)L"Evaluate";
            DISPID dispid = 0;
            if (SUCCEEDED(disp->GetIDsOfNames(IID_NULL, &nm, 1, LOCALE_USER_DEFAULT, &dispid)))
            {
                VARIANT arg; VariantInit(&arg);
                arg.vt = VT_BSTR; arg.bstrVal = SysAllocString(formula.c_str());
                DISPPARAMS params{ &arg, nullptr, 1, 0 };
                VARIANT result; VariantInit(&result);
                const HRESULT hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                                DISPATCH_METHOD, &params, &result, nullptr, nullptr);
                if (SUCCEEDED(hr))
                {
                    out = xloil::variantToExcelObj(result, false);
                    ok = (out.type() != ExcelType::Err);
                }
                VariantClear(&result);
                VariantClear(&arg);
            }
            disp->Release();
            return ok;
        }

        ExcelObj* evalFn(const ExcelObj& formulaObj)
        {
            if (formulaObj.isMissing()) return returnValue(CellError::Value);
            std::wstring f = trimW(formulaObj.toString());
            if (f.empty()) return returnValue(CellError::Value);
            if (f[0] != L'=') f = L"=" + f;

            ExcelObj result;
            if (!comEvaluate(f, result)) return returnValue(CellError::Value);
            if (result.type() == ExcelType::Multi)
                return egtools::core::output(std::move(result));
            return returnValue(std::move(result));
        }

        // ---- SHEETLIST (매크로형, late-bound COM) --------------------------

        HRESULT getDispProp(IDispatch* d, const wchar_t* name, VARIANT* args,
                            UINT nArgs, VARIANT* out)
        {
            DISPID id = 0;
            OLECHAR* nm = const_cast<OLECHAR*>(name);
            HRESULT hr = d->GetIDsOfNames(IID_NULL, &nm, 1, LOCALE_USER_DEFAULT, &id);
            if (FAILED(hr)) return hr;
            // DISPPARAMS는 역순.
            std::vector<VARIANT> rev;
            for (UINT i = 0; i < nArgs; ++i) rev.push_back(args[nArgs - 1 - i]);
            DISPPARAMS dp{ rev.empty() ? nullptr : rev.data(), nullptr, nArgs, 0 };
            return d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                             DISPATCH_PROPERTYGET | DISPATCH_METHOD, &dp, out, nullptr, nullptr);
        }

        ExcelObj* sheetList(const ExcelObj& startObj, const ExcelObj& endObj,
                            const ExcelObj& strictObj, const ExcelObj& exceptObj)
        {
            const bool strictly = strictObj.isMissing() ? false : (strictObj.get<double>(0.0) != 0.0);
            const bool exceptMe = exceptObj.isMissing() ? true : (exceptObj.get<double>(1.0) != 0.0);

            // 호출 셀의 "[통합문서]시트" 이름.
            auto [caller, rc1] = tryCallExcel(kXlfCaller);
            if (rc1 != 0) return returnValue(CellError::Value);
            auto [fullName, rc2] = tryCallExcel(kXlSheetNm, caller);
            if (rc2 != 0) return returnValue(CellError::Value);
            const std::wstring full = fullName.toString();
            const size_t rb = full.find(L']');
            if (full.empty() || full[0] != L'[' || rb == std::wstring::npos)
                return returnValue(CellError::Value);
            const std::wstring wbName = full.substr(1, rb - 1);
            const std::wstring mySheet = full.substr(rb + 1);

            IDispatch* app = nullptr;
            try
            {
                xloil::COM::connectCom();
                IUnknown* unk = xloil::thisApp().ptr();
                if (!unk || FAILED(unk->QueryInterface(IID_IDispatch, (void**)&app)) || !app)
                    return returnValue(CellError::Value);
            }
            catch (...) { return returnValue(CellError::Value); }

            std::vector<std::wstring> names;   // 보이는 시트만
            {
                VARIANT vWbs; VariantInit(&vWbs);
                VARIANT aName; VariantInit(&aName);
                aName.vt = VT_BSTR; aName.bstrVal = SysAllocString(wbName.c_str());
                if (SUCCEEDED(getDispProp(app, L"Workbooks", &aName, 1, &vWbs)) &&
                    vWbs.vt == VT_DISPATCH && vWbs.pdispVal)
                {
                    IDispatch* wb = vWbs.pdispVal;
                    VARIANT vWss; VariantInit(&vWss);
                    if (SUCCEEDED(getDispProp(wb, L"Worksheets", nullptr, 0, &vWss)) &&
                        vWss.vt == VT_DISPATCH && vWss.pdispVal)
                    {
                        IDispatch* wss = vWss.pdispVal;
                        VARIANT vCount; VariantInit(&vCount);
                        long count = 0;
                        if (SUCCEEDED(getDispProp(wss, L"Count", nullptr, 0, &vCount)))
                        {
                            VariantChangeType(&vCount, &vCount, 0, VT_I4);
                            count = vCount.lVal;
                        }
                        for (long i = 1; i <= count; ++i)
                        {
                            VARIANT idx; VariantInit(&idx); idx.vt = VT_I4; idx.lVal = i;
                            VARIANT vSh; VariantInit(&vSh);
                            if (FAILED(getDispProp(wss, L"Item", &idx, 1, &vSh)) ||
                                vSh.vt != VT_DISPATCH || !vSh.pdispVal) continue;
                            IDispatch* sh = vSh.pdispVal;
                            VARIANT vVis; VariantInit(&vVis);
                            long vis = -1;
                            if (SUCCEEDED(getDispProp(sh, L"Visible", nullptr, 0, &vVis)))
                            {
                                VariantChangeType(&vVis, &vVis, 0, VT_I4);
                                vis = vVis.lVal;
                            }
                            if (vis == -1)   // xlSheetVisible
                            {
                                VARIANT vName; VariantInit(&vName);
                                if (SUCCEEDED(getDispProp(sh, L"Name", nullptr, 0, &vName)) &&
                                    vName.vt == VT_BSTR && vName.bstrVal)
                                    names.emplace_back(vName.bstrVal);
                                VariantClear(&vName);
                            }
                            sh->Release();
                        }
                        wss->Release();
                    }
                    wb->Release();
                }
                VariantClear(&aName);
            }
            app->Release();

            if (names.empty()) return returnValue(CellError::Value);
            const int n = (int)names.size();

            // 시작/끝: 번호 또는 이름.
            int iStart = 1, iEnd = n;
            auto resolve = [&](const ExcelObj& v, int def, int& out) -> bool
            {
                if (v.isMissing()) { out = def; return true; }
                if (v.type() == ExcelType::Num || v.type() == ExcelType::Int)
                {
                    const int k = v.get<int>(0);
                    out = k == 0 ? def : k;
                    return true;
                }
                const std::wstring s = v.toString();
                for (int i = 0; i < n; ++i)
                    if (names[(size_t)i] == s) { out = i + 1; return true; }
                return false;
            };
            if (!resolve(startObj, 1, iStart) || !resolve(endObj, n, iEnd))
                return returnValue(CellError::Value);
            if (iStart < 1 || iEnd > n) return returnValue(CellError::Value);
            const int from = iStart + (strictly ? 1 : 0);
            const int to = iEnd - (strictly ? 1 : 0);
            if (from > to) return returnValue(CellError::Value);

            std::vector<ExcelObj> vals;
            for (int i = from; i <= to; ++i)
            {
                if (exceptMe && names[(size_t)i - 1] == mySheet) continue;
                vals.push_back(ExcelObj(std::wstring_view(names[(size_t)i - 1])));
            }
            if (vals.empty()) return returnValue(CellError::NA);
            return egtools::core::output(egtools::core::makeArray(
                (ExcelArrayBuilder::row_t)vals.size(), 1, vals));
        }

        // ---- TOTALPAGES (매크로형, XLM) ------------------------------------

        ExcelObj* totalPages(const ExcelObj& ignoreObj)
        {
            const bool ignoreStart = ignoreObj.isMissing() ? false
                                                           : (ignoreObj.get<double>(0.0) != 0.0);
            auto [startNum, rc1] = tryCallExcel(kXlfGetDocument, 49);
            double startVal = 1;
            if (rc1 == 0 && (startNum.type() == ExcelType::Num || startNum.type() == ExcelType::Int))
                startVal = startNum.get<double>(1);

            auto [pages, rc2] = tryCallExcel(kXlfGetDocument, 50);
            if (rc2 != 0 || (pages.type() != ExcelType::Num && pages.type() != ExcelType::Int))
                return returnValue(CellError::Value);

            const double offset = ignoreStart ? 0 : startVal - 1;
            return returnValue(ExcelObj(pages.get<double>(0) + offset));
        }
    }

    void registerEgUtil()
    {
        egtools::core::registerFn(L"DIRFOLDER",
            [](const ExcelObj& p, const ExcelObj& f, const ExcelObj& d,
               const ExcelObj& s, const ExcelObj& h) -> ExcelObj*
            { return dirFolder(p, f, d, s, h); });

        // 스칼라 네트워크 함수 — legacy 호스트에서도 MTR 병렬 허용.
        egtools::core::registerFn(L"RESTAPI",
            [](const ExcelObj& u, const ExcelObj& m, const ExcelObj& h,
               const ExcelObj& b) -> ExcelObj*
            { return restApi(u, m, h, b); },
            /*macro*/ false, /*threadsafe*/ true);

        egtools::core::registerFn(L"EXRATE",
            [](const ExcelObj& c, const ExcelObj& d) -> ExcelObj*
            { return exRate(c, d); },
            /*macro*/ false, /*threadsafe*/ true);

        egtools::core::registerFn(L"EVAL",
            [](const ExcelObj& f) -> ExcelObj*
            {
                try { return evalFn(f); }
                catch (...) { return returnValue(CellError::Value); }
            }, /*macro*/ true);

        egtools::core::registerFn(L"SHEETLIST",
            [](const ExcelObj& s, const ExcelObj& e, const ExcelObj& st,
               const ExcelObj& ex) -> ExcelObj*
            {
                try { return sheetList(s, e, st, ex); }
                catch (...) { return returnValue(CellError::Value); }
            }, /*macro*/ true);

        egtools::core::registerFn(L"TOTALPAGES",
            [](const ExcelObj& i) -> ExcelObj*
            {
                try { return totalPages(i); }
                catch (...) { return returnValue(CellError::Value); }
            }, /*macro*/ true);
    }
}
