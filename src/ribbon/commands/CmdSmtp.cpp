// CmdSmtp.cpp — SMTP 설정 다이얼로그 + CDO 메일 발송 (plan/23 §3-3 v2).
//
// - 설정(발송인 이름/메일주소/서버/포트/SSL)은 HKCU\Software\EGTools\Smtp에
//   저장한다. **암호는 저장하지 않는다**(사용자 지시 — 발송 시마다 입력).
// - 발송은 CDO.Message late-bound COM(무의존, LateCom 패턴). VB의 .NET
//   SmtpClient 경로는 C++에서 불가하므로 VB의 CDO 폴백(SendMailWithCDO)을
//   본 구현으로 승격. CDO의 smtpusessl은 암시적 SSL(implicit TLS)만 지원하므로
//   기본 포트는 465다(587 STARTTLS 미지원 — 다이얼로그 안내문·manual에 명시).
// - 한글 제목/본문: BodyPart.Charset = "utf-8" 지정으로 인코딩 보장.

#include "Commands.h"
#include "CmdCommon.h"
#include "Smtp.h"

#include <vector>

namespace egtools::commands
{
    namespace
    {
        constexpr const wchar_t* kSmtpPath = L"Software\\EGTools\\Smtp";

        std::wstring regGet(const wchar_t* value, const wchar_t* defVal)
        {
            HKEY h = nullptr;
            if (RegOpenKeyExW(HKEY_CURRENT_USER, kSmtpPath, 0, KEY_READ, &h)
                    != ERROR_SUCCESS)
                return defVal;
            wchar_t buf[512]{};
            DWORD type = 0, size = sizeof(buf) - sizeof(wchar_t);
            const bool ok = RegQueryValueExW(h, value, nullptr, &type,
                                             (BYTE*)buf, &size) == ERROR_SUCCESS
                            && type == REG_SZ;
            RegCloseKey(h);
            return ok ? buf : defVal;
        }

        void regSet(const wchar_t* value, const std::wstring& s)
        {
            HKEY h = nullptr;
            if (RegCreateKeyExW(HKEY_CURRENT_USER, kSmtpPath, 0, nullptr, 0,
                                KEY_WRITE, nullptr, &h, nullptr) != ERROR_SUCCESS)
                return;
            RegSetValueExW(h, value, 0, REG_SZ, (const BYTE*)s.c_str(),
                           (DWORD)((s.size() + 1) * sizeof(wchar_t)));
            RegCloseKey(h);
        }

        // ── 설정 다이얼로그 (ApiKeys와 동일한 in-memory DLGTEMPLATE 패턴) ──
        constexpr WORD IDC_NAME   = 1201;
        constexpr WORD IDC_EMAIL  = 1202;
        constexpr WORD IDC_SERVER = 1203;
        constexpr WORD IDC_PORT   = 1204;
        constexpr WORD IDC_SSL    = 1205;

        INT_PTR CALLBACK smtpProc(HWND dlg, UINT msg, WPARAM wp, LPARAM)
        {
            switch (msg)
            {
            case WM_INITDIALOG:
            {
                const auto s = smtp::load();
                SetDlgItemTextW(dlg, IDC_NAME,   s.name.c_str());
                SetDlgItemTextW(dlg, IDC_EMAIL,  s.email.c_str());
                SetDlgItemTextW(dlg, IDC_SERVER, s.server.c_str());
                SetDlgItemTextW(dlg, IDC_PORT,   s.port.c_str());
                CheckDlgButton(dlg, IDC_SSL, s.ssl ? BST_CHECKED : BST_UNCHECKED);

                RECT ro{}, rd{};
                HWND owner = GetParent(dlg);
                if (owner && GetWindowRect(owner, &ro))
                {
                    GetWindowRect(dlg, &rd);
                    SetWindowPos(dlg, nullptr,
                                 ro.left + ((ro.right - ro.left) - (rd.right - rd.left)) / 2,
                                 ro.top + ((ro.bottom - ro.top) - (rd.bottom - rd.top)) / 2,
                                 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                }
                return TRUE;
            }
            case WM_COMMAND:
                switch (LOWORD(wp))
                {
                case IDOK:
                {
                    auto text = [&](WORD id) {
                        wchar_t buf[512]{};
                        GetDlgItemTextW(dlg, id, buf, 512);
                        std::wstring s = buf;
                        const auto b = s.find_first_not_of(L" \t");
                        const auto e = s.find_last_not_of(L" \t");
                        return b == std::wstring::npos ? std::wstring()
                                                       : s.substr(b, e - b + 1);
                    };
                    regSet(L"UserName",  text(IDC_NAME));
                    regSet(L"UserEmail", text(IDC_EMAIL));
                    regSet(L"Server",    text(IDC_SERVER));
                    regSet(L"Port",      text(IDC_PORT));
                    regSet(L"UseSsl",
                           IsDlgButtonChecked(dlg, IDC_SSL) == BST_CHECKED
                               ? L"1" : L"0");
                    EndDialog(dlg, IDOK);
                    return TRUE;
                }
                case IDCANCEL:
                    EndDialog(dlg, IDCANCEL);
                    return TRUE;
                }
                break;
            }
            return FALSE;
        }

        std::vector<BYTE> buildSmtpTemplate()
        {
            using egtools::i18n::t;
            std::vector<BYTE> buf;
            auto write = [&](const void* p, size_t n) {
                const BYTE* b = (const BYTE*)p;
                buf.insert(buf.end(), b, b + n);
            };
            auto writeWord = [&](WORD w) { write(&w, 2); };
            auto writeStr = [&](const std::wstring& s) {
                write(s.c_str(), (s.size() + 1) * sizeof(wchar_t));
            };
            auto align = [&] { while (buf.size() % 4) buf.push_back(0); };
            auto item = [&](WORD atom, WORD id, DWORD style,
                            int x, int y, int w, int h, const std::wstring& text) {
                align();
                DLGITEMTEMPLATE it{};
                it.style = style | WS_CHILD | WS_VISIBLE;
                it.x = (short)x; it.y = (short)y; it.cx = (short)w; it.cy = (short)h;
                it.id = id;
                write(&it, sizeof(it));
                writeWord(0xFFFF); writeWord(atom);
                writeStr(text);
                writeWord(0);
            };

            DLGTEMPLATE h{};
            h.style = DS_MODALFRAME | DS_SETFONT | WS_POPUP | WS_CAPTION | WS_SYSMENU;
            h.cx = 262; h.cy = 148;
            write(&h, sizeof(h));
            writeWord(0); writeWord(0);
            writeStr(t(L"cmd.smtp.title"));
            writeWord(9); writeStr(L"MS Shell Dlg");

            WORD count = 0;
            auto add = [&](WORD atom, WORD id, DWORD style,
                           int x, int y, int w, int hh, const std::wstring& text) {
                item(atom, id, style, x, y, w, hh, text); ++count;
            };
            const DWORD edit = WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL;
            add(0x0082, 0xFFFF, SS_LEFT, 7, 9, 80, 9, t(L"cmd.smtp.name"));
            add(0x0081, IDC_NAME, edit, 92, 7, 163, 13, L"");
            add(0x0082, 0xFFFE, SS_LEFT, 7, 27, 80, 9, t(L"cmd.smtp.email"));
            add(0x0081, IDC_EMAIL, edit, 92, 25, 163, 13, L"");
            add(0x0082, 0xFFFD, SS_LEFT, 7, 45, 80, 9, t(L"cmd.smtp.server"));
            add(0x0081, IDC_SERVER, edit, 92, 43, 163, 13, L"");
            add(0x0082, 0xFFFC, SS_LEFT, 7, 63, 80, 9, t(L"cmd.smtp.port"));
            add(0x0081, IDC_PORT, edit | ES_NUMBER, 92, 61, 50, 13, L"");
            add(0x0080, IDC_SSL, BS_AUTOCHECKBOX | WS_TABSTOP, 155, 61, 100, 12,
                L"SSL");
            // 안내: 암호 미저장 + 암시적 SSL(465) 권장
            add(0x0082, 0xFFFB, SS_LEFT, 7, 82, 248, 18, t(L"cmd.smtp.note"));
            add(0x0082, 0xFFFA, SS_LEFT, 7, 102, 248, 18, t(L"cmd.smtp.sslNote"));
            add(0x0080, IDOK, BS_DEFPUSHBUTTON | WS_TABSTOP, 145, 127, 50, 14,
                t(L"cmd.smtp.save"));
            add(0x0080, IDCANCEL, BS_PUSHBUTTON | WS_TABSTOP, 200, 127, 55, 14,
                t(L"dlg.cancel"));
            ((DLGTEMPLATE*)buf.data())->cdit = count;
            return buf;
        }

        // ── CDO late-bound 헬퍼 ──
        IDispatch* createDispatch(const wchar_t* progId)
        {
            CLSID clsid{};
            if (FAILED(CLSIDFromProgID(progId, &clsid))) return nullptr;
            IDispatch* d = nullptr;
            if (FAILED(CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER,
                                        IID_IDispatch, (void**)&d)))
                return nullptr;
            return d;
        }

        // invokeRaw + EXCEPINFO 캡처(발송 실패 시 CDO의 설명 텍스트를 얻는다).
        bool invokeEx(IDispatch* d, const wchar_t* name, WORD flags,
                      VARIANT* result, VARIANT* args, UINT nArgs,
                      std::wstring* err)
        {
            const DISPID id = dispidOf(d, name);
            if (id == DISPID_UNKNOWN)
            {
                if (err) *err = std::wstring(name) + L": not found";
                return false;
            }
            std::vector<VARIANT> rev(nArgs);
            for (UINT i = 0; i < nArgs; ++i) rev[i] = args[nArgs - 1 - i];
            DISPPARAMS dp{ nArgs ? rev.data() : nullptr, nullptr, nArgs, 0 };
            DISPID putId = DISPID_PROPERTYPUT;
            if (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF))
            { dp.rgdispidNamedArgs = &putId; dp.cNamedArgs = 1; }
            EXCEPINFO ex{};
            const HRESULT hr = d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                                         &dp, result, &ex, nullptr);
            if (SUCCEEDED(hr)) return true;
            if (err)
            {
                if (hr == DISP_E_EXCEPTION && ex.bstrDescription)
                    *err = ex.bstrDescription;
                else
                {
                    wchar_t buf[64];
                    swprintf_s(buf, L"HRESULT 0x%08X", (unsigned)hr);
                    *err = buf;
                }
            }
            if (ex.bstrSource) SysFreeString(ex.bstrSource);
            if (ex.bstrDescription) SysFreeString(ex.bstrDescription);
            if (ex.bstrHelpFile) SysFreeString(ex.bstrHelpFile);
            return false;
        }

        // 객체 속성 대입(VB의 Set obj.Prop = x — PROPERTYPUTREF 우선).
        bool putRefObj(IDispatch* d, const wchar_t* name, IDispatch* value)
        {
            VARIANT v; VariantInit(&v);
            v.vt = VT_DISPATCH; v.pdispVal = value;
            if (invokeEx(d, name, DISPATCH_PROPERTYPUTREF, nullptr, &v, 1, nullptr))
                return true;
            return invokeEx(d, name, DISPATCH_PROPERTYPUT, nullptr, &v, 1, nullptr);
        }

        // Configuration.Fields.Item(url).Value = v — v의 소유권을 가져와 해제한다.
        bool setCdoField(IDispatch* fields, const wchar_t* url, VARIANT v)
        {
            VARIANT a = varBStr(url);
            IDispatch* field = getObject(fields, L"Item", &a, 1);
            VariantClear(&a);
            if (!field) { VariantClear(&v); return false; }
            Releaser rf{ field };
            const bool ok = putVar(field, L"Value", v);
            VariantClear(&v);
            return ok;
        }
    }

    namespace smtp
    {
        Settings load()
        {
            Settings s;
            s.name   = regGet(L"UserName", L"");
            s.email  = regGet(L"UserEmail", L"");
            s.server = regGet(L"Server", L"");
            s.port   = regGet(L"Port", L"465");
            s.ssl    = regGet(L"UseSsl", L"1") != L"0";
            return s;
        }

        bool complete(const Settings& s)
        {
            return !s.email.empty() && !s.server.empty() && !s.port.empty();
        }

        bool ensure(Settings& s)
        {
            s = load();
            if (complete(s)) return true;
            smtpSettings();             // 설정 다이얼로그를 띄워 채우게 한다
            s = load();
            return complete(s);
        }

        std::wstring send(const Settings& s, const std::wstring& password,
                          const std::wstring& to, const std::wstring& cc,
                          const std::wstring& bcc, const std::wstring& subject,
                          const std::wstring& body, const std::wstring& attachment)
        {
            constexpr const wchar_t* kCfg =
                L"http://schemas.microsoft.com/cdo/configuration/";

            IDispatch* msg = createDispatch(L"CDO.Message");
            if (!msg) return L"CDO.Message unavailable";
            Releaser rMsg{ msg };
            IDispatch* cfg = createDispatch(L"CDO.Configuration");
            if (!cfg) return L"CDO.Configuration unavailable";
            Releaser rCfg{ cfg };

            {
                IDispatch* fields = getObject(cfg, L"Fields");
                if (!fields) return L"CDO fields unavailable";
                Releaser rF{ fields };
                auto url = [&](const wchar_t* leaf) {
                    return std::wstring(kCfg) + leaf;
                };
                setCdoField(fields, url(L"sendusing").c_str(), varLong(2)); // 원격 SMTP
                setCdoField(fields, url(L"smtpserver").c_str(), varBStr(s.server));
                setCdoField(fields, url(L"smtpserverport").c_str(),
                            varLong(_wtoi(s.port.c_str())));
                setCdoField(fields, url(L"smtpauthenticate").c_str(), varLong(1));
                setCdoField(fields, url(L"sendusername").c_str(), varBStr(s.email));
                setCdoField(fields, url(L"sendpassword").c_str(), varBStr(password));
                VARIANT vb; VariantInit(&vb);
                vb.vt = VT_BOOL;
                vb.boolVal = s.ssl ? VARIANT_TRUE : VARIANT_FALSE;
                setCdoField(fields, url(L"smtpusessl").c_str(), vb);
                setCdoField(fields, url(L"smtpconnectiontimeout").c_str(),
                            varLong(60));
                std::wstring err;
                if (!invokeEx(fields, L"Update", DISPATCH_METHOD, nullptr,
                              nullptr, 0, &err))
                    return L"CDO config: " + err;
            }

            if (!putRefObj(msg, L"Configuration", cfg))
                return L"CDO configuration bind failed";

            const std::wstring from =
                s.name.empty() ? s.email
                               : L"\"" + s.name + L"\" <" + s.email + L">";
            putBStr(msg, L"From", from);
            putBStr(msg, L"To", to);
            if (!cc.empty())  putBStr(msg, L"CC", cc);
            if (!bcc.empty()) putBStr(msg, L"BCC", bcc);
            putBStr(msg, L"Subject", subject);
            putBStr(msg, L"TextBody", body);
            // 한글 등 비ASCII 제목/본문 인코딩 지정(TextBody 설정 후).
            if (IDispatch* bp = getObject(msg, L"BodyPart"))
            {
                Releaser rBp{ bp };
                putBStr(bp, L"Charset", L"utf-8");
            }

            if (!attachment.empty())
            {
                VARIANT a = varBStr(attachment);
                std::wstring err;
                VARIANT r; VariantInit(&r);
                const bool ok = invokeEx(msg, L"AddAttachment", DISPATCH_METHOD,
                                         &r, &a, 1, &err);
                VariantClear(&r);
                VariantClear(&a);
                if (!ok) return L"Attachment: " + err;
            }

            std::wstring err;
            if (!invokeEx(msg, L"Send", DISPATCH_METHOD, nullptr, nullptr, 0, &err))
                return err;
            return L"";
        }
    }

    void smtpSettings()
    {
        try
        {
            std::vector<BYTE> tmpl = buildSmtpTemplate();
            DialogBoxIndirectParamW(GetModuleHandleW(nullptr),
                                    (const DLGTEMPLATE*)tmpl.data(),
                                    excelHwnd(), smtpProc, 0);
        }
        catch (...) {}
    }
}
