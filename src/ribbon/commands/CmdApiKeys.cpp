// CmdApiKeys.cpp — 공공 API 키 관리 다이얼로그 (plan/23 §3-5, 사용자 제안).
//
// 서비스별 키를 HKCU\Software\EGTools\ApiKeys에 등록/삭제한다(FxPublicApi.cpp의
// 키 저장소와 동일 — 함수 마지막 인수로 키를 넘기는 기존 방식과 공존).
// 키는 문서/수식에 남지 않고, 표시할 때는 마스킹한다(등록 여부만 노출).

#include "Commands.h"
#include "CmdCommon.h"

#include <vector>

namespace egtools::commands
{
    namespace
    {
        constexpr const wchar_t* kKeyPath = L"Software\\EGTools\\ApiKeys";

        // 서비스 레지스트리 값 이름(FxPublicApi.cpp와 일치) + i18n 설명 키.
        struct Service { const wchar_t* value; const wchar_t* i18nKey; };
        constexpr Service kServices[] = {
            { L"juso",    L"cmd.apikeys.svc.juso" },
            { L"odcloud", L"cmd.apikeys.svc.odcloud" },
            { L"vworld",  L"cmd.apikeys.svc.vworld" },
            { L"datago",  L"cmd.apikeys.svc.datago" },
        };

        bool regHasKey(const wchar_t* service)
        {
            HKEY h = nullptr;
            if (RegOpenKeyExW(HKEY_CURRENT_USER, kKeyPath, 0, KEY_READ, &h) != ERROR_SUCCESS)
                return false;
            DWORD type = 0, size = 0;
            const bool has = RegQueryValueExW(h, service, nullptr, &type, nullptr, &size)
                                 == ERROR_SUCCESS && type == REG_SZ && size > sizeof(wchar_t);
            RegCloseKey(h);
            return has;
        }

        void regSetKey(const wchar_t* service, const std::wstring& key)
        {
            HKEY h = nullptr;
            if (RegCreateKeyExW(HKEY_CURRENT_USER, kKeyPath, 0, nullptr, 0, KEY_WRITE,
                                nullptr, &h, nullptr) != ERROR_SUCCESS)
                return;
            RegSetValueExW(h, service, 0, REG_SZ, (const BYTE*)key.c_str(),
                           (DWORD)((key.size() + 1) * sizeof(wchar_t)));
            RegCloseKey(h);
        }

        void regDeleteKey(const wchar_t* service)
        {
            HKEY h = nullptr;
            if (RegOpenKeyExW(HKEY_CURRENT_USER, kKeyPath, 0, KEY_WRITE, &h) != ERROR_SUCCESS)
                return;
            RegDeleteValueW(h, service);
            RegCloseKey(h);
        }

        // ── 다이얼로그 ──
        constexpr WORD IDC_COMBO  = 1101;
        constexpr WORD IDC_STATUS = 1102;
        constexpr WORD IDC_KEY    = 1103;
        constexpr WORD IDC_SAVE   = 1104;
        constexpr WORD IDC_DELETE = 1105;

        void refreshStatus(HWND dlg)
        {
            using egtools::i18n::t;
            const int i = (int)SendDlgItemMessageW(dlg, IDC_COMBO, CB_GETCURSEL, 0, 0);
            if (i < 0 || i >= (int)_countof(kServices)) return;
            const bool has = regHasKey(kServices[i].value);
            SetDlgItemTextW(dlg, IDC_STATUS,
                            (has ? L"●●●●●●  " + t(L"cmd.apikeys.registered")
                                 : t(L"cmd.apikeys.empty")).c_str());
            EnableWindow(GetDlgItem(dlg, IDC_DELETE), has);
            SetDlgItemTextW(dlg, IDC_KEY, L"");
        }

        INT_PTR CALLBACK apiKeysProc(HWND dlg, UINT msg, WPARAM wp, LPARAM)
        {
            using egtools::i18n::t;
            switch (msg)
            {
            case WM_INITDIALOG:
            {
                for (const auto& s : kServices)
                    SendDlgItemMessageW(dlg, IDC_COMBO, CB_ADDSTRING, 0,
                                        (LPARAM)t(s.i18nKey).c_str());
                SendDlgItemMessageW(dlg, IDC_COMBO, CB_SETCURSEL, 0, 0);
                refreshStatus(dlg);

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
                case IDC_COMBO:
                    if (HIWORD(wp) == CBN_SELCHANGE) refreshStatus(dlg);
                    return TRUE;
                case IDC_SAVE:
                {
                    const int i = (int)SendDlgItemMessageW(dlg, IDC_COMBO, CB_GETCURSEL, 0, 0);
                    if (i < 0 || i >= (int)_countof(kServices)) return TRUE;
                    wchar_t buf[512]{};
                    GetDlgItemTextW(dlg, IDC_KEY, buf, 512);
                    std::wstring key = buf;
                    // 앞뒤 공백 제거
                    const auto b = key.find_first_not_of(L" \t");
                    const auto e = key.find_last_not_of(L" \t");
                    key = (b == std::wstring::npos) ? L"" : key.substr(b, e - b + 1);
                    if (!key.empty())
                    {
                        regSetKey(kServices[i].value, key);
                        refreshStatus(dlg);
                    }
                    return TRUE;
                }
                case IDC_DELETE:
                {
                    const int i = (int)SendDlgItemMessageW(dlg, IDC_COMBO, CB_GETCURSEL, 0, 0);
                    if (i >= 0 && i < (int)_countof(kServices))
                    {
                        regDeleteKey(kServices[i].value);
                        refreshStatus(dlg);
                    }
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

        // Dialogs.cpp와 동일한 in-memory DLGTEMPLATE 조립(콤보 박스 포함이라 별도).
        std::vector<BYTE> buildTemplate()
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
            h.cx = 262; h.cy = 106;
            write(&h, sizeof(h));
            writeWord(0); writeWord(0);
            writeStr(t(L"cmd.apikeys.title"));
            writeWord(9); writeStr(L"MS Shell Dlg");

            WORD count = 0;
            auto add = [&](WORD atom, WORD id, DWORD style,
                           int x, int y, int w, int hh, const std::wstring& text) {
                item(atom, id, style, x, y, w, hh, text); ++count;
            };
            add(0x0082, 0xFFFF, SS_LEFT, 7, 7, 60, 9, t(L"cmd.apikeys.service"));
            add(0x0085, IDC_COMBO, CBS_DROPDOWNLIST | WS_VSCROLL | WS_TABSTOP,
                7, 18, 248, 80, L"");
            add(0x0082, IDC_STATUS, SS_LEFT, 7, 37, 248, 10, L"");
            add(0x0082, 0xFFFE, SS_LEFT, 7, 52, 60, 9, t(L"cmd.apikeys.key"));
            add(0x0081, IDC_KEY, WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL,
                7, 63, 248, 13, L"");
            add(0x0080, IDC_SAVE, BS_DEFPUSHBUTTON | WS_TABSTOP, 90, 85, 50, 14,
                t(L"cmd.apikeys.save"));
            add(0x0080, IDC_DELETE, BS_PUSHBUTTON | WS_TABSTOP, 145, 85, 50, 14,
                t(L"cmd.apikeys.delete"));
            add(0x0080, IDCANCEL, BS_PUSHBUTTON | WS_TABSTOP, 200, 85, 55, 14,
                t(L"cmd.apikeys.close"));
            ((DLGTEMPLATE*)buf.data())->cdit = count;
            return buf;
        }
    }

    void manageApiKeys()
    {
        try
        {
            std::vector<BYTE> tmpl = buildTemplate();
            DialogBoxIndirectParamW(GetModuleHandleW(nullptr),
                                    (const DLGTEMPLATE*)tmpl.data(),
                                    excelHwnd(), apiKeysProc, 0);
        }
        catch (...) {}
    }
}
