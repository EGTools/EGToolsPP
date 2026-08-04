// Dialogs.cpp — 리본 명령 공용 Win32 다이얼로그.
//
// 입력 다이얼로그는 리소스(.rc) 없이 메모리 내 DLGTEMPLATE로 조립한다 —
// 단일 .xll 원칙(리소스는 RCDATA만)과 다국어 라벨(i18n 런타임 문자열)을
// 동시에 만족하는 가장 단순한 방법. 버튼 문구는 i18n "dlg.ok"/"dlg.cancel".

#include "Dialogs.h"
#include "../../core/I18n.h"

#include <algorithm>
#include <commdlg.h>
#include <shobjidl.h>
#include <vector>

#pragma comment(lib, "comdlg32.lib")

namespace egtools::dialogs
{
    namespace
    {
        // ── in-memory DLGTEMPLATE builder ───────────────────────────────────
        // 모든 데이터는 WORD 정렬. 클래스는 표준 아톰(0x0080=button,
        // 0x0081=edit, 0x0082=static)만 사용.
        class DlgBuilder
        {
        public:
            DlgBuilder(const std::wstring& title, int w, int h)
            {
                alignDword();
                DLGTEMPLATE h1{};
                h1.style = DS_MODALFRAME | DS_SETFONT | WS_POPUP | WS_CAPTION | WS_SYSMENU;
                h1.cx = (short)w; h1.cy = (short)h;
                write(&h1, sizeof(h1));
                writeWord(0);               // no menu
                writeWord(0);               // default dialog class
                writeStr(title);
                writeWord(9);               // point size
                writeStr(L"MS Shell Dlg");  // DS_SETFONT font
            }

            void item(WORD classAtom, WORD id, DWORD style, int x, int y, int w, int h,
                      const std::wstring& text = L"")
            {
                ++_count;
                alignDword();
                DLGITEMTEMPLATE it{};
                it.style = style | WS_CHILD | WS_VISIBLE;
                it.x = (short)x; it.y = (short)y; it.cx = (short)w; it.cy = (short)h;
                it.id = id;
                write(&it, sizeof(it));
                writeWord(0xFFFF); writeWord(classAtom);
                writeStr(text);
                writeWord(0);               // no creation data
            }

            // cdit를 채워 완성된 템플릿 반환(수명은 builder와 동일).
            const DLGTEMPLATE* finish()
            {
                ((DLGTEMPLATE*)_buf.data())->cdit = _count;
                return (const DLGTEMPLATE*)_buf.data();
            }

        private:
            std::vector<BYTE> _buf;
            WORD _count = 0;

            void write(const void* p, size_t n)
            {
                const BYTE* b = (const BYTE*)p;
                _buf.insert(_buf.end(), b, b + n);
            }
            void writeWord(WORD w) { write(&w, 2); }
            void writeStr(const std::wstring& s)
            {
                write(s.c_str(), (s.size() + 1) * sizeof(wchar_t));
            }
            void alignDword() { while (_buf.size() % 4) _buf.push_back(0); }
        };

        void centerOnOwner(HWND dlg, HWND owner)
        {
            RECT ro{}, rd{};
            if (!owner || !GetWindowRect(owner, &ro)) SystemParametersInfoW(SPI_GETWORKAREA, 0, &ro, 0);
            GetWindowRect(dlg, &rd);
            const int w = rd.right - rd.left, h = rd.bottom - rd.top;
            int x = ro.left + ((ro.right - ro.left) - w) / 2;
            int y = ro.top + ((ro.bottom - ro.top) - h) / 2;
            SetWindowPos(dlg, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        }

        struct InputCtx
        {
            const std::wstring* label;
            std::wstring* value;
        };

        constexpr WORD IDC_LABEL = 1001;
        constexpr WORD IDC_EDIT  = 1002;

        INT_PTR CALLBACK inputProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
        {
            switch (msg)
            {
            case WM_INITDIALOG:
            {
                SetWindowLongPtrW(dlg, DWLP_USER, lp);
                auto* ctx = (InputCtx*)lp;
                SetDlgItemTextW(dlg, IDC_LABEL, ctx->label->c_str());
                SetDlgItemTextW(dlg, IDC_EDIT, ctx->value->c_str());
                SendDlgItemMessageW(dlg, IDC_EDIT, EM_SETSEL, 0, -1);
                centerOnOwner(dlg, GetParent(dlg));
                SetFocus(GetDlgItem(dlg, IDC_EDIT));
                return FALSE;   // 포커스 직접 지정
            }
            case WM_COMMAND:
                switch (LOWORD(wp))
                {
                case IDOK:
                {
                    auto* ctx = (InputCtx*)GetWindowLongPtrW(dlg, DWLP_USER);
                    wchar_t buf[1024]{};
                    GetDlgItemTextW(dlg, IDC_EDIT, buf, 1024);
                    *ctx->value = buf;
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
    }

    bool inputText(HWND owner, const std::wstring& title,
                   const std::wstring& label, std::wstring& value, bool password)
    {
        using egtools::i18n::t;
        // 라벨 높이는 줄 수에 맞춰 동적 산정(메일머지 옵션 안내처럼 여러 줄
        // 프롬프트가 잘리지 않도록 — 한 줄 ≈ 9 DLU).
        const int lines = 1 + (int)std::count(label.begin(), label.end(), L'\n');
        const int labelH = std::max(18, lines * 9);
        const int editY = 7 + labelH + 3;
        const int btnY = editY + 13 + 6;
        DlgBuilder b(title, 230, btnY + 14 + 7);
        b.item(0x0082, IDC_LABEL, SS_LEFT, 7, 7, 216, labelH);
        b.item(0x0081, IDC_EDIT,
               WS_BORDER | WS_TABSTOP | ES_AUTOHSCROLL | (password ? ES_PASSWORD : 0),
               7, editY, 216, 13);
        b.item(0x0080, IDOK, BS_DEFPUSHBUTTON | WS_TABSTOP, 118, btnY, 50, 14, t(L"dlg.ok"));
        b.item(0x0080, IDCANCEL, BS_PUSHBUTTON | WS_TABSTOP, 173, btnY, 50, 14, t(L"dlg.cancel"));

        InputCtx ctx{ &label, &value };
        return DialogBoxIndirectParamW(GetModuleHandleW(nullptr), b.finish(), owner,
                                       inputProc, (LPARAM)&ctx) == IDOK;
    }

    bool inputNumber(HWND owner, const std::wstring& title,
                     const std::wstring& label, double& value,
                     double minV, double maxV, double defV)
    {
        wchar_t def[32];
        swprintf_s(def, L"%g", defV);
        std::wstring text = def;
        if (!inputText(owner, title, label, text)) return false;

        wchar_t* end = nullptr;
        const double v = wcstod(text.c_str(), &end);
        // 전체가 숫자가 아니면 종료(VB IsNumeric 규약).
        while (end && *end == L' ') ++end;
        if (!end || *end != L'\0' || end == text.c_str()) return false;
        value = v < minV ? minV : (v > maxV ? maxV : v);
        return true;
    }

    bool pickColor(HWND owner, COLORREF& color)
    {
        static COLORREF custom[16] = {};
        CHOOSECOLORW cc{};
        cc.lStructSize = sizeof(cc);
        cc.hwndOwner = owner;
        cc.lpCustColors = custom;
        cc.rgbResult = color;
        cc.Flags = CC_FULLOPEN | CC_RGBINIT;
        if (!ChooseColorW(&cc)) return false;
        color = cc.rgbResult;
        return true;
    }

    bool pickImageFile(HWND owner, std::wstring& path)
    {
        // VB InsertPicture와 동일한 확장자 화이트리스트.
        static const wchar_t* filter =
            L"Images\0*.jpg;*.jpeg;*.gif;*.png;*.bmp;*.tif;*.tiff;*.emf;*.wmf;"
            L"*.jfif;*.jpe;*.dib;*.rle;*.emz;*.wmz;*.pcz;*.svg;*.pct;*.pict;*.wpg\0"
            L"All\0*.*\0";
        wchar_t buf[MAX_PATH]{};
        OPENFILENAMEW ofn{};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = owner;
        ofn.lpstrFilter = filter;
        ofn.lpstrFile = buf;
        ofn.nMaxFile = MAX_PATH;
        ofn.Flags = OFN_FILEMUSTEXIST | OFN_HIDEREADONLY | OFN_NOCHANGEDIR;
        if (!GetOpenFileNameW(&ofn)) return false;
        path = buf;
        return true;
    }

    bool pickFolder(HWND owner, std::wstring& path)
    {
        IFileOpenDialog* dlg = nullptr;
        if (FAILED(CoCreateInstance(CLSID_FileOpenDialog, nullptr,
                                    CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dlg))))
            return false;
        bool ok = false;
        DWORD opts = 0;
        dlg->GetOptions(&opts);
        dlg->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
        if (SUCCEEDED(dlg->Show(owner)))
        {
            IShellItem* item = nullptr;
            if (SUCCEEDED(dlg->GetResult(&item)) && item)
            {
                PWSTR p = nullptr;
                if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &p)) && p)
                {
                    path = p;
                    CoTaskMemFree(p);
                    ok = true;
                }
                item->Release();
            }
        }
        dlg->Release();
        return ok;
    }
}
