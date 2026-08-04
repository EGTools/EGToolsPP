// CmdCommon.h — 리본 명령 구현 공용 헬퍼.
//
// 규약(Convert.cpp에서 확립):
//  - COM은 late-bound(IDispatch)만 사용한다. typed #import 프록시는 이 정적
//    XLL의 지연 COM 컨텍스트에서 크래시(FxImage.cpp 참조).
//  - Application 획득은 xlOil thisApp() → IDispatch. 명령은 리본 콜백(메인
//    스레드)에서만 불리므로 직접 호출해도 안전하다.
//  - Excel 상태 저장/복원은 xlOil PauseExcel RAII(= VB OnOffAll).
#pragma once

#include "../../core/LateCom.h"
#include "../../core/I18n.h"

#include <xlOil/AppObjects.h>
#include <xlOil/State.h>

#include <windows.h>
#include <string>

namespace egtools::commands
{
    using namespace egtools::latecom;

    inline HWND excelHwnd()
    {
        return (HWND)xloil::Environment::excelProcess().hWnd;
    }

    // AddRef된 Application IDispatch (caller releases — Releaser 사용).
    inline IDispatch* appDisp()
    {
        IDispatch* d = nullptr;
        auto& app = xloil::thisApp();
        if (app.ptr()) app.ptr()->QueryInterface(IID_IDispatch, (void**)&d);
        return d;
    }

    // Selection이 셀 Range일 때만 반환(그림 등 다른 선택은 nullptr).
    // Range 판별: Address 속성 보유 여부(ShapeRange/Chart 등에는 없음).
    inline IDispatch* selectionRange(IDispatch* app)
    {
        IDispatch* sel = getObject(app, L"Selection");
        if (!sel) return nullptr;
        if (dispidOf(sel, L"Address") == DISPID_UNKNOWN)
        {
            sel->Release();
            return nullptr;
        }
        return sel;
    }

    inline void statusBar(IDispatch* app, const std::wstring& text)
    {
        putBStr(app, L"StatusBar", text);
    }
    inline void statusBarClear(IDispatch* app)
    {
        putBool(app, L"StatusBar", false);   // = VB "xl.StatusBar = vbNullString"
    }

    // ── 메시지 상자(i18n) ──
    inline void msgInfo(const std::wstring& text)
    {
        MessageBoxW(excelHwnd(), text.c_str(), L"EGTools++", MB_OK | MB_ICONINFORMATION);
    }
    inline void msgWarn(const std::wstring& text)
    {
        MessageBoxW(excelHwnd(), text.c_str(), L"EGTools++", MB_OK | MB_ICONEXCLAMATION);
    }
    inline bool msgOkCancel(const std::wstring& text)
    {
        return MessageBoxW(excelHwnd(), text.c_str(), L"EGTools++",
                           MB_OKCANCEL | MB_ICONINFORMATION | MB_DEFBUTTON1) == IDOK;
    }

    // "라벨 : N" 형태(VB MsgBox(EGT(key) & " : " & n) 관례).
    inline std::wstring withCount(const std::wstring& label, long n)
    {
        return label + L" : " + std::to_wstring(n);
    }

    // ── VARIANT 속성 get/put ──
    inline bool getVar(IDispatch* d, const wchar_t* name, VARIANT* out)
    {
        return invokeRaw(d, name, DISPATCH_PROPERTYGET, out, nullptr, 0);
    }
    inline bool putVar(IDispatch* d, const wchar_t* name, VARIANT v)
    {
        return invokeRaw(d, name, DISPATCH_PROPERTYPUT, nullptr, &v, 1);
    }

    inline bool isEmptyVal(const VARIANT& v)
    {
        return v.vt == VT_EMPTY || v.vt == VT_NULL ||
               (v.vt == VT_BSTR && (!v.bstrVal || !*v.bstrVal));
    }

    inline std::wstring varToString(const VARIANT& v)
    {
        if (v.vt == VT_BSTR) return v.bstrVal ? v.bstrVal : L"";
        VARIANT out; VariantInit(&out);
        std::wstring s;
        if (SUCCEEDED(VariantChangeType(&out, const_cast<VARIANT*>(&v), 0, VT_BSTR)))
            s = out.bstrVal ? out.bstrVal : L"";
        VariantClear(&out);
        return s;
    }

    // Application.InputBox(Type:=8) — Range 선택 입력. 취소/실패 시 nullptr.
    // (반환 VARIANT가 Range(VT_DISPATCH) / False(VT_BOOL) 혼재 — VB 관례 처리)
    inline IDispatch* inputBoxRange(IDispatch* app, const std::wstring& prompt,
                                    const std::wstring& title,
                                    const std::wstring& defaultAddr)
    {
        VARIANT a[8];
        for (auto& x : a) { VariantInit(&x); x.vt = VT_ERROR; x.scode = DISP_E_PARAMNOTFOUND; }
        a[0] = varBStr(prompt);
        a[1] = varBStr(title);
        if (!defaultAddr.empty()) a[2] = varBStr(defaultAddr);
        a[7] = varLong(8);                       // Type:=8 (Range)
        VARIANT r; VariantInit(&r);
        const bool ok = invokeRaw(app, L"InputBox", DISPATCH_METHOD, &r, a, 8);
        VariantClear(&a[0]); VariantClear(&a[1]); VariantClear(&a[2]);
        if (!ok) { VariantClear(&r); return nullptr; }
        if (r.vt == VT_DISPATCH && r.pdispVal) return r.pdispVal;   // 소유권 이전
        VariantClear(&r);
        return nullptr;
    }

    // Range.Resize(rows, cols).
    inline IDispatch* rangeResize(IDispatch* range, long rows, long cols)
    {
        VARIANT a[2];
        VariantInit(&a[0]); VariantInit(&a[1]);
        a[0].vt = VT_I4; a[0].lVal = rows;
        a[1].vt = VT_I4; a[1].lVal = cols;
        return getObject(range, L"Resize", a, 2);
    }

    // Range.Cells(1,1) 등 2-인덱스 파라미터화 속성.
    inline IDispatch* getObjectIdx2(IDispatch* d, const wchar_t* name, long r, long c)
    {
        VARIANT a[2]; VariantInit(&a[0]); VariantInit(&a[1]);
        a[0].vt = VT_I4; a[0].lVal = r;
        a[1].vt = VT_I4; a[1].lVal = c;
        return getObject(d, name, a, 2);
    }

    // Range.Value2/Formula가 돌려주는 2차원 SAFEARRAY(VT_ARRAY|VT_VARIANT,
    // 1-base, 열 우선 저장) 접근자. 단일 셀은 스칼라로 오므로 loaded()가 false —
    // 호출자가 스칼라 경로를 따로 처리한다.
    struct SafeArr2D
    {
        VARIANT v{};
        SAFEARRAY* sa = nullptr;
        VARIANT* data = nullptr;
        long r1 = 0, r2 = -1, c1 = 0, c2 = -1;

        SafeArr2D() { VariantInit(&v); }
        ~SafeArr2D()
        {
            if (data) SafeArrayUnaccessData(sa);
            VariantClear(&v);
        }
        SafeArr2D(const SafeArr2D&) = delete;
        SafeArr2D& operator=(const SafeArr2D&) = delete;

        bool load(IDispatch* range, const wchar_t* prop)
        {
            if (!getVar(range, prop, &v)) return false;
            if (!(v.vt & VT_ARRAY)) return false;
            sa = (v.vt & VT_BYREF) ? *v.pparray : v.parray;
            if (!sa || SafeArrayGetDim(sa) != 2) { sa = nullptr; return false; }
            SafeArrayGetLBound(sa, 1, &r1); SafeArrayGetUBound(sa, 1, &r2);
            SafeArrayGetLBound(sa, 2, &c1); SafeArrayGetUBound(sa, 2, &c2);
            if (FAILED(SafeArrayAccessData(sa, (void**)&data)))
            { data = nullptr; return false; }
            return true;
        }
        bool loaded() const { return data != nullptr; }
        const VARIANT& at(long r, long c) const   // r/c는 배열 자체 인덱스(보통 1-base)
        {
            const long rows = r2 - r1 + 1;
            return data[(size_t)(c - c1) * rows + (r - r1)];
        }
    };

    // 다중 Area 선택의 전체 셀 순회: fn(cell)이 false를 돌려주면 중단.
    // (VB의 For Each Cell In Selection.Cells 대응 — Areas 단위로 순회)
    template <class Fn>
    inline void forEachCell(IDispatch* rangeDisp, Fn fn)
    {
        IDispatch* areas = getObject(rangeDisp, L"Areas");
        if (!areas) return;
        Releaser ra{ areas };
        const long nAreas = getLong(areas, L"Count", 0);
        for (long a = 1; a <= nAreas; ++a)
        {
            IDispatch* area = getObjectIdx(areas, L"Item", a);
            if (!area) continue;
            Releaser rArea{ area };
            IDispatch* cells = getObject(area, L"Cells");
            if (!cells) continue;
            Releaser rCells{ cells };
            const long n = (long)getDouble(cells, L"CountLarge", 0);
            for (long i = 1; i <= n; ++i)
            {
                IDispatch* cell = getObjectIdx(cells, L"Item", i);
                if (!cell) continue;
                Releaser rc{ cell };
                if (!fn(cell)) return;
            }
        }
    }
}
