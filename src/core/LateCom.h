// LateCom.h — minimal late-bound IDispatch helpers shared by macro-type
// functions that walk the Excel object model (FxColor, FxPublicApi, …).
// Late-bound Invoke is deliberate: typed #import proxies crash across deferred
// COM contexts in this static XLL (see FxImage.cpp header note).
#pragma once
#include <windows.h>
#include <oleauto.h>
#include <string>
#include <vector>

namespace egtools::latecom
{
    inline DISPID dispidOf(IDispatch* d, const wchar_t* name)
    {
        DISPID id = DISPID_UNKNOWN;
        OLECHAR* n = const_cast<OLECHAR*>(name);
        if (FAILED(d->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id)))
            return DISPID_UNKNOWN;
        return id;
    }

    // Raw Invoke. `args` are in NATURAL order; reversed into DISPPARAMS here.
    inline bool invokeRaw(IDispatch* d, const wchar_t* name, WORD flags,
                          VARIANT* result, VARIANT* args, UINT nArgs)
    {
        DISPID id = dispidOf(d, name);
        if (id == DISPID_UNKNOWN) return false;
        std::vector<VARIANT> rev(nArgs);
        for (UINT i = 0; i < nArgs; ++i) rev[i] = args[nArgs - 1 - i];
        DISPPARAMS dp{ nArgs ? rev.data() : nullptr, nullptr, nArgs, 0 };
        DISPID putId = DISPID_PROPERTYPUT;
        if (flags & DISPATCH_PROPERTYPUT) { dp.rgdispidNamedArgs = &putId; dp.cNamedArgs = 1; }
        return SUCCEEDED(d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                                   &dp, result, nullptr, nullptr));
    }

    // Property/method returning IDispatch* (caller owns the returned ref).
    inline IDispatch* getObject(IDispatch* d, const wchar_t* name,
                                VARIANT* args = nullptr, UINT n = 0)
    {
        VARIANT r; VariantInit(&r);
        WORD flags = DISPATCH_PROPERTYGET | (n ? DISPATCH_METHOD : 0);
        if (!invokeRaw(d, name, flags, &r, args, n)) return nullptr;
        IDispatch* obj = (r.vt == VT_DISPATCH) ? r.pdispVal : nullptr;
        if (!obj) VariantClear(&r);
        return obj;
    }

    inline IDispatch* getObjectIdx(IDispatch* d, const wchar_t* name, long idx)
    {
        VARIANT a; VariantInit(&a); a.vt = VT_I4; a.lVal = idx;
        IDispatch* o = getObject(d, name, &a, 1);
        VariantClear(&a);
        return o;
    }

    inline long getLong(IDispatch* d, const wchar_t* name, long fallback = 0)
    {
        VARIANT r; VariantInit(&r);
        if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return fallback;
        VARIANT out; VariantInit(&out);
        long v = SUCCEEDED(VariantChangeType(&out, &r, 0, VT_I4)) ? out.lVal : fallback;
        VariantClear(&out); VariantClear(&r);
        return v;
    }

    inline double getDouble(IDispatch* d, const wchar_t* name, double fallback = 0.0)
    {
        VARIANT r; VariantInit(&r);
        if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return fallback;
        VARIANT out; VariantInit(&out);
        double v = SUCCEEDED(VariantChangeType(&out, &r, 0, VT_R8)) ? out.dblVal : fallback;
        VariantClear(&out); VariantClear(&r);
        return v;
    }

    inline bool getBool(IDispatch* d, const wchar_t* name, bool fallback = false)
    {
        VARIANT r; VariantInit(&r);
        if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return fallback;
        VARIANT out; VariantInit(&out);
        bool v = SUCCEEDED(VariantChangeType(&out, &r, 0, VT_BOOL))
                     ? (out.boolVal != VARIANT_FALSE) : fallback;
        VariantClear(&out); VariantClear(&r);
        return v;
    }

    inline std::wstring getBStr(IDispatch* d, const wchar_t* name)
    {
        VARIANT r; VariantInit(&r);
        if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return L"";
        std::wstring s = (r.vt == VT_BSTR && r.bstrVal) ? r.bstrVal : L"";
        VariantClear(&r);
        return s;
    }

    struct Releaser
    {
        IDispatch* p;
        explicit Releaser(IDispatch* d) : p(d) {}
        ~Releaser() { if (p) p->Release(); }
        Releaser(const Releaser&) = delete;
        Releaser& operator=(const Releaser&) = delete;
    };
}
