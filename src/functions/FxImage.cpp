// FxImage.cpp — IMAGE (macro-type), ported from EGTools VB (FXL_2024 IMAGE +
// B06_Image helpers).
//
// Excel <365 has no in-cell image value. Like the VB version, IMAGE downloads /
// loads a picture and inserts it as a FLOATING Picture shape sized to the calling
// (merged) cell, returning "". A worksheet UDF may not modify the workbook during
// calculation, so the insertion is deferred to Excel's main thread (after calc)
// via xloil::runExcelThread.
//
// ── WHY LATE-BOUND COM (2026-06-28, verified on Excel 2016 x86) ───────────────
// The first attempt used TYPED #import COM (Excel::ShapesPtr / _WorksheetPtr).
// Navigating Range→Worksheet→Shapes through #import smart pointers makes raw
// vtable calls with NO automatic apartment marshalling: in the deferred context
// the proxy was invalid and Excel hard-crashed ("RPC server unavailable", not a
// catchable C++ exception). LET (FxLet.cpp) drives COM with LATE-BOUND IDispatch
// (GetIDsOfNames/Invoke) and works reliably — so IMAGE uses the same late-bound
// path for the whole Worksheet/Shapes/AddPicture chain. No typed-COM member calls
// remain on the crashing path. The deferred insert is scheduled on the WINDOW
// queue: in this static XLL the COM_API queue callback never fires, whereas WINDOW
// posts to Excel's main thread where COM is usable after connectCom().

#include "../core/Registry.h"

#include <xlOil/xlOil.h>
#include <xlOil/Caller.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/AppObjects.h>
#include <xlOil/ExcelTypeLib.h>   // only for ExcelRange::com() return type
#include <windows.h>
#include <oleauto.h>
#include <urlmon.h>
#include <string>
#include <vector>
#include <algorithm>

#pragma comment(lib, "urlmon.lib")

using namespace xloil;

// Static XLL does not auto-connect COM (see FxLet.cpp).
namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::functions
{
    namespace
    {
        // MsoTriState / placement constants (avoid pulling in typed enums).
        constexpr int kMsoFalse = 0;
        constexpr int kMsoTrue = -1;
        constexpr int kXlMoveAndSize = 1;

        bool isUrl(const std::wstring& s)
        {
            return s.rfind(L"http://", 0) == 0 || s.rfind(L"https://", 0) == 0;
        }

        // "[Book1]Sheet1!$A$1" / "'[Book1]My Sheet'!$A$1" → sheet="Sheet1", cell="$A$1"
        void parseSheetCell(const std::wstring& addr, std::wstring& sheet, std::wstring& cell)
        {
            size_t bang = addr.rfind(L'!');
            if (bang == std::wstring::npos) { sheet.clear(); cell = addr; return; }
            std::wstring sh = addr.substr(0, bang);
            cell = addr.substr(bang + 1);
            if (sh.size() >= 2 && sh.front() == L'\'' && sh.back() == L'\'')
                sh = sh.substr(1, sh.size() - 2);
            size_t rb = sh.find(L']');
            if (rb != std::wstring::npos) sh = sh.substr(rb + 1);
            sheet = sh;
        }

        // ── Minimal late-bound IDispatch helpers (mirror FxLet.cpp comEvaluate) ──
        // Every Excel object call below goes through IDispatch::Invoke, so there are
        // no #import vtable proxies to break across the deferred COM context.

        DISPID dispidOf(IDispatch* d, const wchar_t* name)
        {
            DISPID id = DISPID_UNKNOWN;
            OLECHAR* n = const_cast<OLECHAR*>(name);
            if (FAILED(d->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id)))
                return DISPID_UNKNOWN;
            return id;
        }

        // Raw Invoke. `args` are in NATURAL order; we reverse into DISPPARAMS here.
        bool invokeRaw(IDispatch* d, const wchar_t* name, WORD flags,
                       VARIANT* result, VARIANT* args, UINT nArgs)
        {
            DISPID id = dispidOf(d, name);
            if (id == DISPID_UNKNOWN) return false;

            std::vector<VARIANT> rev(nArgs);
            for (UINT i = 0; i < nArgs; ++i) rev[i] = args[nArgs - 1 - i];  // COM wants reverse order

            DISPPARAMS dp{ nArgs ? rev.data() : nullptr, nullptr, nArgs, 0 };
            DISPID putId = DISPID_PROPERTYPUT;
            if (flags & DISPATCH_PROPERTYPUT) { dp.rgdispidNamedArgs = &putId; dp.cNamedArgs = 1; }

            return SUCCEEDED(d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                                       &dp, result, nullptr, nullptr));
        }

        // Property/method returning an IDispatch* (caller owns the returned ref).
        IDispatch* getObject(IDispatch* d, const wchar_t* name, VARIANT* args = nullptr, UINT n = 0)
        {
            VARIANT r; VariantInit(&r);
            WORD flags = DISPATCH_PROPERTYGET | (n ? DISPATCH_METHOD : 0);
            if (!invokeRaw(d, name, flags, &r, args, n)) return nullptr;
            IDispatch* obj = (r.vt == VT_DISPATCH) ? r.pdispVal : nullptr;
            if (!obj) VariantClear(&r);   // not a dispatch → release whatever came back
            return obj;                   // VT_DISPATCH: transfer ownership, do not clear
        }

        double getDouble(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return 0.0;
            VARIANT out; VariantInit(&out);
            double v = SUCCEEDED(VariantChangeType(&out, &r, 0, VT_R8)) ? out.dblVal : 0.0;
            VariantClear(&out); VariantClear(&r);
            return v;
        }

        long getLong(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return 0;
            VARIANT out; VariantInit(&out);
            long v = SUCCEEDED(VariantChangeType(&out, &r, 0, VT_I4)) ? out.lVal : 0;
            VariantClear(&out); VariantClear(&r);
            return v;
        }

        std::wstring getBStr(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return L"";
            std::wstring s = (r.vt == VT_BSTR && r.bstrVal) ? r.bstrVal : L"";
            VariantClear(&r);
            return s;
        }

        void putLong(IDispatch* d, const wchar_t* name, long val)
        {
            VARIANT a; VariantInit(&a); a.vt = VT_I4; a.lVal = val;
            invokeRaw(d, name, DISPATCH_PROPERTYPUT, nullptr, &a, 1);
            VariantClear(&a);
        }

        void putBStr(IDispatch* d, const wchar_t* name, const std::wstring& val)
        {
            VARIANT a; VariantInit(&a); a.vt = VT_BSTR; a.bstrVal = SysAllocString(val.c_str());
            invokeRaw(d, name, DISPATCH_PROPERTYPUT, nullptr, &a, 1);
            VariantClear(&a);
        }

        void callVoid(IDispatch* d, const wchar_t* name)
        {
            invokeRaw(d, name, DISPATCH_METHOD, nullptr, nullptr, 0);
        }

        struct Releaser { IDispatch* p; ~Releaser() { if (p) p->Release(); } };

        // Runs on Excel's main thread (deferred, COM_API queue). Inserts the picture.
        // Returns false → ComBusyException-style retry is handled by the queue wrapper;
        // we simply return (best-effort, fire-and-forget like the VB version).
        void insertImage(std::wstring fullAddr, std::wstring cell, std::wstring source, int mode,
                         bool deleteLocalAfter = false)
        {
            std::wstring tempFile;
            bool isTemp = false;
            try
            {
                xloil::COM::connectCom();

                std::wstring path = source;
                if (isUrl(source))
                {
                    wchar_t tmp[MAX_PATH]{};
                    GetTempPathW(MAX_PATH, tmp);
                    std::wstring png = std::wstring(tmp) + L"egimg_" +
                        std::to_wstring(GetTickCount64()) + L".img";
                    if (URLDownloadToFileW(nullptr, source.c_str(), png.c_str(), 0, nullptr) != S_OK)
                        return;
                    path = png; tempFile = png; isTemp = true;
                }
                else if (GetFileAttributesW(source.c_str()) == INVALID_FILE_ATTRIBUTES)
                    return;
                else if (deleteLocalAfter) { tempFile = source; isTemp = true; }

                // Resolve the caller range via xlOil (robust COM handling), then get
                // its IDispatch — everything after this is late-bound.
                xloil::ExcelRange rng(fullAddr);
                IDispatch* rangeDisp = nullptr;
                if (FAILED(rng.com().QueryInterface(IID_IDispatch, (void**)&rangeDisp)) || !rangeDisp)
                {
                    if (isTemp) DeleteFileW(path.c_str());
                    return;
                }
                Releaser rd{ rangeDisp };

                // MergeArea geometry (Left/Top/Width/Height) — all late-bound.
                IDispatch* mergeDisp = getObject(rangeDisp, L"MergeArea");
                IDispatch* geom = mergeDisp ? mergeDisp : rangeDisp;
                Releaser mr{ mergeDisp };
                const double L = getDouble(geom, L"Left");
                const double T = getDouble(geom, L"Top");
                const double W = getDouble(geom, L"Width");
                const double H = getDouble(geom, L"Height");

                // Worksheet → Shapes (the path that crashed under typed COM).
                IDispatch* wsDisp = getObject(rangeDisp, L"Worksheet");
                if (!wsDisp) { if (isTemp) DeleteFileW(path.c_str()); return; }
                Releaser wr{ wsDisp };
                IDispatch* shapesDisp = getObject(wsDisp, L"Shapes");
                if (!shapesDisp) { if (isTemp) DeleteFileW(path.c_str()); return; }
                Releaser sr{ shapesDisp };

                std::wstring shName, c2; parseSheetCell(fullAddr, shName, c2);
                shName.erase(std::remove(shName.begin(), shName.end(), L' '), shName.end());
                const std::wstring imageName = L"EG.Image_" + shName + L"_" + cell;

                // Remove any existing same-named shape (re-calc / formula copy).
                long cnt = getLong(shapesDisp, L"Count");
                for (long i = cnt; i >= 1; --i)
                {
                    VARIANT idx; VariantInit(&idx); idx.vt = VT_I4; idx.lVal = i;
                    IDispatch* shp = getObject(shapesDisp, L"Item", &idx, 1);
                    VariantClear(&idx);
                    if (!shp) continue;
                    Releaser shr{ shp };
                    if (getBStr(shp, L"Name") == imageName) callVoid(shp, L"Delete");
                }

                // Shapes.AddPicture(Filename, LinkToFile, SaveWithDocument, L, T, W, H)
                VARIANT a[7];
                for (auto& v : a) VariantInit(&v);
                a[0].vt = VT_BSTR; a[0].bstrVal = SysAllocString(path.c_str());
                a[1].vt = VT_I4;   a[1].lVal = kMsoFalse;            // LinkToFile
                a[2].vt = VT_I4;   a[2].lVal = kMsoTrue;             // SaveWithDocument (embed)
                a[3].vt = VT_R4;   a[3].fltVal = (float)(L + 0.3);
                a[4].vt = VT_R4;   a[4].fltVal = (float)(T + 0.3);
                a[5].vt = VT_R4;   a[5].fltVal = (float)(W - 0.6);
                a[6].vt = VT_R4;   a[6].fltVal = (float)(H - 0.6);
                VARIANT picV; VariantInit(&picV);
                bool added = invokeRaw(shapesDisp, L"AddPicture",
                                       DISPATCH_METHOD | DISPATCH_PROPERTYGET, &picV, a, 7);
                for (auto& v : a) VariantClear(&v);

                if (added && picV.vt == VT_DISPATCH && picV.pdispVal)
                {
                    IDispatch* pic = picV.pdispVal;
                    Releaser pr{ pic };
                    putBStr(pic, L"Name", imageName);
                    putLong(pic, L"Placement", kXlMoveAndSize);
                    if (mode == 0)
                        putLong(pic, L"LockAspectRatio", kMsoTrue);   // keep ratio within cell
                }
                else VariantClear(&picV);

                if (isTemp) DeleteFileW(path.c_str());
            }
            catch (...)
            {
                if (isTemp && !tempFile.empty()) DeleteFileW(tempFile.c_str());
            }
        }
    }

    // Shared entry point (ImageInsert.h) — also used by the barcode functions.
    void queueInsertPicture(const std::wstring& fullAddr, const std::wstring& cell,
                            const std::wstring& source, int mode, bool deleteLocalAfter)
    {
        // WINDOW (not COM_API): in this static-XLL context the COM_API queue
        // callback never fires, but WINDOW posts to Excel's main thread where
        // COM is usable after connectCom().
        xloil::runExcelThread(
            [fullAddr, cell, source, mode, deleteLocalAfter]
            { insertImage(fullAddr, cell, source, mode, deleteLocalAfter); },
            xloil::ExcelRunQueue::WINDOW | xloil::ExcelRunQueue::ENQUEUE);
    }

    void registerImage()
    {
        // IMAGE(source, [alt_text], [resize_mode], [height], [width]).
        // macro=true for xlfCaller; insertion is deferred to the main thread on the
        // COM_API queue (COM-ready, auto-retrying) — see header note.
        egtools::core::registerFn(L"IMAGE",
            [](const ExcelObj& src, const ExcelObj& /*alt*/, const ExcelObj& modeA,
               const ExcelObj& /*h*/, const ExcelObj& /*w*/) -> ExcelObj*
            {
                try
                {
                    std::wstring source = src.toString();
                    if (source.empty()) return returnValue(CellError::Value);
                    const int mode = modeA.isMissing() ? 0 : modeA.get<int>(0);

                    const std::wstring fullAddr = xloil::CallerInfo().address();
                    std::wstring sheet, cell;
                    parseSheetCell(fullAddr, sheet, cell);
                    if (cell.empty()) return returnValue(CellError::Ref);

                    queueInsertPicture(fullAddr, cell, source, mode, /*deleteLocalAfter*/ false);

                    return returnValue(ExcelObj(std::wstring_view(L"")));
                }
                catch (...) { return returnValue(CellError::Value); }
            },
            /*macro*/ true);
    }
}
