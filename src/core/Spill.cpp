// Spill.cpp — array output that adapts to the host Excel.
//
// Modern Excel (2021/365) spills array returns natively. Legacy Excel (<=2019)
// shows only the top-left cell of an array return unless we force a multi-cell
// CSE array formula over a correctly-sized range. legacyResize() ports EGTools VB
// Resizer.vb (the ExcelDNA ArrayResizer idea): from the caller reference it
// compares the spilled array's dimensions to the formula's current range and, if
// they differ, defers a job that re-enters the formula as an array formula over
// the resized range. The UDF returns the array immediately so no #N/A flashes
// while the resize is queued.
//
// ── WHY COM, NOT THE XLM/C-API PATH ──────────────────────────────────────────
// Resizer.vb does the re-entry with macro-sheet/command XLM (xlfGetCell,
// xlcFormulaArray, …) on ExcelDNA's QueueAsMacro. The equivalent here is
// runExcelThread with the XLL_API / COM_API queue — but in THIS static XLL those
// queue callbacks never fire (xlOil switches to XLL context via COM, which isn't
// available in this build; verified: the deferred job never ran). Only the WINDOW
// queue fires (same finding as IMAGE / FxImage.cpp). WINDOW posts to Excel's main
// thread but does NOT establish an XLM command context, so xlc* commands are out.
// Instead the deferred resize uses the COM object model (Range.FormulaArray, via
// late-bound IDispatch like FxImage) — which works on the main thread regardless
// of XLM context, and accepts the A1-style formula directly (no R1C1 conversion).
//
// ⚠️ Validated on Excel 2016 x86. Build with -DEGTOOLS_SPILL_DEBUG to trace each
//    step to %TEMP%\egtools_spill.log.
//
// KNOWN LIMITATION (parity with EGTools VB Resizer.vb): the resize re-enters the
// formula via a programmatic COM write, which clears Excel's undo stack — so
// Ctrl+Z cannot undo a spilled formula on legacy Excel. This is inherent to
// emulated spill (ExcelDNA's resizer has the same open TODO); accepted as-is.

#include "Spill.h"
#include "Version.h"

#include <xlOil/ExcelObj.h>
#include <xlOil/ExcelArray.h>
#include <xlOil/ExcelCall.h>
#include <xlOil/ExcelRef.h>
#include <xlOil/ExcelThread.h>
#include <xlOil/AppObjects.h>       // ExcelRange (COM resolution)
#include <xlOil/ExcelTypeLib.h>     // Excel::Range (com())
#include <xlOil/Range.h>            // AddressStyle
#include <xlOil/StaticRegister.h>   // returnValue
#include <xlOil/XlCallSlim.h>

#include <windows.h>
#include <oleauto.h>
#include <string>

using namespace xloil;
using namespace msxll;

// Static XLL does not auto-connect COM (see FxLet.cpp / FxImage.cpp).
namespace xloil { namespace COM { bool connectCom(); } }

namespace egtools::core
{
    namespace
    {
        // Excel 2007+ grid limits (2016 target uses these).
        constexpr int kMaxRows = 1048576;
        constexpr int kMaxCols = 16384;

#ifdef EGTOOLS_SPILL_DEBUG
        void dbg(const std::wstring& msg)
        {
            wchar_t tmp[260]{}; GetTempPathW(260, tmp);
            std::wstring p = std::wstring(tmp) + L"egtools_spill.log";
            HANDLE h = CreateFileW(p.c_str(), FILE_APPEND_DATA,
                FILE_SHARE_READ | FILE_SHARE_WRITE, nullptr, OPEN_ALWAYS,
                FILE_ATTRIBUTE_NORMAL, nullptr);
            if (h == INVALID_HANDLE_VALUE) return;
            std::wstring line = msg + L"\r\n";
            int n = WideCharToMultiByte(CP_UTF8, 0, line.c_str(), -1, nullptr, 0, nullptr, nullptr);
            if (n > 1)
            {
                std::string u8(n - 1, '\0');
                WideCharToMultiByte(CP_UTF8, 0, line.c_str(), -1, u8.data(), n, nullptr, nullptr);
                DWORD wrote = 0; WriteFile(h, u8.data(), (DWORD)u8.size(), &wrote, nullptr);
            }
            CloseHandle(h);
        }
#else
        inline void dbg(const std::wstring&) {}
#endif

        // ── Late-bound IDispatch helpers (mirror FxImage.cpp). TODO: unify into a
        //    shared core/ComLateBind.h once a third consumer appears. ───────────
        DISPID dispidOf(IDispatch* d, const wchar_t* name)
        {
            DISPID id = DISPID_UNKNOWN;
            OLECHAR* n = const_cast<OLECHAR*>(name);
            if (FAILED(d->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id)))
                return DISPID_UNKNOWN;
            return id;
        }

        bool invokeRaw(IDispatch* d, const wchar_t* name, WORD flags,
                       VARIANT* result, VARIANT* args, UINT nArgs)
        {
            DISPID id = dispidOf(d, name);
            if (id == DISPID_UNKNOWN) return false;
            DISPPARAMS dp{ args, nullptr, nArgs, 0 };
            DISPID putId = DISPID_PROPERTYPUT;
            if (flags & DISPATCH_PROPERTYPUT) { dp.rgdispidNamedArgs = &putId; dp.cNamedArgs = 1; }
            return SUCCEEDED(d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags,
                                       &dp, result, nullptr, nullptr));
        }

        IDispatch* getObject(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return nullptr;
            IDispatch* obj = (r.vt == VT_DISPATCH) ? r.pdispVal : nullptr;
            if (!obj) VariantClear(&r);
            return obj;
        }

        std::wstring getBStr(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return L"";
            std::wstring s = (r.vt == VT_BSTR && r.bstrVal) ? r.bstrVal : L"";
            VariantClear(&r);
            return s;
        }

        bool getBoolProp(IDispatch* d, const wchar_t* name)
        {
            VARIANT r; VariantInit(&r);
            if (!invokeRaw(d, name, DISPATCH_PROPERTYGET, &r, nullptr, 0)) return false;
            VARIANT o; VariantInit(&o);
            bool v = SUCCEEDED(VariantChangeType(&o, &r, 0, VT_BOOL)) && o.boolVal != VARIANT_FALSE;
            VariantClear(&o); VariantClear(&r);
            return v;
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

        IDispatch* rangeDispatch(const std::wstring& addr)
        {
            xloil::ExcelRange rng(addr);
            IDispatch* d = nullptr;
            if (FAILED(rng.com().QueryInterface(IID_IDispatch, (void**)&d))) return nullptr;
            return d;
        }

        // Deferred (WINDOW queue, main thread): re-enter the first cell's formula as
        // a CSE array formula over `targetAddr` using the COM object model.
        void doResize(std::wstring firstCellAddr, std::wstring targetAddr)
        {
            dbg(L"doResize ENTER first=" + firstCellAddr + L" target=" + targetAddr);
            try
            {
                xloil::COM::connectCom();

                IDispatch* firstDisp = rangeDispatch(firstCellAddr);
                if (!firstDisp) { dbg(L"  firstCell COM NULL"); return; }
                Releaser fr{ firstDisp };

                // The formula currently in the cell (A1-style, includes '=').
                const std::wstring formula = getBStr(firstDisp, L"Formula");
                dbg(L"  formula=" + formula);
                if (formula.empty()) { dbg(L"  empty formula -> abort"); return; }

                // If the cell is already part of an array formula (re-resize), clear
                // the existing array region first so FormulaArray won't error.
                if (getBoolProp(firstDisp, L"HasArray"))
                {
                    if (IDispatch* oldArr = getObject(firstDisp, L"CurrentArray"))
                    {
                        Releaser orl{ oldArr };
                        callVoid(oldArr, L"ClearContents");
                        dbg(L"  cleared old array");
                    }
                }

                IDispatch* tgtDisp = rangeDispatch(targetAddr);
                if (!tgtDisp) { dbg(L"  target COM NULL"); return; }
                Releaser tr{ tgtDisp };

                // Set the array formula over the whole target range. The ensuing
                // recalc makes the caller range match the array size, so the UDF's
                // next legacyResize() sees equal dims and does NOT reschedule.
                putBStr(tgtDisp, L"FormulaArray", formula);
                dbg(L"  DONE FormulaArray set");
            }
            catch (...) { dbg(L"  EXCEPTION in doResize"); }
        }

        // Compare the spilled array to the caller range; queue a resize if needed.
        // Returns the array result either way (immediately, to avoid #N/A flashing).
        ExcelObj* legacyResize(ExcelObj&& result)
        {
            if (result.type() != ExcelType::Multi)
                return returnValue(std::move(result));

            ExcelArray arr(result);
            const int rows = (int)arr.nRows();
            const int cols = (int)arr.nCols();
            if (rows == 0 || cols == 0)
                return returnValue(std::move(result));

            // Caller reference (works in a normal UDF; non-ref → just return array).
            auto [callerObj, rc] = tryCallExcel(xlfCaller);
            if (rc != 0 ||
                !(callerObj.type() == ExcelType::SRef || callerObj.type() == ExcelType::Ref))
            {
                dbg(L"legacyResize: caller not a ref -> return as-is");
                return returnValue(std::move(result));
            }

            ExcelRef caller(callerObj);
            auto [r0, c0, r1, c1] = caller.bounds();
            const int callerRows = r1 - r0 + 1;
            const int callerCols = c1 - c0 + 1;
            dbg(L"legacyResize: array " + std::to_wstring(rows) + L"x" + std::to_wstring(cols) +
                L" caller " + std::to_wstring(callerRows) + L"x" + std::to_wstring(callerCols));

            // Already the right size (a previous resize ran, or user CSE'd it).
            if (callerRows == rows && callerCols == cols)
                return returnValue(std::move(result));

            // Would spill past the sheet edge → #VALUE! (cannot message from here).
            if (r0 + rows - 1 > kMaxRows - 1 || c0 + cols - 1 > kMaxCols - 1)
            {
                dbg(L"legacyResize: beyond sheet edge -> #VALUE!");
                return returnValue(CellError::Value);
            }

            // Defer the array-formula re-entry to the main thread (cannot modify
            // cells during calc). Capture A1 addresses; resolve to COM in doResize.
            const std::wstring firstCellAddr =
                caller.range(0, 0, 0, 0).address(AddressStyle::A1);
            const std::wstring targetAddr =
                caller.range(0, 0, rows - 1, cols - 1).address(AddressStyle::A1);
            dbg(L"legacyResize: scheduling resize to " + targetAddr);

            xloil::runExcelThread(
                [firstCellAddr, targetAddr] { doResize(firstCellAddr, targetAddr); },
                xloil::ExcelRunQueue::WINDOW | xloil::ExcelRunQueue::ENQUEUE);

            return returnValue(std::move(result));
        }
    }

    xloil::ExcelObj* output(xloil::ExcelObj&& result)
    {
        // Modern Excel spills array returns natively — nothing extra to do.
        if (supportsDynamicArrays())
            return xloil::returnValue(std::move(result));

        // Legacy Excel (<=2019): force a CSE array formula sized to the result.
        return legacyResize(std::move(result));
    }
}
