#include "Registry.h"

#include <algorithm>

namespace egtools::functions {
    void registerLookup();   // FxLookup.cpp
    void registerLogic();    // FxLogic.cpp
    void registerText();     // FxText.cpp
    void registerWeb();      // FxWeb.cpp
    void registerDate();     // FxDate.cpp
    void registerInfo();     // FxInfo.cpp
    void registerAgg();      // FxAgg.cpp
    void registerGen();      // FxGen.cpp
    void registerFilterSort(); // FxFilterSort.cpp
    void registerArrayShape(); // FxArrayShape.cpp
    void registerRegex();      // FxRegex.cpp
    void registerGroup();      // FxGroup.cpp
    void registerMath();       // FxMath.cpp
    void registerBit();        // FxBit.cpp
    void registerLet();        // FxLet.cpp
    void registerImage();      // FxImage.cpp
    void registerGoogle();     // FxGoogle.cpp
    void registerEgText();     // FxEgText.cpp
    void registerEgData();     // FxEgData.cpp
    void registerEgQC();       // FxEgQC.cpp
    void registerEgDate();     // FxEgDate.cpp
    void registerEgJson();     // FxEgJson.cpp
    void registerEgUtil();     // FxEgUtil.cpp
    void registerBarcode();    // FxBarcode.cpp
    void registerColor();      // FxColor.cpp
    void registerPublicApi();  // FxPublicApi.cpp
}

namespace egtools::core
{
    namespace
    {
        std::vector<std::shared_ptr<xloil::RegisteredWorksheetFunc>> g_funcs;
    }

    namespace detail
    {
        void keep(std::shared_ptr<xloil::RegisteredWorksheetFunc> reg)
        {
            if (reg) g_funcs.push_back(std::move(reg));
        }
    }

    void registerRawFn(const std::wstring& bareName,
                       xloil::DynamicExcelFunc<> fn,
                       bool macro, bool threadsafe)
    {
        std::wstring desc;
        std::vector<egtools::i18n::ArgMeta> rep;
        egtools::i18n::RepeatSpec spec;
        egtools::i18n::func(bareName, desc, rep, &spec);

        // EG.<F> / x<F> / <F> per host — see core::registeredName().
        const std::wstring name = registeredName(bareName);

        // Record the runtime name while the UDF executes, for core::output()'s
        // outermost-call check on legacy hosts (see detail::namedUdf). Also the
        // last catch before xlOil turns an escaped exception into a TEXT cell
        // value — return #VALUE! instead (plan/22 §M2, same as namedUdfImpl).
        xloil::DynamicExcelFunc<> named =
            [name, fn = std::move(fn)](const xloil::FuncInfo& info,
                                       const xloil::ExcelObj** args)
            {
                UdfNameScope scope(name);
                try { return fn(info, args); }
                catch (...) { return xloil::returnValue(xloil::CellError::Value); }
            };

        // Build the argument list for a given arg-count cap and register. Returns
        // true on success. Wrapping in shared_ptr first: registerFunc() relies on
        // shared_from_this (mirrors xloil::RegisterLambda::registerFunc).
        auto attempt = [&](size_t cap) -> bool
        {
            std::vector<egtools::i18n::ArgMeta> full;
            if (spec.has)
            {
                egtools::i18n::RepeatSpec s = spec;
                if ((int)cap < s.head + s.period + s.tail) cap = s.head + s.period + s.tail;
                s.max = (int)cap;
                full = egtools::i18n::expandRepeatArgs(rep, s);
            }
            else full = rep;

            auto info = std::make_shared<xloil::FuncInfo>();
            info->name = name;
            info->help = desc;
            info->category = L"EGTools";   // 함수 마법사 전용 범주
            info->options = macro ? (unsigned)xloil::FuncInfo::MACRO_TYPE
                                  : (threadsafe ? (unsigned)xloil::FuncInfo::THREAD_SAFE : 0u);
            info->args.reserve(full.size());
            for (const auto& a : full)
            {
                const int type = xloil::FuncArg::Obj | (a.opt ? xloil::FuncArg::Optional : 0);
                info->args.emplace_back(a.name, a.help, type);
            }
            try
            {
                auto reg = std::make_shared<xloil::LambdaSpec<>>(info, named)->registerFunc();
                if (reg) { detail::keep(reg); return true; }
            }
            catch (...) {}
            return false;
        };

        // Try the full arity, then progressively smaller (some hosts may reject a
        // very large argument count). The repeat layout stays valid at each step.
        if (!spec.has) { attempt(rep.size()); return; }
        const int floorN = spec.head + spec.period + spec.tail;
        // NB: xlOil's stock Win32 thunk builder hard-throws above 16 arguments
        // (asmjit FuncDetail limit) — patched out by
        // patches/xloil-0.22.1-win32-thunk-args.patch (hand-rolled x86 thunk),
        // so both platforms register the full arity. Should the patch ever be
        // missing, the descending ladder below still lands on a working cap.
        // Largest head + k*period + tail that fits `cap` (≥ floorN).
        auto aligned = [&](int cap) -> int
        {
            if (cap <= floorN) return floorN;
            return spec.head + spec.tail
                 + (cap - spec.head - spec.tail) / spec.period * spec.period;
        };
        const int top = spec.max;
        for (int cap : { aligned(top),
                         aligned((std::min)(top, spec.head + 60 * spec.period + spec.tail)),
                         aligned((std::min)(top, spec.head + 8 * spec.period + spec.tail)),
                         floorN })
            if (cap <= spec.max && attempt((size_t)cap)) return;
    }

    void registerFunctions()
    {
        // Isolate each group: a throw in one registration must NOT abort the rest
        // (otherwise e.g. a failure in registerLet would silently drop the last
        // group, registerImage — exactly the "some functions missing" symptom).
        auto group = [](void (*fn)()) { try { fn(); } catch (...) {} };

        group(egtools::functions::registerLookup);   // XLOOKUP …
        group(egtools::functions::registerLogic);    // IFNA …
        group(egtools::functions::registerText);     // UNICHAR, UNICODE, NUMBERVALUE …
        group(egtools::functions::registerWeb);      // ENCODEURL …
        group(egtools::functions::registerDate);     // DAYS, ISOWEEKNUM …
        group(egtools::functions::registerInfo);     // FORMULATEXT, ISFORMULA, SHEET, SHEETS
        group(egtools::functions::registerAgg);      // IFS, SWITCH, MAXIFS, MINIFS
        group(egtools::functions::registerGen);      // SEQUENCE, RANDARRAY
        group(egtools::functions::registerFilterSort); // XMATCH, FILTER, SORT, SORTBY, UNIQUE
        group(egtools::functions::registerArrayShape); // VSTACK, HSTACK, TAKE, DROP, TO*, CHOOSE*, EXPAND
        group(egtools::functions::registerRegex);      // REGEXTEST, REGEXEXTRACT, REGEXREPLACE
        group(egtools::functions::registerGroup);      // GROUPBY, PIVOTBY (text aggregator)
        group(egtools::functions::registerMath);       // PERCENTOF, COMBINA, …, CEILING.MATH, BASE, RRI …
        group(egtools::functions::registerBit);        // BITAND/OR/XOR/LSHIFT/RSHIFT
        group(egtools::functions::registerLet);        // LET (macro-type, formula rewrite)
        group(egtools::functions::registerImage);      // IMAGE — late-bound IDispatch + COM_API
        //   queue (retry of typed-COM Worksheet.Shapes crash, Excel 2016/x86). plan/15 §2.
        group(egtools::functions::registerGoogle);     // GOOGLETRANSLATE/GTRS/IMPORT*/QUERY (Google 호환, plan/20)
        group(egtools::functions::registerEgText);     // TEXTNUMSORT/TRIMENDS/STREXT/… (EGTools 텍스트, plan/20 E1)
        group(egtools::functions::registerEgData);     // UNPIVOT/BOMTREE/EXPLODE/… (EGTools 데이터, plan/20 E2)
        group(egtools::functions::registerEgQC);       // SAMPLING*/CP/CPK (EGTools 품질, plan/20 E3)
        group(egtools::functions::registerEgDate);     // TOLUNAR/TOSOLAR/KOREANHOLIDAYS/… (EGTools 날짜, plan/20 E4)
        group(egtools::functions::registerEgJson);     // JSONFILTER (plan/20 E7)
        group(egtools::functions::registerEgUtil);     // DIRFOLDER/RESTAPI/EXRATE/EVAL/SHEETLIST/TOTALPAGES (plan/20 E5·E8)
        group(egtools::functions::registerBarcode);    // BARCODE/QRCODE/DATAMATRIX/CODE128/READBARCODE/READGS1 (plan/22 B2)
        group(egtools::functions::registerColor);      // DISPLAYCOLOR/COUNTIFCOLOR/SUMIFCOLOR/ISVISIBLE/VISIBLEAGGR/ILOOKUP (plan/22 B3)
        group(egtools::functions::registerPublicApi);  // SEARCHADDRESS/BRNSTATUS/GEO*/ADDRESSMAP (plan/22 B4, 키=HKCU ApiKeys)
    }

    void unregisterFunctions()
    {
        g_funcs.clear();   // destructors unregister each function
    }
}
