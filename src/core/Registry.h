// Registry.h — dynamic worksheet-function registration with version-adaptive
// naming and i18n help/argument text.
//
// Functions are registered at AutoOpen (after i18n load) via xlOil's dynamic
// RegisterLambda, because the registered NAME depends on the running Excel
// (core::registeredName): EG.<F> when the host has the native function (avoid
// clash), x<F> for keyword-conflict names on non-native hosts (SORT/FILTER/LET
// — bare entry is refused or hijacked there), bare <F> otherwise (drop-in;
// always bare for EGTools-only UDFs).

#pragma once
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include <xlOil/DynamicRegister.h>

#include "Version.h"
#include "I18n.h"

namespace egtools::core
{
    namespace detail
    {
        // Keep a registration handle alive for the add-in's lifetime.
        void keep(std::shared_ptr<xloil::RegisteredWorksheetFunc> reg);
    }

    // macro=true grants macro-sheet API access. REQUIRED for any function that
    //   * calls XLM helpers (GET.CELL / GET.WORKBOOK / ISFORMULA, …), or
    //   * uses the Excel object model directly (COM: Application/Worksheets/…),
    //   * or reads/inspects other cells.
    // (Plain compute functions and MSXML/WinHTTP — which don't touch Excel's API —
    //  do NOT need it.)
    // threadsafe는 Excel 멀티스레드 재계산(MTR) 참여 여부.
    //
    // 기본값 = arrayThreadsafe(): 배열 반환(core::output) 가능성이 있는 순수
    // 함수에 안전한 최대치 — 모던(DA) 호스트에서는 output()이 순수 반환이라
    // 항상 안전하고, 구버전 호스트에서는 CSE resize 경로(Spill.cpp)가 worker
    // 스레드에서 안전하지 않은 XLM(주소 변환)을 호출하므로 비활성화된다.
    //
    // 명시 지정 규칙:
    //   * true  — 결과가 항상 스칼라인 순수 함수(legacy에서도 resize 경로에
    //             들어가지 않음). 특히 느린 네트워크 스칼라 함수에 효과 큼.
    //   * false — 함수 실행 중 COM을 만드는 경우(FILTERXML/IMPORTFEED/
    //             IMPORTHTML/QUERY): MTR 워커 스레드는 COM 아파트가 초기화되지
    //             않아 CoCreateInstance가 실패/불안정하다.
    //   * macro=true와는 상호 배타 — macro가 우선한다.
    // 전역 가변 상태를 갖는 함수는 없어야 한다(RNG는 thread_local, i18n은 mutex).
    inline bool arrayThreadsafe() { return supportsDynamicArrays(); }

    namespace detail
    {
        // Register one (name -> fn) pair, pulling i18n help/args by bareName.
        template <class TFunc>
        void registerOne(const std::wstring& name, const std::wstring& bareName,
                         const TFunc& fn, bool macro, bool threadsafe)
        {
            std::wstring desc;
            std::vector<egtools::i18n::ArgMeta> args;
            egtools::i18n::func(bareName, desc, args);

            auto builder = xloil::RegisterLambda<>(fn);   // copies fn
            builder.name(name);
            builder.category(L"EGTools");   // 함수 마법사 전용 범주
            if (!desc.empty()) builder.help(desc);
            for (const auto& a : args)
                builder.arg(a.name.c_str(), a.help.empty() ? nullptr : a.help.c_str());
            if (macro) builder.macro();
            else if (threadsafe) builder.threadsafe();

            // Isolate: one function's registration failure must not abort others.
            try { keep(builder.registerFunc()); }
            catch (...) {}
        }
    }

    template <class TFunc>
    void registerFn(const std::wstring& bareName, TFunc&& fn, bool macro = false,
                    bool threadsafe = arrayThreadsafe())
    {
        // EG.<F> / x<F> / <F> per host — see core::registeredName().
        const std::wstring name = registeredName(bareName);
        detail::registerOne(name, bareName, std::forward<TFunc>(fn), macro, threadsafe);

        // NOTE: we do NOT register the "_xlfn." alias of new functions.
        //
        // A modern-authored .xlsx stores new functions under an internal token
        // (e.g. _xlfn.XLOOKUP, dynamic-array ones as _xlfn._xlws.FILTER). It is
        // tempting to also register a UDF under that name so old Excel resolves
        // such tokens to ours — but Excel RESERVES the "_xlfn." prefix and rejects
        // xlfRegister for it (Excel 2016: "command failed"), so the alias can never
        // bind. (Worse, an earlier version let that failure throw and abort all
        // remaining registrations.) A stored _xlfn.NAME token is therefore handled
        // separately — by rewriting it to the bare NAME (which our UDF above does
        // provide) via an explicit compatibility command. See
        // plan/13_x86_2016_런타임수정.md §"_xlfn. 문서 호환". storedName() is kept
        // for that rewrite mapping.
    }

    // Register a function whose argument list is specified at runtime — for
    // variadic-style functions (LET's name/value pairs, TEXTJOIN's text1,text2,…).
    // Reads the i18n description + repeat spec by bareName; when a repeat spec is
    // present it synthesises the argument list up to spec.max (so Excel accepts
    // that many args), otherwise uses the listed args verbatim. Uses the same
    // version-adaptive naming (EG.-prefix when native) as registerFn. The callback
    // receives Excel's raw argument array and the FuncInfo (use info.numArgs()).
    void registerRawFn(const std::wstring& bareName,
                       xloil::DynamicExcelFunc<> fn,
                       bool macro = false,
                       bool threadsafe = arrayThreadsafe());

    // Register all EGTools worksheet functions. Call once at AutoOpen.
    void registerFunctions();
    void unregisterFunctions();
}
