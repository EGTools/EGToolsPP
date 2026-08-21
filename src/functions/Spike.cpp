// Spike.cpp — 진단(디버깅) UDF. Phase 0 스파이크에서 출발해 지원용으로 남긴 것만 유지.
//
//   * EG.VERSIONINFO — 감지된 Excel 에디션·동적배열·네이티브 함수 판정 표시.
//   * EG.SETLANG     — 다국어 디버깅: UI 언어 런타임 전환(저장 안 됨, 재시작 시
//                      자동 감지로 복귀). 인수 생략/미지원 코드면 변경 없이 현재
//                      언어 코드만 반환한다(조회 겸용 — 구 EG.LANG 대체).
//
// EG.HELLO(S1 연결 테스트)·EG.SEQ2D(S5 스필 검증)·EG.LANG(조회 전용)은 실사용이
// 없어 제거(D9, 2026-08-20). 스파이크 검증 이력은 plan/08 참조.

#include <xlOil/xlOil.h>

#include "../core/I18n.h"
#include "../core/Version.h"

#include <string>

using namespace xloil;

// S6/Phase1 diagnostic: report detected Excel edition + native-function decisions.
XLO_FUNC_START( EG_VERSIONINFO() )
{
    std::wstring s = L"year=" + std::to_wstring(egtools::core::excelYear())
        + L"; DA=" + (egtools::core::supportsDynamicArrays() ? L"1" : L"0")
        + L"; XLOOKUP_native=" + (egtools::core::hasNativeFunction(L"XLOOKUP") ? L"1" : L"0")
        + L"; IFS_native=" + (egtools::core::hasNativeFunction(L"IFS") ? L"1" : L"0")
        + L"; FOO_native=" + (egtools::core::hasNativeFunction(L"FOOBAR") ? L"1" : L"0");
    return returnValue(s);
}
XLO_FUNC_END(EG_VERSIONINFO)
    .threadsafe()
    .name(L"EG.VERSIONINFO")
    .category(L"EGTools")
    .help(L"Diagnostic: detected Excel edition and native-function decisions.");

XLO_FUNC_START( EG_SETLANG(const ExcelObj* code) )
{
    egtools::i18n::setCurrent(code->toString());
    return returnValue(egtools::i18n::current());
}
XLO_FUNC_END(EG_SETLANG)
    .name(L"EG.SETLANG")
    .category(L"EGTools")
    .help(L"Diagnostic: switches the EGTools++ UI language for this session "
          L"(ko/en/zh-CN/zh-TW/ja/es); omitted or unknown code just returns "
          L"the current language.")
    .arg(L"code", L"language code (optional)");
