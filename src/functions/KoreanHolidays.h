// KoreanHolidays.h — 대한민국 공휴일 공용 헬퍼 (plan/23 §3-4).
//
// KOREANHOLIDAYS UDF와 리본 달력/일정표(CmdCalendar)가 같은 계산을 공유한다.
// 구현은 FxEgDate.cpp(음력 테이블·직렬화 헬퍼와 같은 TU).
#pragma once
#include <map>
#include <string>

namespace egtools::dates
{
    // 내장 계산(양·음력 공휴일 + 대체공휴일 규칙). serial(1900 기준) → 이름
    // (겹치면 ", "로 연결). korean=false면 영어 명칭.
    void builtinHolidays(int year, bool mayDay, bool korean,
                         std::map<int, std::wstring>& list);

    // 공공데이터포털 특일정보 API(getRestDeInfo)로 임시공휴일·선거일 등
    // 내장 계산에 없는 항목을 병합한다(결정 3). apiKey가 비어 있으면
    // HKCU\Software\EGTools\ApiKeys\datago 저장 키 사용; 인수로 오면 저장 후
    // 사용(공공 API 키 정책과 동일). 키 없음/네트워크 실패는 조용히 0 반환.
    // 연도별 세션 캐시(실패 포함) — 재계산마다 재조회하지 않는다.
    int mergeApiHolidays(int year, std::map<int, std::wstring>& list,
                         const std::wstring& apiKey = L"");

    // serial(양력)의 음력 표기(년/월/일). 범위 밖이면 false.
    // 달력/일정표의 음력 표시용.
    bool lunarOf(int serial, int& lYear, int& lMonth, int& lDay);
}
