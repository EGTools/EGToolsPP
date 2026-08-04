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
    // 내장 계산에 없는 항목을 병합한다(결정 3). apiKey가 비어 있으면 저장된
    // data.go.kr 키 사용; 인수로 오면 저장 후 사용(공공 API 키 정책과 동일).
    // 반환값은 추가된 항목 수. error가 있으면 실패 사유를 담는다(키 없음/거부/
    // 네트워크). 성공만 연도별 세션 캐시 — 실패는 캐시하지 않고 매번 재조회해
    // 사용자가 키를 고치면 즉시 반영된다(사용자 지시 2026-08-04).
    int mergeApiHolidays(int year, std::map<int, std::wstring>& list,
                         const std::wstring& apiKey = L"",
                         std::wstring* error = nullptr);

    // data.go.kr 인증키 — 서비스 구분 없이 계정당 하나(단일 슬롯 "datago").
    // 특일정보(KOREANHOLIDAYS·달력)와 사업자등록(BRNSTATUS)이 공유한다.
    // dataGoKey()는 구 "odcloud" 슬롯 값을 발견하면 datago로 자동 이관한다.
    std::wstring dataGoKey();
    void saveDataGoKey(const std::wstring& key);

    // serial(양력)의 음력 표기(년/월/일). 범위 밖이면 false.
    // 달력/일정표의 음력 표시용.
    bool lunarOf(int serial, int& lYear, int& lMonth, int& lDay);
}
