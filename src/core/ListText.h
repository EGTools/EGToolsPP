// ListText.h — "셀 내 목록/배열"(M365 베타 Lists in Cells)의 텍스트 표기
// 파서·직렬화 (plan/27). 구버전 Excel은 목록 셀을 읽지 못하므로(plan/26)
// 변환 명령이 저장 파일의 richData를 읽어 이 표기로 바꾸고, HAS/HASANY/
// HASALL/FLATTEN 호환 UDF가 이 표기를 평가한다.
//
// 표기 규약(기존 ARRAYTOTEXT 두 모드와 정렬):
//   * 1열 목록(중첩 없음)  → concise: `Mon, Tue, Wed` (`, ` 연결, 따옴표 없음;
//     `,` `;` `{` `}` `"`·앞뒤 공백을 품거나 숫자/논리/오류로 읽힐 텍스트만
//     `"…"` 인용 — concise에서 유일하게 벗어나는 점, 역파싱용).
//   * 2D·중첩 배열          → strict: `{"a",1;"b",2}`, `{"a",{"b","c"},"d"}`
//     (텍스트는 모두 인용, 공백 없음, `;` 행·`,` 열, 중첩은 중괄호 겹침 =
//     네이티브 중첩 배열상수 `={{1,2,3};{4,5,6}}` 문법).
// 파서는 두 모드를 모두 받는다. 중괄호 없는 텍스트는 `,`로 나눈 N×1 목록,
// `{…}`는 행우선 2D(들쭉날쭉 행은 #N/A로 채움). 항목의 숫자 모양은 숫자,
// TRUE/FALSE는 논리, 8종 오류 이름은 오류, 그 외는 텍스트. 단, 콤마·중괄호가
// 없는 단일 텍스트는 분류하지 않고 텍스트 그대로 둔다(일반 셀 "1"은 텍스트;
// 항목 하나짜리 숫자 목록은 변환 명령이 숫자 값 자체로 쓴다). 파싱 불능
// (따옴표 미닫힘 등)은 오류를 내지 않고 전체를 항목 1개 텍스트로 취급한다.
#pragma once
#include <string>
#include <vector>

#include <xlOil/ExcelObj.h>

namespace egtools::core
{
    struct ListNode
    {
        bool isArray = false;
        xloil::ExcelObj value;          // !isArray 일 때의 스칼라
        size_t rows = 0, cols = 0;      // isArray: rows*cols == items.size()
        std::vector<ListNode> items;    // 행우선

        static ListNode scalar(xloil::ExcelObj v);
        static ListNode array(size_t rows, size_t cols, std::vector<ListNode> items);

        const ListNode& at(size_t r, size_t c) const { return items[r * cols + c]; }
        bool hasNested() const;         // 직계 자식 중 배열이 있는가
    };

    // 텍스트 → 트리. 항상 배열 노드(최소 1×1)를 돌려준다.
    ListNode parseListText(const std::wstring& text);

    // 셀 값 → 트리. Str은 parseListText, Multi(범위/배열)는 각 원소를 재귀
    // 변환한 외곽 배열, 그 외 스칼라(숫자·논리·오류·빈 셀)는 스칼라 노드.
    ListNode listFromValue(const xloil::ExcelObj& v);

    // 트리 → 텍스트(위 규약). 스칼라 노드는 셀 표시값 그대로.
    std::wstring formatListText(const ListNode& n);

    // 토큰 한 개 → 스칼라: TRUE/FALSE·8종 오류 이름·숫자 모양은 그 값, 그 외 텍스트
    // (인용 해제는 하지 않음 — 파일 richData의 항목처럼 이미 날것인 텍스트용).
    xloil::ExcelObj scalarFromToken(const std::wstring& token);

    // 스칼라 한 개의 표기. strict=true면 텍스트를 `"…"`(내부 `"`→`""`)로 인용
    // (VALUETOTEXT/ARRAYTOTEXT의 strict와 동일), 빈 값은 빈 문자열.
    std::wstring scalarText(const xloil::ExcelObj& v, bool strict);

    // 목록 항목 비교(HAS 계열): 텍스트는 대소문자 무시 전체 일치, 숫자는 수치
    // 동등, 논리는 논리끼리. 텍스트 "1"과 숫자 1은 다르다(Excel `=` 규칙).
    // 빈 값은 빈 텍스트와 같다. 오류는 어떤 값과도 같지 않다.
    bool listItemEquals(const xloil::ExcelObj& a, const xloil::ExcelObj& b);
}
