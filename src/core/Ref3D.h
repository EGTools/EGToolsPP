// Ref3D.h — 3D 참조(Sheet1:Sheet3!A1:B2) 인수 복원.
//
// Excel은 3D 참조를 XLL 인수로 넘기지 못하고 #VALUE!를 전달한다. 네이티브에서
// 3D 참조를 받는 함수(VSTACK/HSTACK/TOCOL/TOROW/CONCAT/TEXTJOIN)를 위해, 호출
// 셀의 수식 텍스트(GET.CELL 41)에서 그 인수 텍스트를 찾아 시트별 값을 읽는다
// (VB판 방식). xlfCaller/GET.CELL을 쓰므로 함수는 macro=true로 등록해야 한다.
//
// 네이티브 의미: 각 시트의 범위를 따로 읽어 통합 문서의 시트 순서대로 잇는다
// (TOCOL의 열 우선 읽기도 시트별로 적용 — 2026-10 365 실측).

#pragma once
#include <string>
#include <vector>

#include <xlOil/ExcelObj.h>

namespace egtools::core
{
    class Ref3D
    {
    public:
        // bareName: 함수의 i18n 이름(수식 속 이름은 registeredName으로 결정).
        // nArgs: 전달된 인수 개수(끝의 생략 인수 제외) — 같은 함수가 수식에
        //        여러 번 있을 때 인수 개수가 맞는 호출을 고르는 기준.
        Ref3D(const wchar_t* bareName, size_t nArgs);

        // 인수 i(0-based)가 3D 참조면 시트 순서대로 각 시트의 범위 값을, 아니면
        // 빈 벡터. arg가 #VALUE!일 때만 수식을 읽는다(호출당 1회, 지연).
        std::vector<xloil::ExcelObj> sheets(size_t i, const xloil::ExcelObj& arg);

    private:
        std::wstring _name;
        size_t _n;
        bool _tried = false;
        xloil::ExcelObj _caller;
        std::vector<std::wstring> _texts;   // 고른 호출의 인수 텍스트
    };

    // 끝의 생략(Missing) 인수를 뺀 인수 개수.
    size_t passedArgs(const xloil::ExcelObj* const* args, size_t n);
}
