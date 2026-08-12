// FxEgQC.cpp — 품질관리 함수 (EGToolsVB FX6_QC.vb 포팅, plan/20 E3).
//   SAMPLINGLABEL — ISO 2859-1 LOT 크기·검사수준 → 시료 문자.
//   SAMPLINGSIZE  — 시료수(n).   SAMPLINGAC — 합격판정개수(Ac).   SAMPLINGRE — 불합격판정개수(Re).
//   CP / CPK      — 공정능력지수.
// ISO 2859-1:1999 표를 토큰 문자열로 내장: N=아래 화살표, P=위 화살표,
// *=1회 샘플링으로 전환, ++=2회 샘플링으로 전환, #=해당 없음.
// 2회/다회 표의 복합 토큰("a/b;c/d;…")은 차수(Round)로 부분을, Ac/Re로 쌍을 고른다.
// 전부 순수 계산(비매크로형).

#include "../core/Registry.h"
#include "../core/Apply.h"

#include <xlOil/xlOil.h>
#include <xlOil/ExcelArray.h>

#include <algorithm>
#include <cmath>
#include <cwctype>
#include <map>
#include <string>
#include <vector>

using namespace xloil;

namespace egtools::functions
{
    namespace
    {
        bool isNumType(const ExcelObj& e)
        {
            return e.type() == ExcelType::Num || e.type() == ExcelType::Int;
        }

        // ---- 표 데이터 -----------------------------------------------------
        // 각 행 = 공백 구분 토큰 27개 (열0 = 시료수 n, 열1~26 = AQL 등급).

        // 1회 샘플링 — 까다로운(T)/보통(N)/수월한(R).
        static const wchar_t* kT1_T[16] = {
            L"2 N N N N N N N N N N N N N N N N N N 1 2 3 5 8 12 18 27",
            L"3 N N N N N N N N N N N N N N 0 N N 1 2 3 5 8 12 18 27 41",
            L"5 N N N N N N N N N N N N N 0 N N 1 2 3 5 8 12 18 27 41 P",
            L"8 N N N N N N N N N N N N 0 N N 1 2 3 5 8 12 18 27 41 P P",
            L"13 N N N N N N N N N N N 0 N N 1 2 3 5 8 12 18 27 41 P P P",
            L"20 N N N N N N N N N N 0 N N 1 2 3 5 8 12 18 P P P P P P",
            L"32 N N N N N N N N N 0 N N 1 2 3 5 8 12 18 P P P P P P P",
            L"50 N N N N N N N N 0 N N 1 2 3 5 8 12 18 P P P P P P P P",
            L"80 N N N N N N N 0 N N 1 2 3 5 8 12 18 P P P P P P P P P",
            L"125 N N N N N N 0 N N 1 2 3 5 8 12 18 P P P P P P P P P P",
            L"200 N N N N N 0 N N 1 2 3 5 8 12 18 P P P P P P P P P P P",
            L"315 N N N N 0 N N 1 2 3 5 8 12 18 P P P P P P P P P P P P",
            L"500 N N N 0 N N 1 2 3 5 8 12 18 P P P P P P P P P P P P P",
            L"800 N N 0 N N 1 2 3 5 8 12 18 P P P P P P P P P P P P P P",
            L"1250 N 0 N N 1 2 3 5 8 12 18 P P P P P P P P P P P P P P P",
            L"2000 0 P N 1 2 3 5 8 12 18 P P P P P P P P P P P P P P P P" };
        static const wchar_t* kT1_N[16] = {
            L"2 N N N N N N N N N N N N N N 0 N N 1 2 3 5 7 10 14 21 30",
            L"3 N N N N N N N N N N N N N 0 P N 1 2 3 5 7 10 14 21 30 44",
            L"5 N N N N N N N N N N N N 0 P N 1 2 3 5 7 10 14 21 30 44 P",
            L"8 N N N N N N N N N N N 0 P N 1 2 3 5 7 10 14 21 30 44 P P",
            L"13 N N N N N N N N N N 0 P N 1 2 3 5 7 10 14 21 30 44 P P P",
            L"20 N N N N N N N N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P",
            L"32 N N N N N N N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P",
            L"50 N N N N N N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P",
            L"80 N N N N N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P P",
            L"125 N N N N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P P P",
            L"200 N N N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P P P P",
            L"315 N N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P P P P P",
            L"500 N N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P P P P P P",
            L"800 N 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P P P P P P P",
            L"1250 0 P N 1 2 3 5 7 10 14 21 P P P P P P P P P P P P P P P",
            L"2000 P P 1 2 3 5 7 10 14 21 P P P P P P P P P P P P P P P P" };
        static const wchar_t* kT1_R[16] = {
            L"2 N N N N N N N N N N N N N N 0 N N 1 2 3 5 7 10 14 21 30",
            L"2 N N N N N N N N N N N N N 0 P N N 1 2 3 5 7 10 14 21 30",
            L"2 N N N N N N N N N N N N 0 P N N 1 2 3 4 6 8 10 14 21 P",
            L"3 N N N N N N N N N N N 0 P N N 1 2 3 4 6 8 10 14 21 P P",
            L"5 N N N N N N N N N N 0 P N N 1 2 3 4 6 8 10 14 21 P P P",
            L"8 N N N N N N N N N 0 P N N 1 2 3 4 6 8 10 P P P P P P",
            L"13 N N N N N N N N 0 P N N 1 2 3 4 6 8 10 P P P P P P P",
            L"20 N N N N N N N 0 P N N 1 2 3 4 6 8 10 P P P P P P P P",
            L"32 N N N N N N 0 P N N 1 2 3 4 6 8 10 P P P P P P P P P",
            L"50 N N N N N 0 P N N 1 2 3 4 6 8 10 P P P P P P P P P P",
            L"80 N N N N 0 P N N 1 2 3 4 6 8 10 P P P P P P P P P P P",
            L"125 N N N 0 P N N 1 2 3 4 6 8 10 P P P P P P P P P P P P",
            L"200 N N 0 P N N 1 2 3 4 6 8 10 P P P P P P P P P P P P P",
            L"315 N 0 P N N 1 2 3 4 6 8 10 P P P P P P P P P P P P P P",
            L"500 0 P P N 1 2 3 4 6 8 10 P P P P P P P P P P P P P P P",
            L"800 P P P 1 2 3 4 6 8 10 P P P P P P P P P P P P P P P P" };

        // 2회 샘플링 — 복합 토큰은 이름으로 두고 파싱 시 치환.
        static const wchar_t* kT2_T[16] = {
            L"0 N N N N N N N N N N N N N N N N N N * * * * * * * *",
            L"2 N N N N N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 T8 T9",
            L"3 N N N N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 T8 T9 P",
            L"5 N N N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 T8 T9 P P",
            L"8 N N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 T8 T9 P P P",
            L"13 N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P",
            L"20 N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P",
            L"32 N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P",
            L"50 N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P",
            L"80 N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P",
            L"125 N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P",
            L"200 N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P",
            L"315 N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P",
            L"500 N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P P",
            L"800 N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P P P",
            L"1250 * P N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P P P P" };
        static const wchar_t* kT2_N[16] = {
            L"0 N N N N N N N N N N N N N N * N N * * * * * * * * *",
            L"2 N N N N N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 N9 N0",
            L"3 N N N N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 N9 N0 P",
            L"5 N N N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 N9 N0 P P",
            L"8 N N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 N9 N0 P P P",
            L"13 N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P",
            L"20 N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P",
            L"32 N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P",
            L"50 N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P",
            L"80 N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P",
            L"125 N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P",
            L"200 N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P",
            L"315 N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P",
            L"500 N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P P",
            L"800 * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P P P",
            L"1250 P P N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P P P P" };
        static const wchar_t* kT2_R[16] = {
            L"0 N N N N N N N N N N N N N N * N N * * * * * * * * *",
            L"0 N N N N N N N N N N N N N * P N N * * * * * * * * *",
            L"0 N N N N N N N N N N N N * P N N * * * * * * * * * P",
            L"2 N N N N N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 R8 R9 P P",
            L"3 N N N N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 R8 R9 P P P",
            L"5 N N N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P",
            L"8 N N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P",
            L"13 N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P",
            L"20 N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P",
            L"32 N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P",
            L"50 N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P",
            L"80 N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P",
            L"125 N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P",
            L"200 N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P P",
            L"315 * P P N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P P P",
            L"500 P P P R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P P P P" };

        // 다회(3~5회) 샘플링.
        static const wchar_t* kTM_T[16] = {
            L"0 N N N N N N N N N N N N N N N N N N * * * * * * * *",
            L"0 N N N N N N N N N N N N N N * N N ++ ++ ++ ++ ++ ++ ++ ++ ++",
            L"0 N N N N N N N N N N N N N * N N ++ ++ ++ ++ ++ ++ ++ ++ ++ P",
            L"2 N N N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 T8 T9 P P",
            L"3 N N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 T8 T9 P P P",
            L"5 N N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P",
            L"8 N N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P",
            L"13 N N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P",
            L"20 N N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P",
            L"32 N N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P",
            L"50 N N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P",
            L"80 N N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P",
            L"125 N N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P",
            L"200 N N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P P",
            L"315 N * N N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P P P",
            L"500 * P N T1 T2 T3 T4 T5 T6 T7 P P P P P P P P P P P P P P P P" };
        static const wchar_t* kTM_N[16] = {
            L"0 N N N N N N N N N N N N N N * N N * * * * * * * * *",
            L"0 N N N N N N N N N N N N N * P N N ++ ++ ++ ++ ++ ++ ++ ++ ++",
            L"0 N N N N N N N N N N N N * P N N ++ ++ ++ ++ ++ ++ ++ ++ ++ P",
            L"2 N N N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 N9 N0 P P",
            L"3 N N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 N9 N0 P P P",
            L"5 N N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P",
            L"8 N N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P",
            L"13 N N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P",
            L"20 N N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P",
            L"32 N N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P",
            L"50 N N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P",
            L"80 N N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P",
            L"125 N N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P",
            L"200 N * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P P",
            L"315 * P N N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P P P",
            L"500 P P N1 N2 N3 N4 N5 N6 N7 N8 P P P P P P P P P P P P P P P P" };
        static const wchar_t* kTM_R[16] = {
            L"0 N N N N N N N N N N N N N N * N N * * * * * * * * *",
            L"0 N N N N N N N N N N N N N * P N N * * * * * * * * *",
            L"0 N N N N N N N N N N N N * P N N * * * * * * * * * P",
            L"0 N N N N N N N N N N N * P N N ++ ++ ++ ++ ++ ++ ++ ++ ++ P P",
            L"0 N N N N N N N N N N * P N N ++ ++ ++ ++ ++ ++ ++ ++ ++ P P P",
            L"5 N N N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P",
            L"8 N N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P",
            L"13 N N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P",
            L"20 N N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P",
            L"32 N N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P",
            L"50 N N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P",
            L"80 N N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P",
            L"125 N N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P",
            L"200 N * P N N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P P",
            L"315 * P P N R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P P P",
            L"500 P P P R1 R2 R3 R4 R5 R6 R7 P P P P P P P P P P P P P P P P" };

        // 복합 토큰 값 (2회 / 다회).
        const std::map<std::wstring, std::wstring>& tokens2()
        {
            static const std::map<std::wstring, std::wstring> m = {
                { L"T1", L"0/2;1/2" }, { L"T2", L"0/3;3/4" }, { L"T3", L"1/3;4/5" },
                { L"T4", L"2/5;6/7" }, { L"T5", L"4/7;10/11" }, { L"T6", L"6/10;15/16" },
                { L"T7", L"9/14;23/44" }, { L"T8", L"15/20;34/35" }, { L"T9", L"23/29;52/53" },
                { L"N1", L"0/2;1/2" }, { L"N2", L"0/3;3/4" }, { L"N3", L"1/3;4/5" },
                { L"N4", L"2/5;6/7" }, { L"N5", L"3/6;9/10" }, { L"N6", L"5/9;12/13" },
                { L"N7", L"7/11;18/19" }, { L"N8", L"11/16;26/27" }, { L"N9", L"17/22;37/38" },
                { L"N0", L"25/31;56/57" },
                { L"R1", L"0/2;1/2" }, { L"R2", L"0/3;3/4" }, { L"R3", L"1/3;4/5" },
                { L"R4", L"2/4;5/6" }, { L"R5", L"3/6;7/8" }, { L"R6", L"4/7;10/11" },
                { L"R7", L"5/9;12/13" }, { L"R8", L"7/11;18/19" }, { L"R9", L"11/16;26/27" } };
            return m;
        }
        const std::map<std::wstring, std::wstring>& tokensM()
        {
            static const std::map<std::wstring, std::wstring> m = {
                { L"T1", L"#/2;0/2;0/2;0/2;1/2" }, { L"T2", L"#/2;0/3;0/3;1/3;3/4" },
                { L"T3", L"#/3;0/3;1/4;2/5;4/5" }, { L"T4", L"#/4;1/5;2/6;4/7;6/7" },
                { L"T5", L"0/4;2/7;4/9;6/11;10/11" }, { L"T6", L"0/6;3/9;7/12;11/15;15/16" },
                { L"T7", L"1/8;6/12;11/17;16/22;23/24" }, { L"T8", L"3/10;10/17;17/24;25/31;34/35" },
                { L"T9", L"6/15;16/25;26/35;38/45;52/53" },
                { L"N1", L"#/2;0/2;0/2;0/2;1/2" }, { L"N2", L"#/2;0/3;0/3;1/3;3/4" },
                { L"N3", L"#/3;0/3;1/4;2/5;4/5" }, { L"N4", L"#/4;1/5;2/6;4/7;6/7" },
                { L"N5", L"0/4;1/6;3/8;5/9;9/10" }, { L"N6", L"0/5;3/8;6/10;9/12;12/13" },
                { L"N7", L"1/7;4/10;8/13;12/17;18/19" }, { L"N8", L"2/9;7/14;13/19;20/25;26/27" },
                { L"N9", L"4/12;11/19;19/27;28/34;37/38" }, { L"N0", L"6/16;17/27;29/38;40/48;56/57" },
                { L"R1", L"#/2;0/2;0/2;0/2;1/2" }, { L"R2", L"#/2;0/3;0/3;1/3;3/4" },
                { L"R3", L"#/3;0/3;1/4;2/5;4/5" }, { L"R4", L"#/3;1/4;2/5;3/5;5/6" },
                { L"R5", L"0/4;1/6;2/7;4/8;7/8" }, { L"R6", L"0/4;2/7;4/9;6/11;10/11" },
                { L"R7", L"0/5;3/8;6/10;9/12;12/19" } };
            return m;
        }

        // ---- 표 셀 ---------------------------------------------------------

        struct Cell
        {
            enum Kind { Num, Down, Up, Star, Dbl, Hash } kind = Down;
            double val = 0;
        };

        bool numericToken(const std::wstring& t, double& v)
        {
            if (t.empty()) return false;
            wchar_t* end = nullptr;
            v = wcstod(t.c_str(), &end);
            return end && *end == L'\0';
        }

        // (plan군, 검사 엄격도, 차수, Ac/Re) → 셀 그리드. 실패 시 빈 그리드.
        std::vector<std::vector<Cell>> makeGrid(int planKind /*1,2,3*/, int insp,
                                                int round, int acRe)
        {
            const wchar_t** rows = nullptr;
            const std::map<std::wstring, std::wstring>* tok = nullptr;
            if (planKind == 1)
                rows = insp > 0 ? kT1_T : (insp == 0 ? kT1_N : kT1_R);
            else if (planKind == 2)
            {
                rows = insp > 0 ? kT2_T : (insp == 0 ? kT2_N : kT2_R);
                tok = &tokens2();
            }
            else
            {
                rows = insp > 0 ? kTM_T : (insp == 0 ? kTM_N : kTM_R);
                tok = &tokensM();
            }

            std::vector<std::vector<Cell>> grid;
            for (int r = 0; r < 16; ++r)
            {
                std::vector<Cell> row;
                std::wstring s = rows[r], cur;
                std::vector<std::wstring> parts;
                for (wchar_t c : s + L" ")
                {
                    if (c == L' ') { if (!cur.empty()) parts.push_back(cur); cur.clear(); }
                    else cur.push_back(c);
                }
                if (parts.size() != 27) return {};   // 전사 오류 방어

                for (size_t c = 0; c < parts.size(); ++c)
                {
                    std::wstring t = parts[c];
                    if (tok)
                    {
                        const auto it = tok->find(t);
                        if (it != tok->end()) t = it->second;
                    }
                    Cell cell;
                    double v;
                    if (t.find(L';') != std::wstring::npos)
                    {
                        // "a/b;c/d;…" — 차수로 부분, Ac/Re로 쌍 선택.
                        std::vector<std::wstring> ps;
                        std::wstring pc;
                        for (wchar_t ch : t + L";")
                        {
                            if (ch == L';') { ps.push_back(pc); pc.clear(); }
                            else pc.push_back(ch);
                        }
                        if (round < 1 || (size_t)round > ps.size()) return {};
                        const std::wstring& pair = ps[(size_t)round - 1];
                        const size_t slash = pair.find(L'/');
                        const std::wstring pick = (acRe == 0)
                            ? pair.substr(0, slash) : pair.substr(slash + 1);
                        if (pick == L"#") cell.kind = Cell::Hash;
                        else if (numericToken(pick, v)) { cell.kind = Cell::Num; cell.val = v; }
                        else return {};
                    }
                    else if (t == L"N") cell.kind = Cell::Down;
                    else if (t == L"P") cell.kind = Cell::Up;
                    else if (t == L"*") cell.kind = Cell::Star;
                    else if (t == L"++") cell.kind = Cell::Dbl;
                    else if (t == L"#") cell.kind = Cell::Hash;
                    else if (numericToken(t, v))
                    {
                        cell.kind = Cell::Num;
                        // 1회 표의 Re = Ac + 1 (열0 = n은 제외).
                        cell.val = (planKind == 1 && c > 0 && acRe == 1) ? v + 1 : v;
                    }
                    else return {};
                    row.push_back(cell);
                }
                grid.push_back(std::move(row));
            }

            // 까다로운 검사 확장 행.
            if (insp > 0)
            {
                std::vector<Cell> extra(27);   // 기본 Down
                auto num = [](double v) { Cell c; c.kind = Cell::Num; c.val = v; return c; };
                if (planKind == 1) { extra[0] = num(3150); extra[3] = num(1 + (acRe == 1 ? 1 : 0)); }
                else if (planKind == 2) { extra[0] = num(2000); extra[3] = num(round == 1 ? 0 : 1); }
                else
                {
                    extra[0] = num(800);
                    if (round == 1) { Cell c; c.kind = Cell::Hash; extra[3] = c; }
                    else extra[3] = num(round == 5 ? 1 : 0);
                }
                grid.push_back(std::move(extra));
            }
            return grid;
        }

        // ---- 샘플링 계산 ---------------------------------------------------

        int normalizeInspection(const ExcelObj& v)
        {
            if (v.isMissing()) return 0;
            if (v.type() == ExcelType::Str)
            {
                const std::wstring s = v.toString();
                auto has = [&](wchar_t c)
                {
                    for (wchar_t x : s)
                        if (std::towupper(x) == (wchar_t)std::towupper(c) || x == c) return true;
                    return false;
                };
                if (has(L'N') || s.find(L'보') != std::wstring::npos) return 0;
                if (has(L'T') || s.find(L'까') != std::wstring::npos) return 1;
                if (has(L'R') || s.find(L'수') != std::wstring::npos) return -1;
                return 0;
            }
            const double d = v.get<double>(0.0);
            return d > 0 ? 1 : (d < 0 ? -1 : 0);
        }

        std::wstring normalizeSampleType(const ExcelObj& v)
        {
            std::wstring s = v.isMissing() ? L"" : v.toString();
            std::wstring t;
            for (wchar_t c : s)
                if (c != L'-' && c != L' ') t.push_back((wchar_t)std::towupper(c));
            if (t.empty()) t = L"G2";
            return t;
        }

        // 시료 문자 (실패 시 빈 문자열 + err 설정).
        std::wstring samplingLabel(long lotSize, const std::wstring& sType, CellError& err)
        {
            static const long kSizes[15] = { 2, 9, 16, 26, 51, 91, 151, 281, 501,
                                             1201, 3201, 10001, 35001, 150001, 500001 };
            static const std::map<std::wstring, std::wstring> kDict = {
                { L"S1", L"AAAABBBBCCCCDDD" }, { L"S2", L"AAABBBCCCDDDEEE" },
                { L"S3", L"AABBCCDDEEFFGGH" }, { L"S4", L"AABCCDEEFGGHJJK" },
                { L"G1", L"AABCCDEFGHJKLMN" }, { L"G2", L"ABCDEFGHJKLMNPQ" },
                { L"G3", L"BCDEFGHJKLMNPQR" } };
            if (lotSize < 2) { err = CellError::Num; return L""; }
            const auto it = kDict.find(sType);
            if (it == kDict.end()) { err = CellError::Value; return L""; }
            int idx = 0;
            for (int i = 0; i < 15; ++i)
                if (lotSize < kSizes[i]) { idx = i; break; }
            if (idx == 0) idx = 15;
            return std::wstring(1, it->second[(size_t)idx - 1]);
        }

        // acReSize: 0=n, 1=Ac, 2=Re.
        ExcelObj aqlValue(long lotSize, double aql, const std::wstring& sType,
                          int insp, int acReSize, int plan, int round, int depth)
        {
            if (depth > 4) return ExcelObj(CellError::Value);

            bool force = false;
            if (plan > 10) { force = true; plan = plan % 10; }
            if (plan < 1 || plan > 5) return ExcelObj(CellError::Value);
            if (plan == 1 && round != 1) return ExcelObj(CellError::Value);
            if (plan == 2 && (round < 1 || round > 2)) return ExcelObj(CellError::Value);
            if (plan > 2 && (round < 1 || round > 5)) return ExcelObj(CellError::Value);

            CellError err = CellError::Value;
            const std::wstring label = samplingLabel(lotSize, sType, err);
            if (label.empty()) return ExcelObj(err);
            static const std::wstring kLabels = L"ABCDEFGHJKLMNPQRS";
            int L = (int)kLabels.find(label[0]);
            if (L < 0) return ExcelObj(CellError::Value);

            static const double kAQL[26] = { 0.01, 0.015, 0.025, 0.04, 0.065, 0.1,
                0.15, 0.25, 0.4, 0.65, 1, 1.5, 2.5, 4, 6.5, 10, 15, 25, 40, 65,
                100, 150, 250, 400, 650, 1000 };
            int idx = -1;
            for (int i = 0; i < 26; ++i)
                if (aql < kAQL[i]) { idx = i; break; }
            if (idx == -1) idx = 26;
            if (idx == 0) return ExcelObj(CellError::Num);   // AQL < 0.01

            const int planKind = plan == 1 ? 1 : (plan == 2 ? 2 : 3);
            const int acRe = acReSize > 1 ? 1 : 0;
            const auto grid = makeGrid(planKind, insp, round, acRe);
            if (grid.empty()) return ExcelObj(CellError::Value);
            const int rows = (int)grid.size();

            bool on = false;
            for (int guard = 0; guard < 64; ++guard)
            {
                if (L < 0 || L >= rows) return ExcelObj(CellError::NA);
                const Cell cur = grid[(size_t)L][(size_t)idx];
                if (cur.kind == Cell::Num) break;

                if (force && (cur.kind == Cell::Star || cur.kind == Cell::Dbl)) on = true;
                if (on)
                    L += 1;
                else
                    L += (cur.kind == Cell::Up) ? -1 : 1;
                if (L < 0 || L >= rows) return ExcelObj(CellError::NA);

                const Cell next = grid[(size_t)L][(size_t)idx];
                if (next.kind == Cell::Hash) return ExcelObj(CellError::NA);
                if (!on && (next.kind == Cell::Star || next.kind == Cell::Dbl))
                {
                    if (next.kind == Cell::Star)
                        return aqlValue(lotSize, aql, sType, insp, acReSize, 1, 1, depth + 1);
                    return aqlValue(lotSize, aql, sType, insp, acReSize, 2,
                                    round > 1 ? 2 : 1, depth + 1);
                }
            }

            if (L < 0 || L >= rows || grid[(size_t)L][(size_t)idx].kind != Cell::Num)
                return ExcelObj(CellError::NA);

            if (acReSize == 0)
            {
                double n = grid[(size_t)L][0].val;
                if ((double)lotSize < n) n = (double)lotSize;
                return ExcelObj(n);
            }
            return ExcelObj(grid[(size_t)L][(size_t)idx].val);
        }

        // Scalar core — 전 인수 원소별 리프팅 대상(LOT 목록·AQL 표가 흔한 용법).
        ExcelObj samplingOne(const ExcelObj& lotObj, const ExcelObj& aqlObj,
                             const ExcelObj& typeObj, const ExcelObj& inspObj,
                             const ExcelObj& planObj, const ExcelObj& roundObj,
                             int acReSize)
        {
            const long lot = (long)(lotObj.isMissing() ? 0 : lotObj.get<double>(0.0));
            if (lot == 1)
                return ExcelObj((double)(acReSize == 1 ? 0 : 1));
            const double aql = aqlObj.isMissing() ? -1 : aqlObj.get<double>(-1);
            if (aql < 0 || aql > 1000) return ExcelObj(CellError::Value);
            int plan = planObj.isMissing() ? 0 : planObj.get<int>(0);
            if (plan == 0) plan = 1;
            if (plan < 1 || plan > 15) return ExcelObj(CellError::Value);
            int round = roundObj.isMissing() ? 0 : roundObj.get<int>(0);
            if (round == 0) round = 1;
            if (round < 1 || round > 5) return ExcelObj(CellError::Value);

            return aqlValue(lot, aql, normalizeSampleType(typeObj),
                            normalizeInspection(inspObj), acReSize,
                            plan, round, 0);
        }

        ExcelObj* sampling(const ExcelObj& lotObj, const ExcelObj& aqlObj,
                           const ExcelObj& typeObj, const ExcelObj& inspObj,
                           const ExcelObj& planObj, const ExcelObj& roundObj,
                           int acReSize)
        {
            return egtools::core::mapLift(
                [acReSize](const ExcelObj& le, const ExcelObj& ae, const ExcelObj& te,
                           const ExcelObj& ie, const ExcelObj& pe, const ExcelObj& re)
                { return samplingOne(le, ae, te, ie, pe, re, acReSize); },
                lotObj, aqlObj, typeObj, inspObj, planObj, roundObj);
        }

        // ---- CP / CPK ------------------------------------------------------

        bool collectStats(const ExcelObj& data, double& avg, double& std)
        {
            std::vector<double> vals;
            auto push = [&](const ExcelObj& e)
            {
                if (e.type() == ExcelType::Num || e.type() == ExcelType::Int)
                    vals.push_back(e.get<double>(0.0));
            };
            if (data.isType(ExcelType::Multi))
            {
                ExcelArray a(data);
                const size_t n = (size_t)a.nRows() * a.nCols();
                for (size_t k = 0; k < n; ++k) push(a.at(k));
            }
            else push(data);
            if (vals.size() < 2) return false;

            double sum = 0;
            for (double v : vals) sum += v;
            avg = sum / vals.size();
            double ss = 0;
            for (double v : vals) ss += (v - avg) * (v - avg);
            std = sqrt(ss / (vals.size() - 1));
            return true;
        }

        // 규격(USL/LSL)은 원소별 리프팅 대상(규격 세트 여러 개 → Cp/Cpk 배열).
        // 기존에는 배열 규격이 isNumType에 걸러져 조용히 무시되고 그럴듯한
        // 오답이 나왔다(plan/22 A-P0).
        ExcelObj* cpFn(const ExcelObj& data, const ExcelObj& uslObj, const ExcelObj& lslObj)
        {
            if (data.isMissing()) return returnValue(CellError::Value);
            double avg, sd;
            if (!collectStats(data, avg, sd)) return returnValue(CellError::Num);
            if (sd == 0) return returnValue(CellError::Div0);
            return egtools::core::mapLift(
                [avg, sd](const ExcelObj& ue, const ExcelObj& le) -> ExcelObj
                {
                    const bool hasU = !ue.isMissing() && isNumType(ue);
                    const bool hasL = !le.isMissing() && isNumType(le);
                    if (!hasU && !hasL) return ExcelObj(CellError::Value);
                    const double usl = hasU ? ue.get<double>(0.0) : avg;
                    const double lsl = hasL ? le.get<double>(0.0) : avg;
                    return ExcelObj((usl - lsl) / (6 * sd));
                }, uslObj, lslObj);
        }

        ExcelObj* cpkFn(const ExcelObj& data, const ExcelObj& uslObj, const ExcelObj& lslObj)
        {
            if (data.isMissing()) return returnValue(CellError::Value);
            double avg, sd;
            if (!collectStats(data, avg, sd)) return returnValue(CellError::Num);
            if (sd == 0) return returnValue(CellError::Div0);
            return egtools::core::mapLift(
                [avg, sd](const ExcelObj& ue, const ExcelObj& le) -> ExcelObj
                {
                    const bool hasU = !ue.isMissing() && isNumType(ue);
                    const bool hasL = !le.isMissing() && isNumType(le);
                    if (!hasU && !hasL) return ExcelObj(CellError::Value);
                    double cpk = 1e308;
                    if (hasU) cpk = (std::min)(cpk, (ue.get<double>(0.0) - avg) / (3 * sd));
                    if (hasL) cpk = (std::min)(cpk, (avg - le.get<double>(0.0)) / (3 * sd));
                    return ExcelObj(cpk);
                }, uslObj, lslObj);
        }
    }

    void registerEgQC()
    {
        egtools::core::registerFn(L"SAMPLINGLABEL",
            [](const ExcelObj& lot, const ExcelObj& type) -> ExcelObj*
            {
                return egtools::core::mapLift(
                    [](const ExcelObj& le, const ExcelObj& te) -> ExcelObj
                    {
                        const long n = (long)(le.isMissing() ? 0 : le.get<double>(0.0));
                        CellError err = CellError::Value;
                        const std::wstring s = samplingLabel(n, normalizeSampleType(te), err);
                        if (s.empty()) return ExcelObj(err);
                        return ExcelObj(std::wstring_view(s));
                    }, lot, type);
            });

        egtools::core::registerFn(L"SAMPLINGSIZE",
            [](const ExcelObj& lot, const ExcelObj& aql, const ExcelObj& type,
               const ExcelObj& insp, const ExcelObj& plan, const ExcelObj& round) -> ExcelObj*
            { return sampling(lot, aql, type, insp, plan, round, 0); });

        egtools::core::registerFn(L"SAMPLINGAC",
            [](const ExcelObj& lot, const ExcelObj& aql, const ExcelObj& type,
               const ExcelObj& insp, const ExcelObj& plan, const ExcelObj& round) -> ExcelObj*
            { return sampling(lot, aql, type, insp, plan, round, 1); });

        egtools::core::registerFn(L"SAMPLINGRE",
            [](const ExcelObj& lot, const ExcelObj& aql, const ExcelObj& type,
               const ExcelObj& insp, const ExcelObj& plan, const ExcelObj& round) -> ExcelObj*
            { return sampling(lot, aql, type, insp, plan, round, 2); });

        egtools::core::registerFn(L"CP",
            [](const ExcelObj& d, const ExcelObj& u, const ExcelObj& l) -> ExcelObj*
            { return cpFn(d, u, l); });

        egtools::core::registerFn(L"CPK",
            [](const ExcelObj& d, const ExcelObj& u, const ExcelObj& l) -> ExcelObj*
            { return cpkFn(d, u, l); });
    }
}
