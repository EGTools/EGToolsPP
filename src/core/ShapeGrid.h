// ShapeGrid.h — 셀 사각형을 "도형 좌표계"로 구하는 헬퍼 (그림 삽입/맞춤 공용).
//
// 문제: Range.Top/Height는 현재 화면 DPI의 픽셀 반올림 값(예: 55.05pt 행 →
// 120DPI 91px = 54.6pt)이고, Shape.Top이 쓰는 도형 좌표계는 다른 반올림
// (같은 행이 54.75pt)을 쓴다. 그래서 `shape.Top = range.Top`은 행이 내려갈수록
// 누적 오차로 그림이 위로 밀린다(실측: 100행에서 ≈15pt → 한 행 위 셀에 놓임).
// 행 높이가 픽셀 격자에 맞는(0.75pt 배수 등) 파일에서는 드러나지 않는다.
//
// 해법: Shape.TopLeftCell은 도형 좌표계로 판정되므로, 도형을 임시로 옮겨 가며
// `TopLeftCell.Row >= r`이 참이 되는 최소 Top(= r행 상단의 도형 좌표)을 이진
// 탐색한다. 열도 같은 방식. 결과는 인스턴스 수명 동안(명령 1회) 캐시한다.
// 리본 그림 명령(CmdPicture.cpp)과 IMAGE/바코드/ADDRESSMAP(FxImage.cpp)이 공용.
#pragma once
#include "LateCom.h"

#include <algorithm>
#include <map>

namespace egtools::latecom
{
    class ShapeGrid
    {
        std::map<long, double> rowTop_, colLeft_;

        // probe를 임시로 옮겨 가며 `TopLeftCell.{Row|Column} >= idx`가 참이
        // 되는 최소 좌표를 찾는다. guess = Range 좌표계 추정값(탐색 시작점).
        static double search(IDispatch* probe, const wchar_t* prop,
                             const wchar_t* tlcProp, long idx, double guess)
        {
            if (idx <= 1) return 0.0;
            auto atLeast = [&](double v)
            {
                putDouble(probe, prop, v);
                IDispatch* tlc = getObject(probe, L"TopLeftCell");
                if (!tlc) return true;      // 판정 불가 → 괄호 실패 → guess 반환
                Releaser r{ tlc };
                return getLong(tlc, tlcProp, idx) >= idx;
            };
            // 괄호 구간: guess ± d, 실패 시 d를 4배씩 확장(8 → 2048pt).
            double lo = 0, hi = 0;
            bool ok = false;
            for (double d = 8.0; d <= 2048.0 && !ok; d *= 4)
            {
                lo = std::max(0.0, guess - d);
                hi = guess + d;
                ok = !atLeast(lo) && atLeast(hi);
            }
            if (!ok) return guess;          // 비정상(숨김 행 등) → Range 값 사용
            while (hi - lo > 0.02)
            {
                const double mid = (lo + hi) / 2;
                if (atLeast(mid)) hi = mid; else lo = mid;
            }
            return hi;
        }

    public:
        double rowTop(IDispatch* probe, long row, double guess)
        {
            auto it = rowTop_.find(row);
            if (it != rowTop_.end()) return it->second;
            const double v = search(probe, L"Top", L"Row", row, guess);
            rowTop_[row] = v;
            return v;
        }
        double colLeft(IDispatch* probe, long col, double guess)
        {
            auto it = colLeft_.find(col);
            if (it != colLeft_.end()) return it->second;
            const double v = search(probe, L"Left", L"Column", col, guess);
            colLeft_[col] = v;
            return v;
        }

        // range(셀/병합 영역/선택 영역)의 도형 좌표계 사각형. probe는 탐색 중
        // 임시로 이동되는 도형(보통 맞춤 대상 그림 자신).
        bool rect(IDispatch* probe, IDispatch* range,
                  double& l, double& t, double& w, double& h)
        {
            const long r1 = getLong(range, L"Row", 0);
            const long c1 = getLong(range, L"Column", 0);
            if (r1 < 1 || c1 < 1) return false;
            long nr = 1, nc = 1;
            if (IDispatch* rows = getObject(range, L"Rows"))
            {
                Releaser r{ rows }; nr = std::max(1L, getLong(rows, L"Count", 1));
            }
            if (IDispatch* cols = getObject(range, L"Columns"))
            {
                Releaser r{ cols }; nc = std::max(1L, getLong(cols, L"Count", 1));
            }
            const double gT = getDouble(range, L"Top");
            const double gL = getDouble(range, L"Left");
            const double gH = getDouble(range, L"Height");
            const double gW = getDouble(range, L"Width");
            constexpr long kMaxRows = 1048576, kMaxCols = 16384;

            t = rowTop(probe, r1, gT);
            l = colLeft(probe, c1, gL);
            h = (r1 + nr <= kMaxRows) ? rowTop(probe, r1 + nr, gT + gH) - t : gH;
            w = (c1 + nc <= kMaxCols) ? colLeft(probe, c1 + nc, gL + gW) - l : gW;
            if (h <= 0) h = gH;
            if (w <= 0) w = gW;
            return true;
        }
    };

    // 셰이프를 range에 여백 off로 맞춘다(LockAspectRatio 해제). 회전 90/270 보정
    // 포함(VB C03_Picture.vb:304-321과 동일 수식, 좌표만 도형 좌표계로 치환).
    //
    // keepAspect=true: 도형의 현재 가로세로 비율을 유지한 채 셀 안에 최대 크기로
    // 넣고, 남는 방향은 셀 중앙에 둔다(2D 바코드·IMAGE sizing 0). 호출 시점의
    // Width/Height가 원본 비율이어야 하므로 AddPicture는 Width/Height=-1(원본
    // 크기)로 삽입해 둘 것 — 셀 크기로 먼저 늘린 뒤 LockAspectRatio를 켜면
    // 찌그러진 비율이 그대로 고정된다(QR/DataMatrix 모듈이 직사각형이 되던 원인).
    inline void fitShapeToRange(ShapeGrid& grid, IDispatch* shape, IDispatch* range,
                                double off, bool keepAspect = false)
    {
        // TopLeftCell/BottomRightCell은 회전된 "시각적" 사각형으로 판정되지만
        // Top/Left/Width/Height는 회전 전 프레임 기준(실측: 90° 회전 시 TLC가
        // (W-H)/2만큼 어긋남). 경계 탐색과 배치는 회전을 0으로 풀고 하며, 마지막에
        // 복원한다(회전은 중심 기준이라 프레임을 셀 중심에 맞추면 시각적 사각형이
        // 셀과 일치).
        const double rot = getDouble(shape, L"Rotation");
        if (rot != 0.0) putDouble(shape, L"Rotation", 0.0);
        double maL, maT, maW, maH;
        if (!grid.rect(shape, range, maL, maT, maW, maH))
        {
            if (rot != 0.0) putDouble(shape, L"Rotation", rot);
            return;
        }

        // Excel은 속성 put마다 앵커(셀+EMU 오프셋)를 다시 계산하며 화면 배율에
        // 따라 put당 ≈0.15pt(125%: 96↔120DPI 픽셀 차)씩 양의 편향이 붙는다(실측).
        // 크기→위치 순으로 put하고, 위치 put 뒤에 남는 편향으로 아래/오른쪽
        // 모서리가 셀 경계에 닿아 BottomRightCell이 다음 행/열로 잡히지 않도록
        // 크기에 kDrift 여유를 둔다(0.45pt ≈ 0.6px, 육안 식별 불가).
        constexpr double kDrift = 0.45;
        const double boxW = std::max(1.0, maW - off * 2 - kDrift);
        const double boxH = std::max(1.0, maH - off * 2 - kDrift);
        const bool swap = !(rot == 0.0 || rot == 180.0);   // 90/270도: 시각적 가로세로 교환

        // 시각적(회전 후) 크기 vw×vh를 정한 뒤 프레임(회전 전) 크기로 되돌린다.
        double vw = boxW, vh = boxH;
        if (keepAspect)
        {
            const double w0 = getDouble(shape, L"Width");
            const double h0 = getDouble(shape, L"Height");
            if (w0 > 0.0 && h0 > 0.0)
            {
                const double sw = swap ? h0 : w0, sh = swap ? w0 : h0;
                const double s = std::min(boxW / sw, boxH / sh);
                vw = std::max(1.0, sw * s);
                vh = std::max(1.0, sh * s);
            }
        }
        const double fw = swap ? vh : vw, fh = swap ? vw : vh;

        // 회전은 프레임 중심 기준이므로 프레임 중심을 박스 중심에 두면 시각적
        // 사각형이 셀과 일치한다(채움: 박스와 동일, 비율 유지: 셀 중앙 정렬).
        // (회전 0·채움일 때 Top=maT+off, Left=maL+off로 기존 수식과 동일.)
        putLong(shape, L"LockAspectRatio", 0);
        putDouble(shape, L"Width", fw);
        putDouble(shape, L"Height", fh);
        putDouble(shape, L"Top", maT + off + (boxH - fh) / 2);
        putDouble(shape, L"Left", maL + off + (boxW - fw) / 2);
        if (rot != 0.0) putDouble(shape, L"Rotation", rot);
    }
}
