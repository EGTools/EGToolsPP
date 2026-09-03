# VISIBLEAGGR

**종류**: EGTools 전용 함수

범위에서 보이는 셀만 나열하거나 집계합니다.

## 구문

```
=VISIBLEAGGR(범위, [함수], [옵션])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 범위 | 필수 | 계산할 셀 범위 |
| 함수 | 선택 | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S·P/VAR.S·P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC), 생략하면 보이는 셀 나열 |
| 옵션 | 선택 | TEXTJOIN 구분자, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 반환

함수를 생략하면 보이는 셀만 2차원 배열로 스필하고, 함수를 지정하면 스칼라 집계값을 반환합니다. 보이는 셀이 없으면 #N/A, 보이는 셀에 오류가 있으면 COUNT/COUNTA 외에는 그 오류를 그대로 반환하며, 지원하지 않는 함수명은 #VALUE!입니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=VISIBLEAGGR(A1:B10)` |  | 보이는 셀만 나열(숨김 상태 의존) |
| `=VISIBLEAGGR(A1:A10,"SUM")` |  | 보이는 셀 합계(숨김 상태 의존) |
| `=VISIBLEAGGR(A1:A10,"LARGE",2)` |  | 보이는 값 중 2번째 큰 값 |

## 참고

- 지원 집계자(21종): SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE(.INC/.EXC)/QUARTILE(.INC/.EXC). 그 외는 #VALUE!.
- 옵션 인수: TEXTJOIN=구분자, LARGE/SMALL=k, PERCENTILE 계열=p(0~1), QUARTILE 계열=0~4.
- 행/열 숨김이 변경되어도 자동 재계산되지 않으므로 F9로 재계산이 필요합니다. 매크로형이라 멀티스레드 재계산에서 제외됩니다.
- 지원: Excel 2010+. 모든 Excel 버전에서 `VISIBLEAGGR` 이름 그대로 등록됩니다.
