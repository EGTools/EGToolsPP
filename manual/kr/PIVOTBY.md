# PIVOTBY

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Microsoft 365

행·열 키로 데이터를 피벗해 값을 집계합니다. 집계자=텍스트(SUM, COUNT…).

## 구문

```
=PIVOTBY(행필드, 열필드, 값, 함수, [필드머리글], [행총계깊이], [행정렬순서], [열총계깊이], [열정렬순서], [필터배열], [기준])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 행필드 | 필수 | 행 키 열 |
| 열필드 | 필수 | 열 키 열 |
| 값 | 필수 | 집계할 값 |
| 함수 | 필수 | 집계자 텍스트: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 확장(파라미터=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 중 하나 |
| 필드머리글 | 선택 | 0 없음·숨김, 1 있음·숨김, 2 없음·표시, 3 있음·표시 (기본 자동) |
| 행총계깊이 | 선택 | 0 없음, 1 총계(기본); 음수=상단 |
| 행정렬순서 | 선택 | 행 정렬: 열 번호, 음수는 내림차순 |
| 열총계깊이 | 선택 | 0 없음, 1 총계 열(기본); 음수=왼쪽 |
| 열정렬순서 | 선택 | 열 정렬: 1 오름, -1 내림 |
| 필터배열 | 선택 | 행별 TRUE/FALSE 포함 여부 |
| 기준 | 선택 | PERCENTOF 분모 0~2(0 총합 기본); 확장 집계자는 파라미터: TEXTJOIN 구분자, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 반환

열 키 머리글 행을 포함한 행 키 × 열 키 교차 집계표를 2차원 배열로 스필하여 반환합니다. row_fields·col_fields·values가 각각 1열을 넘거나 relative_to가 0~2 밖이면 #VALUE!, 필터를 통과한 행이 없으면 #N/A를 반환하며, 데이터가 없는 교차 셀은 빈 문자열이 됩니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","합계";"a",10,30,40;"b",20,"",20;"합계",30,30,60} | 행×열 합계 피벗 |

## 참고

- row_fields·col_fields·values는 각각 1열만 지원합니다(다중 열 중첩 미지원). 지원 집계자는 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 16종(네이티브와 동일)이며, EGTools 확장으로 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 도 지원합니다 - 확장 집계자의 파라미터(TEXTJOIN 구분자, LARGE/SMALL k, PERCENTILE p, QUARTILE 0~4)는 relative_to 인수로 전달합니다(그 외 이름은 #VALUE!).
- PERCENTOF의 relative_to는 이 구현에서 0(열 합계 기준, 기본)/1(행 합계 기준)/2(총합계 기준)만 지원하며 3·4(부모 기준)는 #VALUE!를 반환합니다.
- 합계 행/열 라벨은 UI 언어를 따릅니다(한국어 UI: 합계, 영어 UI: Total).
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `PIVOTBY` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.PIVOTBY`으로 등록됩니다.
