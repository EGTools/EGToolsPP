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
| 행필드 | 필수 | 행 키 열(들) |
| 열필드 | 필수 | 열 키 열(들) |
| 값 | 필수 | 집계할 값 열(들) |
| 함수 | 필수 | 집계자 텍스트: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 확장(파라미터=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 중 하나 |
| 필드머리글 | 선택 | 0 없음·숨김, 1 있음·숨김, 2 없음·표시, 3 있음·표시 (기본 자동) |
| 행총계깊이 | 선택 | 0 없음, 1 총계(기본), 2 총계+부분합; 음수=상단 |
| 행정렬순서 | 선택 | 행 정렬: 열 번호, 음수는 내림차순 |
| 열총계깊이 | 선택 | 0 없음, 1 총계 열(기본), 2 총계+부분합 열; 음수=왼쪽 |
| 열정렬순서 | 선택 | 열 정렬: ±필드 번호(들), 음수는 내림차순 |
| 필터배열 | 선택 | 행별 TRUE/FALSE 포함 여부 |
| 기준 | 선택 | PERCENTOF 분모 0~2(0 총합 기본); 확장 집계자는 파라미터: TEXTJOIN 구분자, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 반환

열 필드마다 머리글 행 하나를 포함한 행 키 × 열 키 교차 집계표를 2차원 배열로 스필하여 반환하며, 기본으로 합계 행·열이 붙습니다. row_fields·col_fields·values는 네이티브와 같이 각각 여러 열을 허용합니다. 잘못된 집계자·옵션 값이거나 PERCENTOF의 relative_to가 0~2 밖이면 #VALUE!, 필터를 통과한 행이 없으면 #N/A를 반환하며, 데이터가 없는 교차 셀은 부분합 행·열에서도 빈 문자열이 됩니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","합계";"a",10,30,40;"b",20,"",20;"합계",30,30,60} | 행×열 합계 피벗 |
| `=PIVOTBY({"a";"a";"b"},{"x","p";"x","q";"y","p"},{10;20;30},"SUM")` | {"","x","x","y","합계";"","p","q","p","";"a",10,20,"",30;"b","","",30,30;"합계",10,20,30,60} | 열필드 2개 → 열 키 머리글 2행 |

## 참고

- function 인수는 네이티브의 람다 대신 텍스트로 지정하며, 지원 집계자는 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 16종(네이티브와 동일)이며, EGTools 확장으로 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 도 지원합니다 - 확장 집계자의 파라미터(TEXTJOIN 구분자, LARGE/SMALL k, PERCENTILE p, QUARTILE 0~4)는 relative_to 인수로 전달합니다(그 외 이름은 #VALUE!).
- row_fields·col_fields·values는 각각 여러 열을 허용합니다(네이티브 동일): 열 필드마다 머리글 행이 하나씩 생기고, 각 열 그룹 아래에 값 열이 values의 열 수만큼 중첩됩니다. 값이 2열 이상이어도 합계·부분합 '열'의 데이터 셀을 값 열별로 집계해 채웁니다 - 네이티브는 이 셀들을 비워 두므로 이 부분만 네이티브 출력과 다릅니다(의도적 차이). field_headers 2·3(표시 모드)에서는 맨 위에 열 필드명을 ", "로 연결한 행, 열 키 행들, 행 필드명·값 이름 행 순서로 머리글이 출력됩니다.
- row_total_depth·col_total_depth의 2는 첫 번째 행/열 필드 블록별 부분합 행·열을 추가합니다(해당 축에 필드 2개 이상 필요, 아니면 1로 동작). 음수는 총계를 상단·왼쪽에 배치합니다. row_sort_order·col_sort_order는 ±필드 번호(배열 허용)이며, 정렬은 필드 계층 순서를 유지한 채 지정한 필드의 방향만 바꿉니다.
- PERCENTOF의 relative_to는 이 구현에서 0(열 합계 기준, 기본)/1(행 합계 기준)/2(총합계 기준)만 지원하며 3·4(부모 기준)는 #VALUE!를 반환합니다. 합계 행/열 라벨은 UI 언어를 따릅니다(한국어 UI: 합계/총합계, 영어 UI: Total/Grand Total).
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `PIVOTBY` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.PIVOTBY`으로 등록됩니다.
