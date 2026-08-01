# GROUPBY

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Excel 2024 / Microsoft 365

키로 행을 그룹화해 값을 집계합니다. 집계자는 텍스트: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF.

## 구문

```
=GROUPBY(행필드, 값, 함수, [필드머리글], [총계깊이], [정렬순서], [필터배열], [필드관계])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 행필드 | 필수 | 키 열(들) |
| 값 | 필수 | 집계할 값 열(들) |
| 함수 | 필수 | 집계자 텍스트: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 중 하나 |
| 필드머리글 | 선택 | 0 없음·숨김, 1 있음·숨김, 2 없음·생성표시, 3 있음·표시 (기본 자동) |
| 총계깊이 | 선택 | 0 없음, 1 총계(기본), 2 총계+부분합; 음수=상단 배치 |
| 정렬순서 | 선택 | 출력 열 번호(들), 음수는 내림차순, 예 {2,-1} |
| 필터배열 | 선택 | 행별 TRUE/FALSE 포함 여부 |
| 필드관계 | 선택 | 호환용으로 받되 무시됨 |

## 반환

그룹 키 열과 집계 값 열로 구성된 2차원 배열을 스필로 반환합니다. 필수 인수 누락, 키·값 행 수 불일치, 잘못된 집계자·옵션 값이면 #VALUE!, 필터를 통과한 행이 없으면 #N/A를 반환하며, AVERAGE·PERCENTOF에서 분모가 0인 셀은 #DIV/0!이 됩니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=GROUPBY({"a";"b";"a"},{10;20;30},"SUM")` | {"a",40;"b",20;"합계",60} | 키별 합계와 총계 행 |
| `=GROUPBY({"a";"b";"a"},{10;20;30},"COUNT",0,0)` | {"a",2;"b",1} | 총계 없이 개수만 |

## 참고

- function 인수는 네이티브의 람다 대신 텍스트로 지정하며, 지원 집계자는 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 16종(네이티브와 동일)입니다(그 외는 #VALUE!). field_relationship 인수는 받되 무시됩니다.
- 총계·부분합 라벨은 UI 언어를 따릅니다(한국어 UI: 합계/총합계, 영어 UI: Total/Grand Total). 부분합 행의 라벨은 첫 번째 키 값입니다.
- field_headers를 생략하면 첫 행이 모두 텍스트이고 그 아래에 비텍스트 값이 있을 때 헤더로 자동 인식합니다.
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `GROUPBY` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.GROUPBY`으로 등록됩니다.
