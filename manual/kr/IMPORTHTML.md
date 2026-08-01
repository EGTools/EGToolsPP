# IMPORTHTML

**종류**: Google Sheets 호환 함수

웹 페이지에서 표(table)나 목록(list)을 추출합니다.

## 구문

```
=IMPORTHTML(주소, [종류], [번호], [로캘])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 주소 | 필수 | 웹 페이지 URL |
| 종류 | 선택 | "table"(기본) 또는 "list"(ul/ol) |
| 번호 | 선택 | 몇 번째 요소를 추출할지, 1부터 (기본 1) |
| 로캘 | 선택 | 숫자 해석 로캘, 예 "de-DE" |

## 반환

웹 페이지의 표(2차원 배열) 또는 목록(N×1 배열)을 스필하여 반환하며 숫자 텍스트는 자동 변환됩니다. URL 누락, search가 "table"/"list"가 아니거나 index가 1 미만이면 #VALUE!, 다운로드 실패나 해당 index의 표/목록이 없으면 #N/A를 반환합니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=IMPORTHTML("https://en.wikipedia.org/wiki/Microsoft_Excel","table",1)` |  | 페이지 내용에 따름 |

## 참고

- <script>를 제거한 뒤 MSHTML로 파싱하므로 스크립트가 만들어 내는 동적 콘텐츠는 가져올 수 없습니다.
- 표의 rowspan/colspan은 그리드로 정규화되며 값은 좌상단 셀에만 들어갑니다.
- COM 사용으로 다중 스레드 재계산에서 제외되며 인터넷 연결이 필요합니다.
- 지원: Excel 2010+. 모든 Excel 버전에서 `IMPORTHTML` 이름 그대로 등록됩니다.
