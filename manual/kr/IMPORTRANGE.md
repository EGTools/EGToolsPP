# IMPORTRANGE

**종류**: Google Sheets 호환 함수

공유된 Google 스프레드시트의 자료를 가져옵니다.

## 구문

```
=IMPORTRANGE(시트주소, [범위주소])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 시트주소 | 필수 | 구글시트 URL — 공유가 설정되어 있어야 합니다 |
| 범위주소 | 선택 | 가져올 범위, 예 "A1:C10" (기본: 시트 전체) |

## 반환

공유된 Google Sheets의 값을 2차원 배열로 스필하여 반환하며 숫자·날짜 텍스트는 자동 변환됩니다. URL이 비었거나 /d/<ID> 형식이 아니면 #VALUE!, 다운로드 실패·공유 미설정(HTML 응답)·빈 시트·범위가 데이터 밖이면 #N/A를 반환합니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1AbCdEf.../edit#gid=0","A1:C10")` |  | 시트 내용에 따름 |

## 참고

- 링크 공유가 켜진 Google Sheets만 가져올 수 있습니다 — 미공유 문서는 로그인 페이지(HTML)가 내려와 #N/A가 됩니다.
- 시트는 URL의 gid로 선택합니다(기본 gid=0). range_address의 시트 이름 부분은 무시됩니다.
- TSV export 엔드포인트를 사용하므로 값만 가져오며 서식은 유지되지 않습니다.
- 지원: Excel 2010+. 모든 Excel 버전에서 `IMPORTRANGE` 이름 그대로 등록됩니다.
