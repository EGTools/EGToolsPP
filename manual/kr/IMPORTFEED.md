# IMPORTFEED

**종류**: Google Sheets 호환 함수

RSS/ATOM 피드 자료를 가져옵니다.

## 구문

```
=IMPORTFEED(주소, [질의], [제목행], [항목수])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 주소 | 필수 | 피드 URL |
| 질의 | 선택 | "items"(기본)·"feed" 또는 필드: title/summary/description/author/url/created |
| 제목행 | 선택 | TRUE면 제목 행 포함 |
| 항목수 | 선택 | 출력할 항목 수(기본: 전체) |

## 반환

<!-- TODO: 반환 형태(스칼라/배열·스필)와 의미, 오류 조건(#VALUE!, #N/A 등) -->

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=IMPORTFEED(...)` | | <!-- TODO --> |

## 참고

<!-- TODO: 원본(Excel/Google)과의 차이, 관련 함수 링크 -->
- 지원: Excel 2010+. 모든 Excel 버전에서 `IMPORTFEED` 이름 그대로 등록됩니다(`EG.`/`x` 접두어 없음).
