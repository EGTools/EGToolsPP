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

피드 항목(또는 피드 정보·단일 필드)을 2차원 배열로 스필하여 반환합니다. URL 누락이나 지원하지 않는 query 필드면 #VALUE!, 다운로드 실패·XML 파싱 실패·항목 없음이면 #N/A를 반환합니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=IMPORTFEED("https://blog.example.com/rss","items",TRUE,5)` |  | 피드 내용에 따름 |

## 참고

- RSS 2.0과 Atom을 자동 판별합니다. query "items"(기본)는 Title/Summary/URL/Author/Created 5열을 반환합니다.
- created 필드의 RFC822/ISO8601 날짜는 Excel 직렬값으로 변환합니다(시간대 오프셋은 무시).
- COM(MSXML) 사용으로 다중 스레드 재계산에서 제외되며 인터넷 연결이 필요합니다.
- url에 배열을 지정하면 #VALUE!를 반환합니다 — 여러 건은 수식을 행별로 복사해 사용합니다.
- 지원: Excel 2010+. 모든 Excel 버전에서 `IMPORTFEED` 이름 그대로 등록됩니다.
