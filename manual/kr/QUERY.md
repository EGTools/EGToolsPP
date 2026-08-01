# QUERY

**종류**: Google Sheets 호환 함수

데이터에 SQL 쿼리를 실행합니다(테이블명: EGTOOLS).

## 구문

```
=QUERY(데이터, 쿼리문, [제목행수])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 데이터 | 필수 | 쿼리할 범위 또는 배열 |
| 쿼리문 | 필수 | SQL 문, 예 "SELECT F1, SUM(F2) FROM EGTOOLS GROUP BY F1" |
| 제목행수 | 선택 | 헤더 행 수: N≥1이면 위 N행이 필드명, 0 없음(F1, F2, …), -1/생략 자동 추정 |

## 반환

SQL 실행 결과를 2차원 배열(값 행만, 머리글 행 없음)로 스필하여 반환합니다. data/query 누락이나 빈 쿼리는 #VALUE!, 결과 행이 없으면 #N/A이며, ADO/ACE 공급자 오류·임시 CSV 기록 실패는 "ERROR: …" 텍스트로 반환됩니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=QUERY(A1:C10,"SELECT F1, SUM(F3) FROM EGTOOLS GROUP BY F1",0)` |  | 데이터에 따라 다름 |

## 참고

- Google Query Language가 아니라 ACE 텍스트 드라이버 + ADO 기반 SQL을 사용하며, 테이블 이름은 EGTOOLS입니다(쿼리 안의 EGTOOLS는 자동으로 EGTOOLS#csv로 치환).
- 열 이름은 headers로 지정한 헤더 행 텍스트, 없으면 F1, F2, …입니다(headers: N≥1=위 N행이 필드명, 0=없음, 생략/-1=자동 추정).
- 데이터를 %TEMP%\EGTOOLS.csv에 기록한 뒤 조회하므로 Microsoft ACE OLEDB(12.0/16.0) 공급자가 필요하며, 공용 임시 파일 사용으로 다중 스레드 재계산에서 제외됩니다.
- 지원: Excel 2010+. 모든 Excel 버전에서 `QUERY` 이름 그대로 등록됩니다.
