# SAMPLINGLABEL

**종류**: EGTools 전용 함수

ISO 2859-1의 LOT 크기·검사수준에 따른 시료 문자를 구합니다.

## 구문

```
=SAMPLINGLABEL(LOT수량, [검사수준])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| LOT수량 | 필수 | 검사할 LOT/배치 수량(2 이상) |
| 검사수준 | 선택 | 보통 Ⅰ,Ⅱ,Ⅲ→G1/G2/G3, 특별→S1~S4 (기본 G2) |

## 반환

ISO 2859-1 시료 문자 한 글자(텍스트 스칼라)를 반환합니다. LOT수량이 2 미만이면 #NUM!, 검사수준이 S1~S4·G1~G3이 아니면 #VALUE! 오류를 반환합니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=SAMPLINGLABEL(1000)` | J | 기본 G2, LOT 1000 |
| `=SAMPLINGLABEL(1000,"S3")` | E | 특별검사수준 S-3 |

## 참고

- 검사수준 표기의 하이픈·공백은 무시합니다("G-2"="G2").
- ISO 2859-1:1999 표를 내장하여 계산합니다.
- 관련 함수: SAMPLINGSIZE, SAMPLINGAC, SAMPLINGRE
- 지원: Excel 2010+. 모든 Excel 버전에서 `SAMPLINGLABEL` 이름 그대로 등록됩니다.
