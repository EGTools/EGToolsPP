# CEILING.MATH

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Excel 2013

수를 가장 가까운 정수 또는 지정한 배수로 올림합니다.

## 구문

```
=CEILING.MATH(number, [significance], [mode])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| number | 필수 | 올림할 수 |
| significance | 선택 | [선택] 배수(기본값 1) |
| mode | 선택 | [선택] 0이 아니면 음수를 0에서 멀어지게 올림 |

## 반환

<!-- TODO: 반환 형태(스칼라/배열·스필)와 의미, 오류 조건(#VALUE!, #N/A 등) -->

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=CEILING.MATH(...)` | | <!-- TODO --> |

## 참고

<!-- TODO: 원본(Excel/Google)과의 차이, 관련 함수 링크 -->
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `CEILING.MATH` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.CEILING.MATH`으로 등록됩니다.
