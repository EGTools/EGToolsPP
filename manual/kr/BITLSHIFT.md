# BITLSHIFT

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Excel 2013

수를 지정한 비트 수만큼 왼쪽으로 시프트한 값을 반환합니다.

## 구문

```
=BITLSHIFT(number, shift_amount)
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| number | 필수 | 시프트할 수(0 이상) |
| shift_amount | 필수 | 왼쪽 시프트 비트 수(음수는 오른쪽) |

## 반환

<!-- TODO: 반환 형태(스칼라/배열·스필)와 의미, 오류 조건(#VALUE!, #N/A 등) -->

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=BITLSHIFT(...)` | | <!-- TODO --> |

## 참고

<!-- TODO: 원본(Excel/Google)과의 차이, 관련 함수 링크 -->
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `BITLSHIFT` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.BITLSHIFT`으로 등록됩니다.
