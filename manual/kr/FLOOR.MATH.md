# FLOOR.MATH

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Excel 2013

수를 가장 가까운 정수 또는 지정한 배수로 내림합니다.

## 구문

```
=FLOOR.MATH(number, [significance], [mode])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| number | 필수 | 내림할 수 |
| significance | 선택 | [선택] 배수(기본값 1) |
| mode | 선택 | [선택] 0이 아니면 음수를 0에 가깝게 내림 |

## 반환

significance의 배수로 내림한 숫자(스칼라)를 반환합니다. number가 숫자가 아니면 #VALUE! 오류를 반환하고, significance가 0이면 0을 반환합니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=FLOOR.MATH(6.7)` | 6 | 기본 배수 1로 내림 |
| `=FLOOR.MATH(-5.5,2)` | -6 | 음수는 0에서 먼 쪽 |
| `=FLOOR.MATH(-5.5,2,1)` | -4 | mode≠0: 0 방향으로 |

## 참고

- significance의 부호는 무시하고 절댓값을 사용합니다.
- 관련 함수: CEILING.MATH
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `FLOOR.MATH` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.FLOOR.MATH`으로 등록됩니다.
