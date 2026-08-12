# PERMUTATIONA

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Excel 2013

중복을 허용한 순열의 수(number^number_chosen)를 반환합니다.

## 구문

```
=PERMUTATIONA(number, number_chosen)
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| number | 필수 | 전체 항목 수 |
| number_chosen | 필수 | 각 순열에서 선택할 항목 수 |

## 반환

중복을 허용한 순열의 수 number^number_chosen(스칼라 숫자)을 반환합니다. 인수가 숫자가 아니거나 음수이면 #NUM! 오류를 반환합니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=PERMUTATIONA(3,2)` | 9 | 3^2 |
| `=PERMUTATIONA(2,5)` | 32 | 2^5 |

## 참고

- 소수 부분은 버립니다.
- 모든 인수에 배열을 지정하면 원소별로 계산해 같은 모양의 배열로 스필됩니다 — 스칼라는 반복, 열 벡터×행 벡터는 외적으로 확장, 크기가 맞지 않는 원소는 #N/A, 오류 원소는 그 오류를 그대로 반환합니다.
- 관련 함수: COMBINA
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `PERMUTATIONA` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.PERMUTATIONA`으로 등록됩니다.
