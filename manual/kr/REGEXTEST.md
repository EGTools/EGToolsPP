# REGEXTEST

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Excel 2024 / Microsoft 365

텍스트가 정규식과 일치하는지 검사합니다.

## 구문

```
=REGEXTEST(텍스트, 패턴, [대소문자구분])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| 텍스트 | 필수 | 검사할 텍스트 |
| 패턴 | 필수 | 정규식 패턴(ECMAScript) |
| 대소문자구분 | 선택 | 0 구분(기본), 1 무시 |

## 반환

일치 여부를 나타내는 논리값 TRUE/FALSE(스칼라)를 반환합니다. 정규식이 잘못되면 #VALUE!를 반환합니다.

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=REGEXTEST("abc123","\d+")` | TRUE | 숫자 포함 |
| `=REGEXTEST("ABC","[a-z]+",1)` | TRUE | 대소문자 무시 |
| `=REGEXTEST("abc","^\d")` | FALSE | 일치 없음 |

## 참고

- 정규식 문법은 std::wregex의 ECMAScript입니다(네이티브 365의 PCRE2와 일부 다를 수 있습니다).
- 관련 함수: REGEXEXTRACT, REGEXREPLACE
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `REGEXTEST` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.REGEXTEST`으로 등록됩니다.
