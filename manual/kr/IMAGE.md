# IMAGE

**종류**: MS Excel 호환 함수 · **네이티브 도입**: Excel 2024 / Microsoft 365

셀에 URL 또는 파일 경로의 그림을 삽입합니다.

## 구문

```
=IMAGE(source, [alt_text], [sizing], [height], [width])
```

## 인수

| 인수 | 필수 | 설명 |
|---|---|---|
| source | 필수 | 이미지 URL 또는 파일 경로 |
| alt_text | 선택 | 대체(접근성) 텍스트 (선택) |
| sizing | 선택 | 0 비율맞춤, 1 채움, 2 원본 크기, 3 사용자지정 (선택) |
| height | 선택 | sizing이 3일 때 높이(픽셀) (선택) |
| width | 선택 | sizing이 3일 때 너비(픽셀) (선택) |

## 반환

<!-- TODO: 반환 형태(스칼라/배열·스필)와 의미, 오류 조건(#VALUE!, #N/A 등) -->

## 예제

| 수식 | 결과 | 설명 |
|---|---|---|
| `=IMAGE(...)` | | <!-- TODO --> |

## 참고

<!-- TODO: 원본(Excel/Google)과의 차이, 관련 함수 링크 -->
- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 `IMAGE` 그대로(드롭인), 네이티브가 있는 최신 Excel에서는 `EG.IMAGE`으로 등록됩니다.
