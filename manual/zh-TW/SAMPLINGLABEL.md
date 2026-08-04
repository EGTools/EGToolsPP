# SAMPLINGLABEL

**類別**: EGTools 專用函數

依 ISO 2859-1 的批量與檢驗水準傳回樣本量字碼。

## 語法

```
=SAMPLINGLABEL(批量, [檢驗水準])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 批量 | 必要 | 批/批次數量(2 以上) |
| 檢驗水準 | 選用 | 一般 Ⅰ,Ⅱ,Ⅲ→G1/G2/G3，特殊→S1~S4(預設 G2) |

## 傳回

傳回一個 ISO 2859-1 樣本代字字母(文字純量)。LOT 數量小於 2 時傳回 #NUM!，檢驗水準不是 S1~S4·G1~G3 時傳回 #VALUE! 錯誤。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=SAMPLINGLABEL(1000)` | J | 預設 G2，LOT 1000 |
| `=SAMPLINGLABEL(1000,"S3")` | E | 特殊檢驗水準 S-3 |

## 備註

- 忽略檢驗水準表示中的連字號與空格("G-2"="G2")。
- 內建 ISO 2859-1:1999 表格進行計算。
- 相關函數: SAMPLINGSIZE, SAMPLINGAC, SAMPLINGRE
- 支援: Excel 2010+。在所有 Excel 版本中一律以 `SAMPLINGLABEL` 原名註冊。
