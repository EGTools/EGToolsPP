# VISIBLEAGGR

**類別**: EGTools 專用函數

僅列出或彙總範圍中可見的儲存格。

## 語法

```
=VISIBLEAGGR(範圍, [函數], [選項])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 範圍 | 必要 | 要計算的儲存格範圍 |
| 函數 | 選用 | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S·P/VAR.S·P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC)，省略時列出可見儲存格 |
| 選項 | 選用 | TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE p(0~1)、QUARTILE 0~4 |

## 傳回

省略函數時，僅將可見儲存格以二維陣列溢出；指定函數時，傳回純量彙總值。沒有可見儲存格時傳回 #N/A；可見儲存格含有錯誤時，除 COUNT/COUNTA 外會原樣傳回該錯誤；不支援的函數名稱則為 #VALUE!。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=VISIBLEAGGR(A1:B10)` |  | 僅列出可見儲存格（依隱藏狀態而定） |
| `=VISIBLEAGGR(A1:A10,"SUM")` |  | 可見儲存格的總和（依隱藏狀態而定） |
| `=VISIBLEAGGR(A1:A10,"LARGE",2)` |  | 可見值中第 2 大的值 |

## 備註

- 支援的彙總函數（21 種）：SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE(.INC/.EXC)/QUARTILE(.INC/.EXC)。其餘為 #VALUE!。
- 選項引數：TEXTJOIN=分隔符、LARGE/SMALL=k、PERCENTILE 系列=p(0~1)、QUARTILE 系列=0~4。
- 列/欄的隱藏狀態變更後不會自動重新計算，需按 F9 重新計算。屬巨集型函數，不參與多執行緒重新計算。
- 支援: Excel 2010+。在所有 Excel 版本中一律以 `VISIBLEAGGR` 原名註冊。
