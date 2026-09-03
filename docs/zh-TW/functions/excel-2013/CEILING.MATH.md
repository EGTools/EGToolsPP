# CEILING.MATH

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2013

將數字無條件進位到最接近的整數或指定基準的倍數。

## 語法

```
=CEILING.MATH(number, [significance], [mode])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| number | 必要 | 要進位的數字 |
| significance | 選用 | [選用] 倍數（預設 1） |
| mode | 選用 | [選用] 非零時負數遠離零進位 |

## 傳回

傳回進位到 significance 倍數的數字(純量)。number 不是數字時傳回 #VALUE! 錯誤，significance 為 0 時傳回 0。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=CEILING.MATH(6.3)` | 7 | 以預設倍數 1 進位 |
| `=CEILING.MATH(-5.5,2)` | -4 | 負數朝 0 的方向 |
| `=CEILING.MATH(-5.5,2,1)` | -6 | mode≠0: 朝遠離 0 的方向 |

## 備註

- 忽略 significance 的正負號，使用其絕對值。
- 所有引數指定為陣列時會逐元素計算，以相同形狀的陣列溢出傳回 — 純量會重複使用，欄向量×列向量按外積擴充，大小不符的元素為 #N/A，錯誤元素會原樣傳回該錯誤。
- 相關函數: FLOOR.MATH
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `CEILING.MATH` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.CEILING.MATH`。
