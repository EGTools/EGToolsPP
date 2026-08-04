# FLOOR.MATH

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2013

將數字無條件捨去到最接近的整數或指定基準的倍數。

## 語法

```
=FLOOR.MATH(number, [significance], [mode])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| number | 必要 | 要捨去的數字 |
| significance | 選用 | [選用] 倍數（預設 1） |
| mode | 選用 | [選用] 非零時負數趨向零捨去 |

## 傳回

傳回捨位到 significance 倍數的數字(純量)。number 不是數字時傳回 #VALUE! 錯誤，significance 為 0 時傳回 0。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=FLOOR.MATH(6.7)` | 6 | 以預設倍數 1 捨位 |
| `=FLOOR.MATH(-5.5,2)` | -6 | 負數朝遠離 0 的方向 |
| `=FLOOR.MATH(-5.5,2,1)` | -4 | mode≠0: 朝 0 的方向 |

## 備註

- 忽略 significance 的正負號，使用其絕對值。
- 相關函數: CEILING.MATH
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `FLOOR.MATH` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.FLOOR.MATH`。
