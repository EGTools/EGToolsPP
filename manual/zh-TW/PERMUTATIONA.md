# PERMUTATIONA

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2013

傳回可重複排列數（number^number_chosen）。

## 語法

```
=PERMUTATIONA(number, number_chosen)
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| number | 必要 | 項目總數 |
| number_chosen | 必要 | 每個排列的項目數 |

## 傳回

傳回允許重複的排列數 number^number_chosen(純量數字)。引數不是數字或為負數時傳回 #NUM! 錯誤。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=PERMUTATIONA(3,2)` | 9 | 3^2 |
| `=PERMUTATIONA(2,5)` | 32 | 2^5 |

## 備註

- 捨去小數部分。
- 相關函數: COMBINA
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `PERMUTATIONA` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.PERMUTATIONA`。
