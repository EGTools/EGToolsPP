# PERCENTOF

**類別**: MS Excel 相容函數 · **原生導入**: Microsoft 365

傳回子集合之和佔全部資料之和的比例。

## 語法

```
=PERCENTOF(data_subset, data_all)
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| data_subset | 必要 | 資料子集合 |
| data_all | 必要 | 全部資料 |

## 傳回

以純量數字傳回子集合資料之和除以全部資料之和的比例。缺少引數時傳回 #VALUE!；全部之和為 0 或全部資料中沒有數字時傳回 #DIV/0!。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=PERCENTOF({10;20},{10;20;30;40})` | 0.3 | 子集合之和÷全部之和的比例 |
| `=PERCENTOF(50,200)` | 0.25 | 純量值的比例 |

## 備註

- 非數字的值在計算合計時會被忽略。
- 在 GROUPBY/PIVOTBY 的 function 引數指定 "PERCENTOF" 即可求得各群組的組成比例。
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `PERCENTOF` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.PERCENTOF`。
