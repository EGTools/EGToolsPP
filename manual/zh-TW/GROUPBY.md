# GROUPBY

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2024 / Microsoft 365

依鍵將列分組並彙總值。彙總器為文字: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 擴充(參數=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC)。

## 語法

```
=GROUPBY(列欄位, 值, 函數, [欄位標題], [總計深度], [排序], [篩選陣列], [欄位關係], [基準值])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 列欄位 | 必要 | 鍵欄 |
| 值 | 必要 | 要彙總的值欄 |
| 函數 | 必要 | 彙總器文字: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 擴充(參數=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 之一 |
| 欄位標題 | 選用 | 0 無·隱藏, 1 有·隱藏, 2 無·產生顯示, 3 有·顯示 (預設自動) |
| 總計深度 | 選用 | 0 無, 1 總計(預設), 2 總計+小計; 負數=頂部 |
| 排序 | 選用 | 輸出欄號，負數遞減，如 {2,-1} |
| 篩選陣列 | 選用 | 每列 TRUE/FALSE |
| 欄位關係 | 選用 | 相容接受但忽略 |
| 基準值 | 選用 | 擴充彙總器參數: TEXTJOIN 分隔符, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 傳回

以溢出方式傳回由群組鍵欄與彙總值欄組成的二維陣列。缺少必要引數、鍵與值的列數不一致、彙總函數或選項值錯誤時傳回 #VALUE!，沒有通過篩選的列時傳回 #N/A；AVERAGE·PERCENTOF 中分母為 0 的儲存格會成為 #DIV/0!。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=GROUPBY({"a";"b";"a"},{10;20;30},"SUM")` | {"a",40;"b",20;"Total",60} | 各鍵的合計與總計列 |
| `=GROUPBY({"a";"b";"a"},{10;20;30},"COUNT",0,0)` | {"a",2;"b",1} | 不含總計，只計個數 |

## 備註

- function 引數以文字指定（而非原生的 lambda），支援的彙總函數為 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 共 16 種（與原生相同），並以 EGTools 擴充另外支援 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) — 擴充彙總函數的參數（TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE 的 p、QUARTILE 0~4）以 relative_to 引數傳入（其他名稱則為 #VALUE!）。field_relationship 引數會接受但被忽略。
- 總計·小計的標籤依 UI 語言而定（韓文 UI: `합계`/`총합계`，英文 UI: `Total`/`Grand Total`）。小計列的標籤為第一個鍵的值。
- 省略 field_headers 時，若第一列全為文字且其下方含有非文字值，會自動辨識為標題。
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `GROUPBY` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.GROUPBY`。
