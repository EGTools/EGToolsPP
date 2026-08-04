# PIVOTBY

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2024 / Microsoft 365

依列鍵與欄鍵樞紐分析資料並彙總值。彙總器為文字(SUM、COUNT…)。

## 語法

```
=PIVOTBY(列欄位, 欄欄位, 值, 函數, [欄位標題], [列總計深度], [列排序], [欄總計深度], [欄排序], [篩選陣列], [基準])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 列欄位 | 必要 | 列鍵欄 |
| 欄欄位 | 必要 | 欄鍵欄 |
| 值 | 必要 | 要彙總的值 |
| 函數 | 必要 | 彙總器文字: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 擴充(參數=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 之一 |
| 欄位標題 | 選用 | 0 無·隱藏, 1 有·隱藏, 2 無·顯示, 3 有·顯示 (預設自動) |
| 列總計深度 | 選用 | 0 無, 1 總計(預設); 負數=頂部 |
| 列排序 | 選用 | 列排序: 欄號，負數遞減 |
| 欄總計深度 | 選用 | 0 無, 1 總計欄(預設); 負數=左側 |
| 欄排序 | 選用 | 1 遞增, -1 遞減 |
| 篩選陣列 | 選用 | 每列 TRUE/FALSE |
| 基準 | 選用 | PERCENTOF 分母 0~2(預設 0 總計); 擴充彙總器的參數: TEXTJOIN 分隔符, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 傳回

以二維陣列溢出傳回包含欄鍵標題列的「列鍵 × 欄鍵」交叉彙總表。row_fields·col_fields·values 各超過 1 欄，或 relative_to 超出 0~2 範圍時傳回 #VALUE!；沒有通過篩選的列時傳回 #N/A；沒有資料的交叉儲存格為空字串。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | 列×欄合計樞紐分析 |

## 備註

- row_fields·col_fields·values 各僅支援 1 欄（不支援多欄巢狀）。支援的彙總函數為 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 共 16 種（與原生相同），並以 EGTools 擴充另外支援 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) — 擴充彙總函數的參數（TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE 的 p、QUARTILE 0~4）以 relative_to 引數傳入（其他名稱則為 #VALUE!）。
- 本實作中 PERCENTOF 的 relative_to 僅支援 0（以欄合計為基準，預設）/1（以列合計為基準）/2（以總計為基準），3·4（以父層為基準）會傳回 #VALUE!。
- 合計列/欄的標籤依 UI 語言而定（韓文 UI: `합계`，英文 UI: `Total`）。
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `PIVOTBY` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.PIVOTBY`。
