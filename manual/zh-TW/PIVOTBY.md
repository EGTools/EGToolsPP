# PIVOTBY

**類別**: MS Excel 相容函數 · **原生導入**: Microsoft 365

依列鍵與欄鍵樞紐分析資料並彙總值。彙總器為文字(SUM、COUNT…)。

## 語法

```
=PIVOTBY(列欄位, 欄欄位, 值, 函數, [欄位標題], [列總計深度], [列排序], [欄總計深度], [欄排序], [篩選陣列], [基準])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 列欄位 | 必要 | 列鍵欄(可多欄) |
| 欄欄位 | 必要 | 欄鍵欄(可多欄) |
| 值 | 必要 | 要彙總的值欄(可多欄) |
| 函數 | 必要 | 彙總器文字: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 擴充(參數=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 之一 |
| 欄位標題 | 選用 | 0 無·隱藏, 1 有·隱藏, 2 無·顯示, 3 有·顯示 (預設自動) |
| 列總計深度 | 選用 | 0 無, 1 總計(預設), 2 總計+小計; 負數=頂部 |
| 列排序 | 選用 | 列排序: 欄號，負數遞減 |
| 欄總計深度 | 選用 | 0 無, 1 總計欄(預設), 2 總計+小計欄; 負數=左側 |
| 欄排序 | 選用 | ±欄位編號(可多個)，負數遞減 |
| 篩選陣列 | 選用 | 每列 TRUE/FALSE |
| 基準 | 選用 | PERCENTOF 分母 0~2(預設 0 總計); 擴充彙總器的參數: TEXTJOIN 分隔符, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 傳回

以二維陣列溢出傳回「列鍵 × 欄鍵」交叉彙總表，每個欄欄位對應一列標題列，預設附總計列/欄；row_fields·col_fields·values 與原生相同，各可為多欄。彙總函數·選項值錯誤，或 PERCENTOF 的 relative_to 超出 0~2 範圍時傳回 #VALUE!；沒有通過篩選的列時傳回 #N/A；沒有資料的交叉儲存格（小計列/欄中亦同）為空字串。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | 列×欄合計樞紐分析 |
| `=PIVOTBY({"a";"a";"b"},{"x","p";"x","q";"y","p"},{10;20;30},"SUM")` | {"","x","x","y","Total";"","p","q","p","";"a",10,20,"",30;"b","","",30,30;"Total",10,20,30,60} | 兩個欄欄位 → 兩列標題 |

## 備註

- function 引數以文字指定（而非原生的 lambda），支援的彙總函數為 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 共 16 種（與原生相同），並以 EGTools 擴充另外支援 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) — 擴充彙總函數的參數（TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE 的 p、QUARTILE 0~4）以 relative_to 引數傳入（其他名稱則為 #VALUE!）。
- row_fields·col_fields·values 各可為多欄（與原生相同）：每個欄欄位增加一列標題列，每個欄群組下依 values 的欄數巢狀值欄。值為 2 欄以上時，總計·小計「欄」的資料儲存格也會依各值欄彙總填入 — 原生會將這些儲存格留空，因此僅此處與原生輸出不同（刻意差異）。field_headers 為 2·3（顯示模式）時，標題區塊依序為：以 ", " 連接欄欄位名稱的一列、欄鍵列、列欄位名稱與值名稱列。
- row_total_depth·col_total_depth 的 2 會依第一個列/欄欄位的區塊加入小計列/欄（該軸需 2 個以上欄位，否則視為 1）；負數會將總計置於頂部/左側。row_sort_order·col_sort_order 接受 ±欄位編號（可為陣列）；排序始終保持欄位階層順序，僅改變指定欄位的方向。
- 本實作中 PERCENTOF 的 relative_to 僅支援 0（以欄合計為基準，預設）/1（以列合計為基準）/2（以總計為基準），3·4（以父層為基準）會傳回 #VALUE!。合計列/欄的標籤依 UI 語言而定（韓文 UI: `합계`/`총합계`，英文 UI: `Total`/`Grand Total`）。
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `PIVOTBY` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.PIVOTBY`。
