# IMPORTHTML

**類別**: Google Sheets 相容函數

從網頁中擷取表格(table)或清單(list)。

## 語法

```
=IMPORTHTML(網址, [類型], [序號], [地區設定])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 網址 | 必要 | 網頁 URL |
| 類型 | 選用 | "table"(預設) 或 "list"(ul/ol) |
| 序號 | 選用 | 擷取第幾個元素，從 1 開始(預設 1) |
| 地區設定 | 選用 | 數字解析地區，如 "de-DE" |

## 傳回

將網頁的表格（二維陣列）或清單（N×1 陣列）以溢出傳回，數字文字會自動轉換。缺少 URL、search 不是 "table"/"list" 或 index 小於 1 時傳回 #VALUE!；下載失敗或該 index 的表格/清單不存在時傳回 #N/A。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=IMPORTHTML("https://en.wikipedia.org/wiki/Microsoft_Excel","table",1)` |  | 視頁面內容而定 |

## 備註

- 會先移除 <script> 再以 MSHTML 剖析，因此無法取得由指令碼產生的動態內容。
- 表格的 rowspan/colspan 會被正規化為格線，值只填入左上角的儲存格。
- 因使用 COM，不參與多執行緒重算，且需要網際網路連線。
- 支援: Excel 2010+。在所有 Excel 版本中一律以 `IMPORTHTML` 原名註冊。
