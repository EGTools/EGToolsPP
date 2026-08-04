# IMPORTFEED

**類別**: Google Sheets 相容函數

匯入 RSS/ATOM 摘要。

## 語法

```
=IMPORTFEED(網址, [查詢], [標題列], [項目數])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 網址 | 必要 | 摘要 URL |
| 查詢 | 選用 | "items"(預設)、"feed" 或欄位: title/summary/description/author/url/created |
| 標題列 | 選用 | TRUE 時包含標題列 |
| 項目數 | 選用 | 傳回的項目數(預設: 全部) |

## 傳回

將摘要項目（或摘要資訊·單一欄位）以二維陣列溢出傳回。缺少 URL 或 query 欄位不受支援時傳回 #VALUE!；下載失敗·XML 剖析失敗·沒有項目時傳回 #N/A。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=IMPORTFEED("https://blog.example.com/rss","items",TRUE,5)` |  | 視摘要內容而定 |

## 備註

- 自動判別 RSS 2.0 與 Atom。query "items"（預設）傳回 Title/Summary/URL/Author/Created 5 欄。
- created 欄位的 RFC822/ISO8601 日期會轉換為 Excel 序列值（忽略時區位移）。
- 因使用 COM(MSXML)，不參與多執行緒重算，且需要網際網路連線。
- 支援: Excel 2010+。在所有 Excel 版本中一律以 `IMPORTFEED` 原名註冊。
