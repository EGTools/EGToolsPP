# QUERY

**類別**: Google Sheets 相容函數

對資料執行 SQL 查詢(資料表名: EGTOOLS)。

## 語法

```
=QUERY(資料, 查詢語句, [標題列數])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 資料 | 必要 | 要查詢的範圍或陣列 |
| 查詢語句 | 必要 | SQL 語句，如 "SELECT F1, SUM(F2) FROM EGTOOLS GROUP BY F1" |
| 標題列數 | 選用 | 標題列數: N≥1 時頂部 N 列為欄位名，0 無(F1, F2, …)，-1/省略 自動偵測 |

## 傳回

將 SQL 執行結果以二維陣列（僅值列，不含標題列）溢出傳回。缺少 data/query 或查詢為空時傳回 #VALUE!；結果沒有任何列時傳回 #N/A；ADO/ACE 提供者錯誤·暫存 CSV 寫入失敗會以 "ERROR: …" 文字傳回。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=QUERY(A1:C10,"SELECT F1, SUM(F3) FROM EGTOOLS GROUP BY F1",0)` |  | 視資料而定 |

## 備註

- 並非 Google Query Language，而是使用 ACE 文字驅動程式 + 以 ADO 為基礎的 SQL；資料表名稱為 EGTOOLS（查詢中的 EGTOOLS 會自動替換為 EGTOOLS#csv）。
- 欄名為 headers 指定的標題列文字，若無則為 F1, F2, …（headers: N≥1=頂部 N 列為欄位名，0=無，省略/-1=自動推測）。
- 資料會先寫入 %TEMP%\EGTOOLS.csv 再查詢，因此需要 Microsoft ACE OLEDB(12.0/16.0) 提供者；因使用共用暫存檔案，不參與多執行緒重算。
- 支援: Excel 2010+。在所有 Excel 版本中一律以 `QUERY` 原名註冊。
