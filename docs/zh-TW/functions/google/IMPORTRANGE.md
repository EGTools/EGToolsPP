# IMPORTRANGE

**類別**: Google Sheets 相容函數

從共用的 Google 試算表匯入資料。

## 語法

```
=IMPORTRANGE(試算表網址, [範圍位址])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 試算表網址 | 必要 | Google 試算表 URL — 必須已啟用共用 |
| 範圍位址 | 選用 | 要匯入的範圍，如 "A1:C10"(預設: 整個工作表) |

## 傳回

將已共用的 Google Sheets 的值以二維陣列溢出傳回，數字·日期文字會自動轉換。URL 為空或不是 /d/<ID> 格式時傳回 #VALUE!；下載失敗·未設定共用（HTML 回應）·空工作表·範圍超出資料時傳回 #N/A。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1AbCdEf.../edit#gid=0","A1:C10")` |  | 視工作表內容而定 |

## 備註

- 只能取得已開啟連結共用的 Google Sheets — 未共用的文件會收到登入頁面（HTML）而成為 #N/A。
- 工作表以 URL 中的 gid 選擇（預設 gid=0）。range_address 的工作表名稱部分會被忽略。
- 因使用 TSV 匯出端點，只會取得值，不保留格式。
- url指定為陣列時傳回 #VALUE! — 多筆資料請將公式逐列複製使用。
- 支援: Excel 2010+。在所有 Excel 版本中一律以 `IMPORTRANGE` 原名註冊。
