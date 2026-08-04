# GOOGLETRANSLATE

**類別**: Google Sheets 相容函數

使用 Google 翻譯服務翻譯文字。

## 語法

```
=GOOGLETRANSLATE(文字, [來源語言], [目標語言])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 文字 | 必要 | 要翻譯的文字(或範圍) |
| 來源語言 | 選用 | 來源語言代碼，如 "en"(預設 "auto" 自動偵測) |
| 目標語言 | 選用 | 結果語言代碼(預設: 介面語言) |

## 傳回

以純量傳回翻譯後的文字。HTTP 要求本身失敗時傳回 #VALUE!；服務錯誤·頁面格式變更·語言代碼錯誤·超過 5000 字時以 "ERROR:…" 文字傳回；輸入空文字時傳回空字串。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=GOOGLETRANSLATE("Hello","en","ko")` |  | 視翻譯結果而定 |

## 備註

- 以抓取非官方 translate.google.com/m 頁面的方式運作 — 若 Google 變更格式可能失效，且需要網際網路連線。
- source_language 預設為 auto，target_language 預設為增益集的 UI 語言。
- 範圍/陣列輸入會以換行連接後一次翻譯；超過 5000 字時傳回 "ERROR:Exceed max length"。
- 支援: Excel 2010+。在所有 Excel 版本中一律以 `GOOGLETRANSLATE` 原名註冊。
