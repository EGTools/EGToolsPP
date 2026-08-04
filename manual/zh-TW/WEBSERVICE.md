# WEBSERVICE

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2013

傳回指定 URL 處 Web 服務的回應（HTTP GET）。

## 語法

```
=WEBSERVICE(url)
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| url | 必要 | 要要求的 URL |

## 傳回

傳回將 HTTP(S) GET 回應本文以 UTF-8 解讀後的文字純量。URL 格式錯誤或連線·要求失敗時傳回 #VALUE!（連線 5 秒、接收回應 15 秒逾時）。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=WEBSERVICE("https://api.frankfurter.app/latest")` |  | 視外部回應而定 |

## 備註

- 以 WinHTTP GET 運作，需要網際網路連線。
- 回應本文一律以 UTF-8 解讀。
- 已註冊為執行緒安全，會在多執行緒重算中平行執行。
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `WEBSERVICE` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.WEBSERVICE`。
