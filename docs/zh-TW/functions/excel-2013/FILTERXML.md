# FILTERXML

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2013

使用指定的 XPath 從 XML 內容傳回資料。

## 語法

```
=FILTERXML(xml, xpath)
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| xml | 必要 | XML 文字 |
| xpath | 必要 | XPath 查詢 |

## 傳回

與 XPath 相符的節點只有 1 個時以純量傳回其文字，多個時以溢出傳回 N×1 的直向文字陣列。XML/XPath 為空，或 XML 剖析·XPath 處理失敗時傳回 #VALUE!；沒有相符的節點時傳回 #N/A。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=FILTERXML("<r><a>1</a><a>2</a></r>","//a")` | {"1";"2"} | 所有 a 節點的文字 |
| `=FILTERXML(WEBSERVICE("https://example.com/rss.xml"),"//item/title")` |  | 視外部 XML 而定 |

## 備註

- 以 MSXML6 剖析，函數本身不會存取網路 — 遠端 XML 請與 WEBSERVICE 搭配使用。
- 結果為節點的文字，且一律以文字格式傳回。
- 因使用 COM，不參與多執行緒重算（非執行緒安全）。
- xml·xpath指定為陣列時，會傳回與該陣列相同形狀的陣列，每個元素為該元素結果的第一個值（與原生一致的降級提升）。
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `FILTERXML` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.FILTERXML`。
