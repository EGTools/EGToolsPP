# REGEXTEST

**類別**: MS Excel 相容函數 · **原生導入**: Microsoft 365

檢驗文字是否與正規表示式相符。

## 語法

```
=REGEXTEST(文字, 模式, [區分大小寫])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| 文字 | 必要 | 要檢驗的文字 |
| 模式 | 必要 | 正規模式(ECMAScript) |
| 區分大小寫 | 選用 | 0 區分(預設)，1 忽略 |

## 傳回

傳回表示是否相符的邏輯值 TRUE/FALSE（純量）。正規表示式無效時傳回 #VALUE!。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=REGEXTEST("abc123","\d+")` | TRUE | 包含數字 |
| `=REGEXTEST("ABC","[a-z]+",1)` | TRUE | 忽略大小寫 |
| `=REGEXTEST("abc","^\d")` | FALSE | 沒有相符項 |

## 備註

- 正規表示式語法為 std::wregex 的 ECMAScript（可能與原生 365 的 PCRE2 略有差異）。
- 相關函數：REGEXEXTRACT、REGEXREPLACE
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `REGEXTEST` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.REGEXTEST`。
