# IMPORTFEED

**種類**: Google スプレッドシート互換関数

RSS/ATOM フィードを取り込みます。

## 構文

```
=IMPORTFEED(URL, [クエリ], [見出し], [項目数])
```

## 引数

| 引数 | 必須 | 説明 |
|---|---|---|
| URL | 必須 | フィード URL |
| クエリ | 省略可能 | "items"(既定)・"feed" またはフィールド: title/summary/description/author/url/created |
| 見出し | 省略可能 | TRUE で見出し行を含める |
| 項目数 | 省略可能 | 出力する項目数(既定: すべて) |

## 戻り値

フィードの項目(またはフィード情報・単一フィールド)を 2 次元配列としてスピルで返します。URL の欠落やサポートしていない query フィールドの場合は #VALUE!、ダウンロードの失敗・XML の解析失敗・項目なしの場合は #N/A を返します。

## 例

| 数式 | 結果 | 説明 |
|---|---|---|
| `=IMPORTFEED("https://blog.example.com/rss","items",TRUE,5)` |  | フィードの内容に依存します |

## 解説

- RSS 2.0 と Atom を自動判別します。query "items"(既定)は Title/Summary/URL/Author/Created の 5 列を返します。
- created フィールドの RFC822/ISO8601 日付は Excel のシリアル値へ変換します(タイムゾーンオフセットは無視)。
- COM(MSXML) を使用するためマルチスレッド再計算からは除外され、インターネット接続が必要です。
- urlに配列を指定すると #VALUE! を返します — 複数件は数式を行ごとにコピーして使用してください。
- 対応: Excel 2010+。すべての Excel バージョンで `IMPORTFEED` の名前のまま登録されます。
