# GROUPBY

**種類**: MS Excel 互換関数 · **ネイティブ導入**: Microsoft 365

キーで行をグループ化して値を集計します。集計子はテキスト: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 拡張(パラメータ=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC)。

## 構文

```
=GROUPBY(行フィールド, 値, 関数, [見出し], [合計深さ], [並び順], [フィルター配列], [フィールド関係], [基準値])
```

## 引数

| 引数 | 必須 | 説明 |
|---|---|---|
| 行フィールド | 必須 | キー列 |
| 値 | 必須 | 集計する値列 |
| 関数 | 必須 | 集計子テキスト: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 拡張(パラメータ=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) のいずれか |
| 見出し | 省略可能 | 0 なし・非表示, 1 あり・非表示, 2 なし・生成表示, 3 あり・表示 (既定 自動) |
| 合計深さ | 省略可能 | 0 なし, 1 総計(既定), 2 総計+小計; 負数=上部 |
| 並び順 | 省略可能 | 出力列番号、負数は降順、例 {2,-1} |
| フィルター配列 | 省略可能 | 行ごとの TRUE/FALSE |
| フィールド関係 | 省略可能 | 互換のため受け付けるが無視 |
| 基準値 | 省略可能 | 拡張集計子のパラメータ: TEXTJOIN 区切り, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 戻り値

グループキー列と集計値列で構成される 2 次元配列をスピルで返します。必須引数の欠落、キーと値の行数不一致、無効な集計子・オプション値の場合は #VALUE!、フィルターを通過した行がない場合は #N/A を返し、AVERAGE・PERCENTOF で分母が 0 のセルは #DIV/0! になります。

## 例

| 数式 | 結果 | 説明 |
|---|---|---|
| `=GROUPBY({"a";"b";"a"},{10;20;30},"SUM")` | {"a",40;"b",20;"Total",60} | キーごとの合計と総計行 |
| `=GROUPBY({"a";"b";"a"},{10;20;30},"COUNT",0,0)` | {"a",2;"b",1} | 総計なしで件数のみ |

## 解説

- function 引数はネイティブのラムダの代わりにテキストで指定します。サポートされる集計子は SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF の 16 種(ネイティブと同一)で、EGTools 拡張として TEXTJOIN/LARGE/SMALL/PERCENTILE・QUARTILE(.INC/.EXC) もサポートします - 拡張集計子のパラメーター(TEXTJOIN の区切り文字、LARGE/SMALL の k、PERCENTILE の p、QUARTILE の 0~4)は relative_to 引数で渡します(それ以外の名前は #VALUE!)。field_relationship 引数は受け取りますが無視されます。
- sort_order はキーフィールドの階層順を保ちます: キーフィールド番号は該当フィールドの並び方向のみを変え、値列の番号を指定すると最後のキーフィールドの代わりにその集計値で並べ替えます(同点はそのキーの昇順)。
- 総計・小計のラベルは UI 言語に従います(韓国語 UI: 합계/총합계、英語 UI: Total/Grand Total)。小計行のラベルは最初のキーの値です。
- field_headers を省略した場合、先頭行がすべてテキストで、その下に非テキスト値があるときにヘッダーとして自動認識します。
- 対応: Excel 2010+。ネイティブ関数がない旧バージョンでは `GROUPBY` の名前のまま（ドロップイン）、ネイティブ関数がある新しい Excel では `EG.GROUPBY` として登録されます。
