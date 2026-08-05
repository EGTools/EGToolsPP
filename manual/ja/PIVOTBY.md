# PIVOTBY

**種類**: MS Excel 互換関数 · **ネイティブ導入**: Microsoft 365

行キーと列キーでデータをピボットし値を集計します。集計子はテキスト(SUM, COUNT…)。

## 構文

```
=PIVOTBY(行フィールド, 列フィールド, 値, 関数, [見出し], [行合計深さ], [行並び順], [列合計深さ], [列並び順], [フィルター配列], [基準])
```

## 引数

| 引数 | 必須 | 説明 |
|---|---|---|
| 行フィールド | 必須 | 行キー列 |
| 列フィールド | 必須 | 列キー列 |
| 値 | 必須 | 集計する値 |
| 関数 | 必須 | 集計子テキスト: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 拡張(パラメータ=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) のいずれか |
| 見出し | 省略可能 | 0 なし・非表示, 1 あり・非表示, 2 なし・表示, 3 あり・表示 (既定 自動) |
| 行合計深さ | 省略可能 | 0 なし, 1 総計(既定); 負数=上部 |
| 行並び順 | 省略可能 | 行の並び: 列番号、負数は降順 |
| 列合計深さ | 省略可能 | 0 なし, 1 合計列(既定); 負数=左 |
| 列並び順 | 省略可能 | 1 昇順, -1 降順 |
| フィルター配列 | 省略可能 | 行ごとの TRUE/FALSE |
| 基準 | 省略可能 | PERCENTOF の分母 0~2(既定 0 総計); 拡張集計子ではパラメータ: TEXTJOIN 区切り, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 戻り値

列キーの見出し行を含む、行キー × 列キーのクロス集計表を 2 次元配列としてスピルで返します。row_fields・col_fields・values がそれぞれ 1 列を超える場合や relative_to が 0~2 の範囲外の場合は #VALUE!、フィルターを通過した行がない場合は #N/A を返し、データのない交差セルは空文字列になります。

## 例

| 数式 | 結果 | 説明 |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | 行×列の合計ピボット |

## 解説

- row_fields・col_fields・values はそれぞれ 1 列のみサポートします(複数列のネストはサポートしていません)。サポートされる集計子は SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF の 16 種(ネイティブと同一)で、EGTools 拡張として TEXTJOIN/LARGE/SMALL/PERCENTILE・QUARTILE(.INC/.EXC) もサポートします - 拡張集計子のパラメーター(TEXTJOIN の区切り文字、LARGE/SMALL の k、PERCENTILE の p、QUARTILE の 0~4)は relative_to 引数で渡します(それ以外の名前は #VALUE!)。
- PERCENTOF の relative_to は本実装では 0(列合計基準、既定)/1(行合計基準)/2(総合計基準)のみをサポートし、3・4(親基準)は #VALUE! を返します。
- 合計行/列のラベルは UI 言語に従います(韓国語 UI: 합계、英語 UI: Total)。
- 対応: Excel 2010+。ネイティブ関数がない旧バージョンでは `PIVOTBY` の名前のまま（ドロップイン）、ネイティブ関数がある新しい Excel では `EG.PIVOTBY` として登録されます。
