# VISIBLEAGGR

**種類**: EGTools 専用関数

範囲の表示セルのみを列挙または集計します。

## 構文

```
=VISIBLEAGGR(範囲, [関数], [オプション])
```

## 引数

| 引数 | 必須 | 説明 |
|---|---|---|
| 範囲 | 必須 | 計算するセル範囲 |
| 関数 | 省略可能 | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S·P/VAR.S·P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC)、省略時は表示セルを列挙 |
| オプション | 省略可能 | TEXTJOIN 区切り文字、LARGE/SMALL の k、PERCENTILE p(0~1)、QUARTILE 0~4 |

## 戻り値

関数を省略すると表示セルのみを 2 次元配列としてスピルし、関数を指定するとスカラーの集計値を返します。表示セルがない場合は #N/A を返し、表示セルにエラーがある場合は COUNT/COUNTA 以外ではそのエラーをそのまま返し、サポートしていない関数名は #VALUE! になります。

## 例

| 数式 | 結果 | 説明 |
|---|---|---|
| `=VISIBLEAGGR(A1:B10)` |  | 表示セルのみを列挙(非表示の状態に依存) |
| `=VISIBLEAGGR(A1:A10,"SUM")` |  | 表示セルの合計(非表示の状態に依存) |
| `=VISIBLEAGGR(A1:A10,"LARGE",2)` |  | 表示されている値のうち 2 番目に大きい値 |

## 解説

- サポートする集計関数(21 種): SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE(.INC/.EXC)/QUARTILE(.INC/.EXC)。それ以外は #VALUE! です。
- オプション引数: TEXTJOIN=区切り文字、LARGE/SMALL=k、PERCENTILE 系=p(0～1)、QUARTILE 系=0～4。
- 行/列の表示・非表示が変更されても自動では再計算されないため、F9 での再計算が必要です。マクロ型のためマルチスレッド再計算からは除外されます。
- 対応: Excel 2010+。すべての Excel バージョンで `VISIBLEAGGR` の名前のまま登録されます。
