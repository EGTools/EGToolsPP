# CEILING.MATH

**種類**: MS Excel 互換関数 · **ネイティブ導入**: Excel 2013

数値を最も近い整数または基準値の倍数に切り上げます。

## 構文

```
=CEILING.MATH(number, [significance], [mode])
```

## 引数

| 引数 | 必須 | 説明 |
|---|---|---|
| number | 必須 | 丸める数値 |
| significance | 省略可能 | [省略可] 倍数（既定 1） |
| mode | 省略可能 | [省略可] 0 以外で負数を 0 から離れる方向に丸める |

## 戻り値

significance の倍数に切り上げた数値(スカラー)を返します。number が数値でない場合は #VALUE! エラーを返し、significance が 0 の場合は 0 を返します。

## 例

| 数式 | 結果 | 説明 |
|---|---|---|
| `=CEILING.MATH(6.3)` | 7 | 既定の倍数 1 で切り上げ |
| `=CEILING.MATH(-5.5,2)` | -4 | 負数は 0 に近い方向へ |
| `=CEILING.MATH(-5.5,2,1)` | -6 | mode≠0: 0 から遠い方向へ |

## 解説

- significance の符号は無視し、絶対値を使用します。
- 関連関数: FLOOR.MATH。
- 対応: Excel 2010+。ネイティブ関数がない旧バージョンでは `CEILING.MATH` の名前のまま（ドロップイン）、ネイティブ関数がある新しい Excel では `EG.CEILING.MATH` として登録されます。
