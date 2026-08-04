# BITLSHIFT

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2013

傳回數字左移指定位元數後的結果。

## 語法

```
=BITLSHIFT(number, shift_amount)
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| number | 必要 | 要移位的數字（≥0） |
| shift_amount | 必要 | 左移位元數（負數為右移） |

## 傳回

以數字（純量）傳回左移的結果。若 number 不是 [0, 2^48-1] 範圍內的整數，或 shift_amount 不是整數、絕對值超過 53，或結果超過 2^53，則傳回 #NUM!。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=BITLSHIFT(4,2)` | 16 | 左移 2 位元 |
| `=BITLSHIFT(4,-1)` | 2 | 負數為右移 |

## 備註

- shift_amount 為負數時，會以右移方式運作。
- 相關函數：BITRSHIFT
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `BITLSHIFT` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.BITLSHIFT`。
