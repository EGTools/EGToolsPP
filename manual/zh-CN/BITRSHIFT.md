# BITRSHIFT

**类别**：MS Excel 兼容函数 · **原生引入**：Excel 2013

返回数字右移指定位数后的结果。

## 语法

```
=BITRSHIFT(number, shift_amount)
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| number | 必需 | 要移位的数字（≥0） |
| shift_amount | 必需 | 右移位数（负数为左移） |

## 返回

以数字（标量）返回右移结果（向下取整）。number 不是 [0, 2^48-1] 范围内的整数、shift_amount 不是整数或绝对值超过 53、或结果超过 2^53 时返回 #NUM!。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=BITRSHIFT(13,2)` | 3 | 右移 2 位 |
| `=BITRSHIFT(3,-2)` | 12 | 负数为左移 |

## 备注

- shift_amount 为负数时按左移处理。
- 相关函数：BITLSHIFT
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `BITRSHIFT` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.BITRSHIFT`。
