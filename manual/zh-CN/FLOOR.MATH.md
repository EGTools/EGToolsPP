# FLOOR.MATH

**类别**：MS Excel 兼容函数 · **原生引入**：Excel 2013

将数字向下舍入到最接近的整数或指定基数的倍数。

## 语法

```
=FLOOR.MATH(number, [significance], [mode])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| number | 必需 | 要舍入的数字 |
| significance | 可选 | [可选] 倍数（默认 1） |
| mode | 可选 | [可选] 非零时负数趋向零舍入 |

## 返回

返回向下舍入到 significance 倍数的数字（标量）。number 非数字时返回 #VALUE! 错误；significance 为 0 时返回 0。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=FLOOR.MATH(6.7)` | 6 | 按默认倍数 1 向下舍入 |
| `=FLOOR.MATH(-5.5,2)` | -6 | 负数远离 0 |
| `=FLOOR.MATH(-5.5,2,1)` | -4 | mode≠0：向 0 方向 |

## 备注

- 忽略 significance 的符号，使用其绝对值。
- 相关函数：CEILING.MATH
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `FLOOR.MATH` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.FLOOR.MATH`。
