# VISIBLEAGGR

**类别**：EGTools 专用函数

仅列出或聚合区域中可见的单元格。

## 语法

```
=VISIBLEAGGR(区域, [函数], [选项])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 区域 | 必需 | 要计算的单元格区域 |
| 函数 | 可选 | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S·P/VAR.S·P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC)，省略时列出可见单元格 |
| 选项 | 可选 | TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE p(0~1)、QUARTILE 0~4 |

## 返回

省略函数时仅将可见单元格以二维数组溢出，指定函数时返回标量聚合值。没有可见单元格时返回 #N/A；可见单元格含错误时除 COUNT/COUNTA 外原样返回该错误；不支持的函数名返回 #VALUE!。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=VISIBLEAGGR(A1:B10)` |  | 仅列出可见单元格（取决于隐藏状态） |
| `=VISIBLEAGGR(A1:A10,"SUM")` |  | 可见单元格求和（取决于隐藏状态） |
| `=VISIBLEAGGR(A1:A10,"LARGE",2)` |  | 可见值中第 2 大的值 |

## 备注

- 支持的聚合器（21 种）：SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/TEXTJOIN/LARGE/SMALL/PERCENTILE(.INC/.EXC)/QUARTILE(.INC/.EXC)。其余返回 #VALUE!。
- 选项参数：TEXTJOIN=分隔符，LARGE/SMALL=k，PERCENTILE 系=p(0~1)，QUARTILE 系=0~4。
- 行/列隐藏状态变化不会自动重算，需按 F9 重算。宏类函数，被排除在多线程重算之外。
- 支持：Excel 2010+。在所有 Excel 版本中均按 `VISIBLEAGGR` 名称注册。
