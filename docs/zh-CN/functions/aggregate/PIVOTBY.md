# PIVOTBY

**类别**：MS Excel 兼容函数 · **原生引入**：Microsoft 365

按行键和列键透视数据并聚合值。聚合器为文本(SUM、COUNT…)。

## 语法

```
=PIVOTBY(行字段, 列字段, 值, 函数, [字段标题], [行总计深度], [行排序], [列总计深度], [列排序], [筛选数组], [基准])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 行字段 | 必需 | 行键列(可多列) |
| 列字段 | 必需 | 列键列(可多列) |
| 值 | 必需 | 要聚合的值列(可多列) |
| 函数 | 必需 | 聚合器文本: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 扩展(参数=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 之一 |
| 字段标题 | 可选 | 0 无·隐藏, 1 有·隐藏, 2 无·显示, 3 有·显示 (默认自动) |
| 行总计深度 | 可选 | 0 无, 1 总计(默认), 2 总计+小计; 负数=顶部 |
| 行排序 | 可选 | 行排序: 列号，负数降序 |
| 列总计深度 | 可选 | 0 无, 1 总计列(默认), 2 总计+小计列; 负数=左侧 |
| 列排序 | 可选 | ±字段号(可多个)，负数降序 |
| 筛选数组 | 可选 | 每行 TRUE/FALSE |
| 基准 | 可选 | PERCENTOF 分母 0~2(默认 0 总计); 扩展聚合器的参数: TEXTJOIN 分隔符, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 返回

以二维数组溢出返回行键×列键交叉汇总表，每个列字段对应一行标题行，默认带总计行/列；row_fields·col_fields·values 与原生相同，各可为多列。聚合器·选项值无效或 PERCENTOF 的 relative_to 超出 0~2 时返回 #VALUE!；没有通过筛选的行时返回 #N/A；无数据的交叉单元格（小计行/列中亦然）为空字符串。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | 行×列求和透视 |
| `=PIVOTBY({"a";"a";"b"},{"x","p";"x","q";"y","p"},{10;20;30},"SUM")` | {"","x","x","y","Total";"","p","q","p","";"a",10,20,"",30;"b","","",30,30;"Total",10,20,30,60} | 两个列字段 → 两行标题 |

## 备注

- function 参数以文本而非原生的 lambda 指定；支持的聚合器为 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 共 16 种（与原生相同），并作为 EGTools 扩展支持 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) — 扩展聚合器的参数（TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE 的 p、QUARTILE 0~4）通过 relative_to 参数传入（其他名称返回 #VALUE!）。
- row_fields·col_fields·values 各可为多列（与原生相同）：每个列字段增加一行标题行，每个列组下按 values 的列数嵌套值列。值为 2 列以上时，总计·小计“列”的数据单元格也会按各值列聚合填充 — 原生会将这些单元格留空，因此仅此处与原生输出不同（有意差异）。field_headers 为 2·3（显示模式）时，标题块依次为：以 ", " 连接列字段名的一行、列键行、行字段名与值名称行。
- row_total_depth·col_total_depth 的 2 会按第一个行/列字段的块添加小计行/列（该轴需 2 个以上字段，否则按 1 处理）；负数将总计置于顶部/左侧。row_sort_order·col_sort_order 接受 ±字段号（可为数组）；排序始终保持字段层级顺序，仅改变指定字段的方向。
- 本实现中 PERCENTOF 的 relative_to 仅支持 0（按列总计，默认）/1（按行总计）/2（按总计），3·4（按父级）返回 #VALUE!。总计行/列的标签遵循 UI 语言（中文 UI：总计，英语 UI：Total/Grand Total）。
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `PIVOTBY` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.PIVOTBY`。
