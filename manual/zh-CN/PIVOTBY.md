# PIVOTBY

**类别**：MS Excel 兼容函数 · **原生引入**：Excel 2024 / Microsoft 365

按行键和列键透视数据并聚合值。聚合器为文本(SUM、COUNT…)。

## 语法

```
=PIVOTBY(行字段, 列字段, 值, 函数, [字段标题], [行总计深度], [行排序], [列总计深度], [列排序], [筛选数组], [基准])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 行字段 | 必需 | 行键列 |
| 列字段 | 必需 | 列键列 |
| 值 | 必需 | 要聚合的值 |
| 函数 | 必需 | 聚合器文本: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 扩展(参数=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 之一 |
| 字段标题 | 可选 | 0 无·隐藏, 1 有·隐藏, 2 无·显示, 3 有·显示 (默认自动) |
| 行总计深度 | 可选 | 0 无, 1 总计(默认); 负数=顶部 |
| 行排序 | 可选 | 行排序: 列号，负数降序 |
| 列总计深度 | 可选 | 0 无, 1 总计列(默认); 负数=左侧 |
| 列排序 | 可选 | 1 升序, -1 降序 |
| 筛选数组 | 可选 | 每行 TRUE/FALSE |
| 基准 | 可选 | PERCENTOF 分母 0~2(默认 0 总计); 扩展聚合器的参数: TEXTJOIN 分隔符, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 返回

以二维数组溢出返回包含列键标题行的行键×列键交叉汇总表。row_fields·col_fields·values 各超过 1 列或 relative_to 超出 0~2 时返回 #VALUE!；没有通过筛选的行时返回 #N/A；无数据的交叉单元格为空字符串。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=PIVOTBY({"a";"b";"a"},{"x";"x";"y"},{10;20;30},"SUM")` | {"","x","y","Total";"a",10,30,40;"b",20,"",20;"Total",30,30,60} | 行×列求和透视 |

## 备注

- row_fields·col_fields·values 各仅支持 1 列（不支持多列嵌套）。支持的聚合器为 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 共 16 种（与原生相同），并作为 EGTools 扩展支持 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) — 扩展聚合器的参数（TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE 的 p、QUARTILE 0~4）通过 relative_to 参数传入（其他名称返回 #VALUE!）。
- 本实现中 PERCENTOF 的 relative_to 仅支持 0（按列总计，默认）/1（按行总计）/2（按总计），3·4（按父级）返回 #VALUE!。
- 总计行/列的标签遵循 UI 语言（中文 UI：总计，英语 UI：Total）。
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `PIVOTBY` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.PIVOTBY`。
