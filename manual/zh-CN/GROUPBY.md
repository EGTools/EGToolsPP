# GROUPBY

**类别**：MS Excel 兼容函数 · **原生引入**：Microsoft 365

按键对行分组并聚合值。聚合器为文本: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 扩展(参数=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC)。

## 语法

```
=GROUPBY(行字段, 值, 函数, [字段标题], [总计深度], [排序], [筛选数组], [字段关系], [基准值])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 行字段 | 必需 | 键列 |
| 值 | 必需 | 要聚合的值列 |
| 函数 | 必需 | 聚合器文本: SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF + 扩展(参数=relative_to): TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) 之一 |
| 字段标题 | 可选 | 0 无·隐藏, 1 有·隐藏, 2 无·生成显示, 3 有·显示 (默认自动) |
| 总计深度 | 可选 | 0 无, 1 总计(默认), 2 总计+小计; 负数=顶部 |
| 排序 | 可选 | 输出列号，负数降序，如 {2,-1} |
| 筛选数组 | 可选 | 每行 TRUE/FALSE |
| 字段关系 | 可选 | 兼容接受但忽略 |
| 基准值 | 可选 | 扩展聚合器参数: TEXTJOIN 分隔符, LARGE/SMALL k, PERCENTILE p(0~1), QUARTILE 0~4 |

## 返回

以溢出返回由分组键列和聚合值列组成的二维数组。缺少必需参数、键·值行数不一致、聚合器·选项值无效时返回 #VALUE!；没有通过筛选的行时返回 #N/A；AVERAGE·PERCENTOF 中分母为 0 的单元格为 #DIV/0!。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=GROUPBY({"a";"b";"a"},{10;20;30},"SUM")` | {"a",40;"b",20;"Total",60} | 按键求和并带总计行 |
| `=GROUPBY({"a";"b";"a"},{10;20;30},"COUNT",0,0)` | {"a",2;"b",1} | 仅计数不带总计 |

## 备注

- function 参数以文本而非原生的 lambda 指定；支持的聚合器为 SUM/AVERAGE/COUNT/COUNTA/MAX/MIN/PRODUCT/MEDIAN/MODE.SNGL/STDEV.S/STDEV.P/VAR.S/VAR.P/CONCAT/ARRAYTOTEXT/PERCENTOF 共 16 种（与原生相同），并作为 EGTools 扩展支持 TEXTJOIN/LARGE/SMALL/PERCENTILE·QUARTILE(.INC/.EXC) — 扩展聚合器的参数（TEXTJOIN 分隔符、LARGE/SMALL 的 k、PERCENTILE 的 p、QUARTILE 0~4）通过 relative_to 参数传入（其他名称返回 #VALUE!）。field_relationship 参数接受但被忽略。
- 总计·小计标签遵循 UI 语言（中文 UI：总计，英语 UI：Total/Grand Total）。小计行的标签为第一个键的值。
- 省略 field_headers 时，若首行全为文本且其下存在非文本值，则自动识别为标题。
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `GROUPBY` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.GROUPBY`。
