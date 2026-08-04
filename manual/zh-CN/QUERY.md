# QUERY

**类别**：Google 表格兼容函数

对数据执行 SQL 查询(表名: EGTOOLS)。

## 语法

```
=QUERY(数据, 查询语句, [标题行数])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 数据 | 必需 | 要查询的区域或数组 |
| 查询语句 | 必需 | SQL 语句，如 "SELECT F1, SUM(F2) FROM EGTOOLS GROUP BY F1" |
| 标题行数 | 可选 | 标题行数: N≥1 时顶部 N 行为字段名，0 无(F1, F2, …)，-1/省略 自动检测 |

## 返回

将 SQL 执行结果以二维数组（仅值行，无标题行）溢出返回。缺少 data/query 或查询为空时返回 #VALUE!，无结果行时返回 #N/A；ADO/ACE 提供程序错误·临时 CSV 写入失败以 "ERROR: …" 文本返回。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=QUERY(A1:C10,"SELECT F1, SUM(F3) FROM EGTOOLS GROUP BY F1",0)` |  | 取决于数据 |

## 备注

- 使用的不是 Google Query Language，而是基于 ACE 文本驱动程序 + ADO 的 SQL；表名为 EGTOOLS（查询中的 EGTOOLS 自动替换为 EGTOOLS#csv）。
- 列名为 headers 指定的标题行文本，没有则为 F1、F2、…（headers：N≥1=顶部 N 行为字段名，0=无，省略/-1=自动推断）。
- 将数据写入 %TEMP%\EGTOOLS.csv 后查询，因此需要 Microsoft ACE OLEDB（12.0/16.0）提供程序；因使用公共临时文件被排除在多线程重算之外。
- 支持：Excel 2010+。在所有 Excel 版本中均按 `QUERY` 名称注册。
