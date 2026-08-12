# IMPORTFEED

**类别**：Google 表格兼容函数

导入 RSS/ATOM 订阅源。

## 语法

```
=IMPORTFEED(网址, [查询], [标题行], [条目数])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 网址 | 必需 | 订阅源 URL |
| 查询 | 可选 | "items"(默认)、"feed" 或字段: title/summary/description/author/url/created |
| 标题行 | 可选 | TRUE 时包含标题行 |
| 条目数 | 可选 | 返回的条目数(默认: 全部) |

## 返回

将订阅源条目（或订阅源信息·单一字段）以二维数组溢出返回。缺少 URL 或 query 字段不受支持时返回 #VALUE!，下载失败·XML 解析失败·无条目时返回 #N/A。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=IMPORTFEED("https://blog.example.com/rss","items",TRUE,5)` |  | 取决于订阅源内容 |

## 备注

- 自动判别 RSS 2.0 与 Atom。query "items"（默认）返回 Title/Summary/URL/Author/Created 5 列。
- created 字段的 RFC822/ISO8601 日期转换为 Excel 序列值（忽略时区偏移）。
- 因使用 COM（MSXML）被排除在多线程重算之外，且需要互联网连接。
- url指定为数组时返回 #VALUE! — 多条数据请将公式按行复制使用。
- 支持：Excel 2010+。在所有 Excel 版本中均按 `IMPORTFEED` 名称注册。
