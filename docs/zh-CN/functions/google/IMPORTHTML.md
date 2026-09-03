# IMPORTHTML

**类别**：Google 表格兼容函数

从网页中提取表格(table)或列表(list)。

## 语法

```
=IMPORTHTML(网址, [类型], [序号], [区域设置])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 网址 | 必需 | 网页 URL |
| 类型 | 可选 | "table"(默认) 或 "list"(ul/ol) |
| 序号 | 可选 | 提取第几个元素，从 1 开始(默认 1) |
| 区域设置 | 可选 | 数字解析区域，如 "de-DE" |

## 返回

溢出返回网页的表格（二维数组）或列表（N×1 数组），数字文本自动转换。缺少 URL、search 不是 "table"/"list" 或 index 小于 1 时返回 #VALUE!，下载失败或该 index 的表格/列表不存在时返回 #N/A。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=IMPORTHTML("https://en.wikipedia.org/wiki/Microsoft_Excel","table",1)` |  | 取决于页面内容 |

## 备注

- 先移除 <script> 再用 MSHTML 解析，因此无法获取由脚本生成的动态内容。
- 表格的 rowspan/colspan 规范化为网格，值只写入左上角单元格。
- 因使用 COM 被排除在多线程重算之外，且需要互联网连接。
- url指定为数组时返回 #VALUE! — 多条数据请将公式按行复制使用。
- 支持：Excel 2010+。在所有 Excel 版本中均按 `IMPORTHTML` 名称注册。
