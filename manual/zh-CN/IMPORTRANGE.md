# IMPORTRANGE

**类别**：Google 表格兼容函数

从共享的 Google 表格导入数据。

## 语法

```
=IMPORTRANGE(表格网址, [区域地址])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 表格网址 | 必需 | Google 表格 URL — 必须已启用共享 |
| 区域地址 | 可选 | 要导入的区域，如 "A1:C10"(默认: 整个工作表) |

## 返回

将共享 Google 表格的值以二维数组溢出返回，数字·日期文本自动转换。URL 为空或不是 /d/<ID> 格式时返回 #VALUE!，下载失败·未设置共享（HTML 响应）·空工作表·区域超出数据范围时返回 #N/A。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1AbCdEf.../edit#gid=0","A1:C10")` |  | 取决于表格内容 |

## 备注

- 只能导入已开启链接共享的 Google 表格 — 未共享的文档会返回登录页（HTML）而得到 #N/A。
- 工作表通过 URL 的 gid 选择（默认 gid=0）。range_address 中的工作表名部分被忽略。
- 使用 TSV 导出端点，因此仅获取值，不保留格式。
- 支持：Excel 2010+。在所有 Excel 版本中均按 `IMPORTRANGE` 名称注册。
