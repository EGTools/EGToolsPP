# GOOGLETRANSLATE

**类别**：Google 表格兼容函数

使用 Google 翻译服务翻译文本。

## 语法

```
=GOOGLETRANSLATE(文本, [源语言], [目标语言])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 文本 | 必需 | 要翻译的文本(或区域) |
| 源语言 | 可选 | 源语言代码，如 "en"(默认 "auto" 自动检测) |
| 目标语言 | 可选 | 结果语言代码(默认: 界面语言) |

## 返回

以标量返回翻译后的文本。HTTP 请求本身失败时返回 #VALUE!；服务错误·页面格式变化·语言代码无效·超过 5000 字符时以 "ERROR:…" 文本返回；输入空文本返回空字符串。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=GOOGLETRANSLATE("Hello","en","ko")` |  | 取决于翻译结果 |

## 备注

- 抓取非官方的 translate.google.com/m 页面 — Google 更改格式后可能失效，且需要互联网连接。
- source_language 默认值为 auto，target_language 默认值为加载项 UI 语言。
- 区域/数组输入以换行连接后一次性翻译，超过 5000 字符时返回 "ERROR:Exceed max length"。
- 支持：Excel 2010+。在所有 Excel 版本中均按 `GOOGLETRANSLATE` 名称注册。
