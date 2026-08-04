# FILTERXML

**类别**：MS Excel 兼容函数 · **原生引入**：Excel 2013

使用指定的 XPath 从 XML 内容返回数据。

## 语法

```
=FILTERXML(xml, xpath)
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| xml | 必需 | XML 文本 |
| xpath | 必需 | XPath 查询 |

## 返回

与 XPath 匹配的节点为 1 个时以标量返回其文本，多个时以 N×1 纵向文本数组溢出返回。XML/XPath 为空或 XML 解析·XPath 处理失败时返回 #VALUE!，没有匹配节点时返回 #N/A。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=FILTERXML("<r><a>1</a><a>2</a></r>","//a")` | {"1";"2"} | 所有 a 节点的文本 |
| `=FILTERXML(WEBSERVICE("https://example.com/rss.xml"),"//item/title")` |  | 取决于外部 XML |

## 备注

- 使用 MSXML6 解析，函数本身不访问网络 — 远程 XML 请与 WEBSERVICE 组合使用。
- 结果为节点的文本，始终以文本形式返回。
- 因使用 COM 而被排除在多线程重算之外（非线程安全）。
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `FILTERXML` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.FILTERXML`。
