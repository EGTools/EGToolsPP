# REGEXTEST

**类别**：MS Excel 兼容函数 · **原生引入**：Microsoft 365

检验文本是否与正则表达式匹配。

## 语法

```
=REGEXTEST(文本, 模式, [区分大小写])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| 文本 | 必需 | 要检验的文本 |
| 模式 | 必需 | 正则模式(ECMAScript) |
| 区分大小写 | 可选 | 0 区分(默认)，1 忽略 |

## 返回

返回表示是否匹配的逻辑值 TRUE/FALSE（标量）。正则表达式无效时返回 #VALUE!。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=REGEXTEST("abc123","\d+")` | TRUE | 包含数字 |
| `=REGEXTEST("ABC","[a-z]+",1)` | TRUE | 忽略大小写 |
| `=REGEXTEST("abc","^\d")` | FALSE | 无匹配 |

## 备注

- 正则语法为 std::wregex 的 ECMAScript（与原生 365 的 PCRE2 可能略有差异）。
- 相关函数：REGEXEXTRACT、REGEXREPLACE
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `REGEXTEST` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.REGEXTEST`。
