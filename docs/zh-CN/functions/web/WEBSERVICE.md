# WEBSERVICE

**类别**：MS Excel 兼容函数 · **原生引入**：Excel 2013

返回指定 URL 处 Web 服务的响应（HTTP GET）。

## 语法

```
=WEBSERVICE(url)
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| url | 必需 | 要请求的 URL |

## 返回

以文本标量返回按 UTF-8 解释的 HTTP(S) GET 响应正文。URL 格式错误或连接·请求失败时返回 #VALUE!（连接 5 秒、接收响应 15 秒超时）。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=WEBSERVICE("https://api.frankfurter.app/latest")` |  | 取决于外部响应 |

## 备注

- 通过 WinHTTP GET 工作，需要互联网连接。
- 响应正文始终按 UTF-8 解释。
- 注册为线程安全，在多线程重算中并行执行。
- url指定为数组时返回 #VALUE! — 多条数据请将公式按行复制使用。
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `WEBSERVICE` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.WEBSERVICE`。
