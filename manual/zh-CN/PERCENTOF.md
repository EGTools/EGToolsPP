# PERCENTOF

**类别**：MS Excel 兼容函数 · **原生引入**：Microsoft 365

返回子集之和占全部数据之和的比例。

## 语法

```
=PERCENTOF(data_subset, data_all)
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| data_subset | 必需 | 数据子集 |
| data_all | 必需 | 全部数据 |

## 返回

以标量数字返回子集数据之和除以全部数据之和的比例。缺少参数时返回 #VALUE!；全部之和为 0 或全部数据中没有数字时返回 #DIV/0!。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=PERCENTOF({10;20},{10;20;30;40})` | 0.3 | 子集和÷全部和的比例 |
| `=PERCENTOF(50,200)` | 0.25 | 标量值的比例 |

## 备注

- 非数字的值在求和时被忽略。
- 在 GROUPBY/PIVOTBY 的 function 参数中指定 "PERCENTOF" 可求各组占比。
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `PERCENTOF` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.PERCENTOF`。
