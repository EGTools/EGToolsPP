# PERMUTATIONA

**类别**：MS Excel 兼容函数 · **原生引入**：Excel 2013

返回可重复排列数（number^number_chosen）。

## 语法

```
=PERMUTATIONA(number, number_chosen)
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| number | 必需 | 项目总数 |
| number_chosen | 必需 | 每个排列的项目数 |

## 返回

返回可重复排列数 number^number_chosen（标量数字）。参数非数字或为负数时返回 #NUM! 错误。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=PERMUTATIONA(3,2)` | 9 | 3^2 |
| `=PERMUTATIONA(2,5)` | 32 | 2^5 |

## 备注

- 舍去小数部分。
- 所有参数指定为数组时按元素计算，以相同形状的数组溢出返回 — 标量重复使用，列向量×行向量按外积扩展，大小不匹配的元素为 #N/A，错误元素原样返回该错误。
- 相关函数：COMBINA
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `PERMUTATIONA` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.PERMUTATIONA`。
