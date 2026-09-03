# IMAGE

**类别**：MS Excel 兼容函数 · **原生引入**：Excel 2024

将来自 URL 或文件路径的图片插入单元格。

## 语法

```
=IMAGE(source, [alt_text], [sizing], [height], [width])
```

## 参数

| 参数 | 必需 | 说明 |
|---|---|---|
| source | 必需 | 图片 URL 或文件路径 |
| alt_text | 可选 | 替代（无障碍）文本（可选） |
| sizing | 可选 | 0 适应并保持比例, 1 填充, 2 原始大小, 3 自定义（可选） |
| height | 可选 | sizing 为 3 时的高度（像素）（可选） |
| width | 可选 | sizing 为 3 时的宽度（像素）（可选） |

## 返回

始终以标量返回空字符串（""），并在计算结束后按调用单元格（合并区域）大小插入浮动图片形状。source 为空时返回 #VALUE!，无法确认调用单元格地址时返回 #REF!。

## 示例

| 公式 | 结果 | 说明 |
|---|---|---|
| `=IMAGE("https://www.google.com/favicon.ico")` |  | 取决于外部图片 |

## 备注

- 不像 365 原生那样把图片作为单元格的值，而是在单元格上方插入浮动图片形状（EG.Image_工作表_单元格），重算时替换同名的现有形状。
- alt_text·height·width 参数为兼容而接受但被忽略；resize_mode 为 0（默认）时保持纵横比。
- 在具有原生 IMAGE 的新版 Excel 中注册为 EG.IMAGE。
- source指定为数组时返回 #VALUE!（每个单元格一张图片）— 多条数据请将公式按行复制使用。
- 支持：Excel 2010+。在没有原生函数的旧版本中按 `IMAGE` 原名注册（直接替换），在具有原生函数的新版 Excel 中注册为 `EG.IMAGE`。
