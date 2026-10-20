# 应用旧版兼容

**位置**：功能区 `EGTools` 选项卡 → `EGTools` 组 → 拆分按钮下拉菜单

在旧版 Excel 中打开由新版 Excel 编写的工作簿时，新函数会以 `_xlfn.XLOOKUP`
这样的损坏标记出现并产生 `#NAME?` 错误。此按钮将活动工作簿所有公式中的
现代/原生函数标记批量转换为 EGTools 兼容函数名，使公式在旧版 Excel 中
也能照常计算。

## 操作流程

1. **保存确认** — 因为会直接修改公式，仅能在已保存的工作簿中执行。
   如果尚未保存，会先询问是否保存后再继续（拒绝/取消则中止）。
2. **预扫描** — 以只读方式遍历全部工作表的公式和定义的名称，
   按函数汇总转换对象的数量并显示。没有转换对象时提示后结束。
3. **备份** — 选择继续后，会在转换前自动保存备份文件
   （与工作簿同一文件夹，`原名_egbak_<时间>.<扩展名>`）。备份失败则不转换。
4. **转换与结果** — 汇总显示转换位置数和备份文件路径。

## 转换规则

| 对象 | 转换 |
|---|---|
| 一般兼容函数 | `_xlfn.函数` / `_xlfn._xlws.函数` / `EG.函数` → `函数`（EGTools 直接替换名）。冲突名称 SORT/FILTER/LET 转换为 `xSORT`/`xFILTER`/`xLET` |
| IMAGE | `_xlfn.IMAGE` / `EG.IMAGE` → `IMAGE` |
| GROUPBY / PIVOTBY | 将函数名改为 `GROUPBY`/`PIVOTBY`，并给聚合器参数加引号（`SUM` 或 `_xleta.SUM` → `"SUM"`；仅限原生支持的 16 种聚合器，LAMBDA 等则不改动该公式）。即使后面跟有可选参数也会转换（参数顺序与原生相同） |
| `_xlpm.*`（LET/LAMBDA 参数） | 新版 Excel 保存 LET/LAMBDA **参数名**的内部前缀。在旧版中会显示为 `=_xlfn.LET(_xlpm.x, …)`，而含 `_xlpm.` 的公式会被 Excel 拒绝重写，因此去除该前缀（`_xlpm.x` → `x`）。在转换汇总中计入 `_xlpm.*` 项 |
| `_xll.*` / `_xludf.*`（加载项 UDF 保存前缀） | 去除在未注册 EGTools 函数的环境中打开过的工作簿（尤其是**定义的名称**）中残留的前缀，重新连接到已注册的 EGTools 函数（`_xll.xSORT(…)` → `xSORT(…)`）。不改动其他加载项的函数。在转换汇总中计入 `_xll.*` 项 |
| `_xlfn._ARRAY` / `_xlfn._ARRAYROW`（单元格内数组包装） | Microsoft 365 测试版把溢出结果装进单个单元格时保存的内部函数。在没有原生函数的主机中会变成 `#NAME?`，因此改写为以文本表示法输出相同值的 `ARRAYTOTEXT(…,1)`：`_xlfn._ARRAYROW(a)` → `a`、`_xlfn._ARRAYROW(a,b,…)` → `HSTACK(a,b,…)`、`_xlfn._ARRAY(r)` → `ARRAYTOTEXT(r,1)`、`_xlfn._ARRAY(r1,r2,…)` → `ARRAYTOTEXT(VSTACK(r1,r2,…),1)`。例：`=_xlfn._ARRAY(_xlfn._ARRAYROW(FILTER(A:A,B:B=1)))` → `=ARRAYTOTEXT(FILTER(A:A,B:B=1),1)`。结果文本与[列表/数组单元格 → 文本](list-to-text.md)的严格表示法（`{"a",1;"b",2}`）相同，HAS/HASANY/HASALL/FLATTEN 可直接读取。由内向外递归处理，嵌套的 `_ARRAY` 会成为外层文本中带引号的文本项。在转换汇总中计入 `_ARRAY`/`_ARRAYROW` 项 |
| 分阶段推送的函数（HAS/HASANY/HASALL/FLATTEN、IMPORTTEXT/IMPORTCSV） | 若此 Excel 确实没有原生函数（运行时确认），则把 `_xlfn.HAS` / 不带前缀的 `HAS` 改为**此主机上已注册的 EG. 名称**，如 `EG.HAS`（避免在 Microsoft 365 非测试版通道中出现 `#NAME?`）。在 2024 及更早的旧版中仍按原样使用不带前缀的名称（`HAS`）。已改为 EG. 名称的文档之后在旧版中再次转换时会变为不带前缀的名称 |

字符串字面量内部或其他标识符的一部分不会被更改（按标识符为单位替换）。

## 备注

- 要还原时，请执行[还原为内置函数](restore-native.md)或打开备份文件。
- 转换对象函数列表与加载项的函数注册列表按单一来源管理，始终保持一致。
- 在测试版中创建的列表/数组单元格本身不由此命令处理。为避免引用单元格内列表的 HAS 等公式变成
  `#VALUE!`，请**先执行[列表/数组单元格 → 文本](list-to-text.md)**，再执行此命令。
