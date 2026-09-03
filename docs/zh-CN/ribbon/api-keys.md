# 管理 API 密钥

**位置**：功能区 `EGTools` 选项卡 → `EGTools` 组 → 拆分按钮下拉菜单

用于注册/删除公共 API 函数·功能区功能所用 API 密钥的对话框。

列表中有 **3 个**服务。

| 服务 | 用途 | 相关功能 |
|---|---|---|
| juso | 道路名地址搜索 | [SEARCHADDRESS](../functions/korea/SEARCHADDRESS.md) |
| data.go.kr | 节假日·营业登记状态 | [KOREANHOLIDAYS](../functions/korea/KOREANHOLIDAYS.md) · [日历/日程表](calendar.md#节假日说明) · [BRNSTATUS](../functions/korea/BRNSTATUS.md) |
| vworld | 地图/坐标 | [GEOSEARCH](../functions/korea/GEOSEARCH.md) · [GEOCODER](../functions/korea/GEOCODER.md) · [GEOADDRESS](../functions/korea/GEOADDRESS.md) · [ADDRESSMAP](../functions/korea/ADDRESSMAP.md) |

- 选择服务后，密钥是否已注册**仅以掩码（●●●●●●）** 显示。
  已保存密钥的原文不会再次显示。
- 密钥**保存在用户账户中**，**不会留在工作簿或公式里**。
- 与**在函数最后一个参数直接传入密钥**的既有方式共存
  — 以参数传入的密钥会自动保存到同一存储位置，之后即可省略。
- **密钥不随 EGTools++ 分发。** 需要用户在各服务（juso.go.kr、
  data.go.kr、vworld.kr）自行申领。

## data.go.kr 密钥只有一个

公共数据门户（data.go.kr）**每个账户只发放一个**认证密钥，
申请使用的所有服务共用该密钥。因此 EGTools++ 也不按服务细分，
统一以 `data.go.kr` 一栏管理。

- 这一个密钥由 [KOREANHOLIDAYS](../functions/korea/KOREANHOLIDAYS.md)（特日信息）、
  [日历/日程表](calendar.md#节假日说明)的临时节假日合并、
  [BRNSTATUS](../functions/korea/BRNSTATUS.md)（国税厅营业登记状态）**共同使用**。
- 旧版本中**为营业登记单独注册的密钥**会在首次使用时
  **自动迁移**到统一栏位 — 用户无需重新注册或整理。
- 门户显示的编码（Encoding）/解码（Decoding）密钥粘贴哪个都可以。
  必要时内部会进行编码后使用。
- 函数查询失败并**确认密钥被拒时，会自动删除已保存的密钥**
  并返回重新申领指引。请申领新密钥后重新注册。
