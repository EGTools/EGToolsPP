# 管理 API 金鑰

**位置**: 功能區 `EGTools` 索引標籤 → `EGTools` 群組 → 分割按鈕下拉式選單

用來註冊/刪除公共 API 函數與功能區功能所用 API 金鑰的對話方塊。

清單中有 **3 個**服務。

| 服務 | 用途 | 相關功能 |
|---|---|---|
| juso | 道路名地址搜尋 | [SEARCHADDRESS](../functions/korea/SEARCHADDRESS.md) |
| data.go.kr | 節假日、營業登記狀態 | [KOREANHOLIDAYS](../functions/korea/KOREANHOLIDAYS.md) · [行事曆/日程表](calendar.md#節假日說明) · [BRNSTATUS](../functions/korea/BRNSTATUS.md) |
| vworld | 地圖/座標 | [GEOSEARCH](../functions/korea/GEOSEARCH.md) · [GEOCODER](../functions/korea/GEOCODER.md) · [GEOADDRESS](../functions/korea/GEOADDRESS.md) · [ADDRESSMAP](../functions/korea/ADDRESSMAP.md) |

- 選擇服務後，金鑰的註冊狀態**只以遮罩（●●●●●●）** 顯示。
  已儲存的金鑰原文不會再次顯示。
- 金鑰**儲存於使用者帳戶**，**不會留在活頁簿或公式中**。
- 與**在函數的最後一個引數直接傳入金鑰**的既有方式並存
  — 以引數傳入的金鑰會自動儲存到同一個儲存區，之後即可省略。
- **金鑰不隨 EGTools++ 一併發佈。** 須由使用者自行向各服務（juso.go.kr、
  data.go.kr、vworld.kr）申請核發。

## data.go.kr 金鑰只有一個

公共資料入口網站 (data.go.kr) **每個帳戶只核發一把認證金鑰**，所有申請
使用的服務都共用這把金鑰。因此 EGTools++ 也不按服務區分，統一以一個
`data.go.kr` 欄位管理。

- 這一把金鑰由 [KOREANHOLIDAYS](../functions/korea/KOREANHOLIDAYS.md)（特定日資訊）、
  [行事曆/日程表](calendar.md#節假日說明)的臨時節假日合併、
  [BRNSTATUS](../functions/korea/BRNSTATUS.md)（國稅廳營業登記狀態）**共同使用**。
- 舊版本中**另外為營業登記註冊的金鑰**會在首次使用時
  **自動移轉**到整合後的欄位 — 使用者不需重新註冊或
  整理。
- 入口網站顯示的編碼 (Encoding)/解碼 (Decoding) 金鑰貼上任一種
  皆可。必要時會在內部編碼後使用。
- 函數查詢失敗且**確認金鑰遭拒時，會自動刪除已儲存的
  金鑰**並傳回重新核發的指引。請重新申請金鑰後再次註冊。
