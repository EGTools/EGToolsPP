# IMAGE

**類別**: MS Excel 相容函數 · **原生導入**: Excel 2024

將來自 URL 或檔案路徑的圖片插入儲存格。

## 語法

```
=IMAGE(source, [alt_text], [sizing], [height], [width])
```

## 引數

| 引數 | 必要 | 說明 |
|---|---|---|
| source | 必要 | 圖片 URL 或檔案路徑 |
| alt_text | 選用 | 替代（無障礙）文字（選用） |
| sizing | 選用 | 0 符合並保持比例, 1 填滿, 2 原始大小, 3 自訂（選用） |
| height | 選用 | sizing 為 3 時的高度（像素）（選用） |
| width | 選用 | sizing 為 3 時的寬度（像素）（選用） |

## 傳回

一律以純量傳回空字串（""），並在計算結束後插入符合呼叫儲存格（合併範圍）大小的浮動圖片圖案。source 為空時傳回 #VALUE!；無法確認呼叫儲存格位址時傳回 #REF!。

## 範例

| 公式 | 結果 | 說明 |
|---|---|---|
| `=IMAGE("https://www.google.com/favicon.ico")` |  | 視外部圖片而定 |

## 備註

- 並非像 365 原生那樣成為儲存格的值，而是在儲存格上插入浮動圖片圖案（EG.Image_工作表_儲存格），重算時會取代同名的既有圖案。
- alt_text·height·width 引數僅為相容而接受但會被忽略；resize_mode 為 0（預設）時維持長寬比例。
- 在具有原生 IMAGE 的較新 Excel 中會註冊為 EG.IMAGE。
- source指定為陣列時傳回 #VALUE!（每個儲存格一張圖片）— 多筆資料請將公式逐列複製使用。
- 支援: Excel 2010+。在沒有原生函數的舊版 Excel 中以 `IMAGE` 原名註冊（可直接替換），在已內建原生函數的新版 Excel 中則註冊為 `EG.IMAGE`。
