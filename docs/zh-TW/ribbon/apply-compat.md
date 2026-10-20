# 套用舊版相容

**位置**: 功能區 `EGTools` 索引標籤 → `EGTools` 群組 → 分割按鈕下拉式選單

在最新 Excel 中製作的活頁簿若在舊版 Excel 開啟，新函數會變成 `_xlfn.XLOOKUP`
之類的損毀符記，並出現 `#NAME?` 錯誤。此按鈕會將作用中活頁簿所有公式裡的
現代/原生函數符記批次轉換為 EGTools 相容函數名稱，讓公式在舊版 Excel
也能照常計算。

## 操作流程

1. **儲存確認** — 因為會直接修改公式，只在已儲存的活頁簿上執行。
   若尚未儲存，會先詢問是否儲存再繼續（拒絕/取消則中止）。
2. **事前掃描** — 以唯讀方式掃描全部工作表的公式與名稱，
   依函數彙總顯示轉換對象的數量。若無對象則提示後結束。
3. **備份** — 選擇繼續後，轉換前會自動儲存備份檔案
   （與活頁簿相同的資料夾，`原檔名_egbak_<時間>.<副檔名>`）。備份失敗則不進行轉換。
4. **轉換與結果** — 彙總顯示轉換的位置數與備份檔案路徑。

## 轉換規則

| 對象 | 轉換 |
|---|---|
| 一般相容函數 | `_xlfn.函數` / `_xlfn._xlws.函數` / `EG.函數` → `函數`（EGTools 直接替換名稱）。衝突名稱 SORT/FILTER/LET 會轉換為 `xSORT`/`xFILTER`/`xLET` |
| IMAGE | `_xlfn.IMAGE` / `EG.IMAGE` → `IMAGE` |
| GROUPBY / PIVOTBY | 將函數名稱改為 `GROUPBY`/`PIVOTBY`，並為彙總函數引數加上引號（`SUM` 或 `_xleta.SUM` → `"SUM"`；僅限原生支援的 16 種彙總函數，LAMBDA 等則不更動該公式）。即使後面帶有選擇性引數也會轉換（引數順序與原生相同） |
| `_xlpm.*`（LET/LAMBDA 參數） | 這是最新 Excel 用來儲存 LET/LAMBDA **參數名稱**的內部前綴。在舊版中會顯示為 `=_xlfn.LET(_xlpm.x, …)`，且含 `_xlpm.` 的公式 Excel 會拒絕改寫，因此會移除此前綴（`_xlpm.x` → `x`）。在轉換摘要中計入 `_xlpm.*` 項目 |
| `_xll.*` / `_xludf.*`（增益集 UDF 儲存前綴） | 移除曾在未註冊 EGTools 函數的環境中開啟過的活頁簿（尤其是**名稱**）所殘留的前綴，重新連結到已註冊的 EGTools 函數（`_xll.xSORT(…)` → `xSORT(…)`）。不會更動其他增益集的函數。在轉換摘要中計入 `_xll.*` 項目 |
| `_xlfn._ARRAY` / `_xlfn._ARRAYROW`（儲存格內陣列包裝） | Microsoft 365 測試版把溢出結果裝進單一儲存格時儲存的內部函數。在沒有原生函數的主機上會變成 `#NAME?`，因此改寫為以文字表示法輸出相同值的 `ARRAYTOTEXT(…,1)`：`_xlfn._ARRAYROW(a)` → `a`、`_xlfn._ARRAYROW(a,b,…)` → `HSTACK(a,b,…)`、`_xlfn._ARRAY(r)` → `ARRAYTOTEXT(r,1)`、`_xlfn._ARRAY(r1,r2,…)` → `ARRAYTOTEXT(VSTACK(r1,r2,…),1)`。例：`=_xlfn._ARRAY(_xlfn._ARRAYROW(FILTER(A:A,B:B=1)))` → `=ARRAYTOTEXT(FILTER(A:A,B:B=1),1)`。結果文字與[清單/陣列儲存格 → 文字](list-to-text.md)的嚴格表示法（`{"a",1;"b",2}`）相同，HAS/HASANY/HASALL/FLATTEN 可直接讀取。由內向外遞迴處理，巢狀的 `_ARRAY` 會成為外層文字中帶引號的文字項目。在轉換摘要中計入 `_ARRAY`/`_ARRAYROW` 項目 |
| 分階段推出的函數（HAS/HASANY/HASALL/FLATTEN、IMPORTTEXT/IMPORTCSV） | 若此 Excel 確實沒有原生函數（執行時確認），會把 `_xlfn.HAS` / 不帶前綴的 `HAS` 改為**此主機上已註冊的 EG. 名稱**，例如 `EG.HAS`（避免在 Microsoft 365 非測試版通道出現 `#NAME?`）。在 2024 以下的舊版中仍照舊使用不帶前綴的名稱（`HAS`）。已改為 EG. 名稱的文件之後在舊版中再次轉換時會變回不帶前綴的名稱 |

字串常值內部或其他識別字的一部分不會被變更（以識別字為單位取代）。

## 備註

- 若要復原，請執行[還原為內建函數](restore-native.md)或開啟備份檔案。
- 轉換對象函數清單與增益集的函數註冊清單以單一來源管理，永遠保持一致。
- 在測試版中建立的清單/陣列儲存格本身不由此命令處理。為避免參照儲存格內清單的 HAS 等公式變成
  `#VALUE!`，請**先執行[清單/陣列儲存格 → 文字](list-to-text.md)**，再執行此命令。
