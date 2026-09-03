# EGTools++ 函數手冊（繁體中文）

**EGTools++** 是單一檔案的 `.xll` 增益集,讓舊版 Excel(2010 及以上)也能使用 Microsoft 365 的新版函數(僅限 Windows,無需安裝程式或系統管理員權限)。它在任何版本上都以相同方式提供 XLOOKUP、FILTER、TEXTSPLIT、GROUPBY、正規表示式等函數:主機沒有原生函數時按原名註冊(直接替換),已有原生函數時以 `EG.` 前綴註冊。另附實用功能區(可見儲存格複製、合併列印、行事曆產生等)。

下載:[GitHub Releases](https://github.com/EGTools/EGToolsPP/releases) · 原始碼與回報:[GitHub 儲存庫](https://github.com/EGTools/EGToolsPP)

各函數的詳細說明。在 IntelliSense 工具提示中按一下函數名稱，即可開啟對應頁面。

## MS Excel 相容函數 (69)

### Excel 2013 (28)

- [ARABIC](functions/excel-2013/ARABIC.md)
- [BASE](functions/excel-2013/BASE.md)
- [BITAND](functions/excel-2013/BITAND.md)
- [BITLSHIFT](functions/excel-2013/BITLSHIFT.md)
- [BITOR](functions/excel-2013/BITOR.md)
- [BITRSHIFT](functions/excel-2013/BITRSHIFT.md)
- [BITXOR](functions/excel-2013/BITXOR.md)
- [CEILING.MATH](functions/excel-2013/CEILING.MATH.md)
- [COMBINA](functions/excel-2013/COMBINA.md)
- [DAYS](functions/excel-2013/DAYS.md)
- [DECIMAL](functions/excel-2013/DECIMAL.md)
- [ENCODEURL](functions/excel-2013/ENCODEURL.md)
- [FILTERXML](functions/excel-2013/FILTERXML.md)
- [FLOOR.MATH](functions/excel-2013/FLOOR.MATH.md)
- [FORMULATEXT](functions/excel-2013/FORMULATEXT.md)
- [IFNA](functions/excel-2013/IFNA.md)
- [ISFORMULA](functions/excel-2013/ISFORMULA.md)
- [ISOWEEKNUM](functions/excel-2013/ISOWEEKNUM.md)
- [NUMBERVALUE](functions/excel-2013/NUMBERVALUE.md)
- [PDURATION](functions/excel-2013/PDURATION.md)
- [PERMUTATIONA](functions/excel-2013/PERMUTATIONA.md)
- [RRI](functions/excel-2013/RRI.md)
- [SHEET](functions/excel-2013/SHEET.md)
- [SHEETS](functions/excel-2013/SHEETS.md)
- [UNICHAR](functions/excel-2013/UNICHAR.md)
- [UNICODE](functions/excel-2013/UNICODE.md)
- [WEBSERVICE](functions/excel-2013/WEBSERVICE.md)
- [XOR](functions/excel-2013/XOR.md)

### Excel 2016/2019 (6)

- [CONCAT](functions/excel-2016-2019/CONCAT.md)
- [IFS](functions/excel-2016-2019/IFS.md)
- [MAXIFS](functions/excel-2016-2019/MAXIFS.md)
- [MINIFS](functions/excel-2016-2019/MINIFS.md)
- [SWITCH](functions/excel-2016-2019/SWITCH.md)
- [TEXTJOIN](functions/excel-2016-2019/TEXTJOIN.md)

### Excel 2021 (9)

- [FILTER](functions/excel-2021/FILTER.md)
- [LET](functions/excel-2021/LET.md)
- [RANDARRAY](functions/excel-2021/RANDARRAY.md)
- [SEQUENCE](functions/excel-2021/SEQUENCE.md)
- [SORT](functions/excel-2021/SORT.md)
- [SORTBY](functions/excel-2021/SORTBY.md)
- [UNIQUE](functions/excel-2021/UNIQUE.md)
- [XLOOKUP](functions/excel-2021/XLOOKUP.md)
- [XMATCH](functions/excel-2021/XMATCH.md)

### Excel 2024 (17)

- [ARRAYTOTEXT](functions/excel-2024/ARRAYTOTEXT.md)
- [CHOOSECOLS](functions/excel-2024/CHOOSECOLS.md)
- [CHOOSEROWS](functions/excel-2024/CHOOSEROWS.md)
- [DROP](functions/excel-2024/DROP.md)
- [EXPAND](functions/excel-2024/EXPAND.md)
- [HSTACK](functions/excel-2024/HSTACK.md)
- [IMAGE](functions/excel-2024/IMAGE.md)
- [TAKE](functions/excel-2024/TAKE.md)
- [TEXTAFTER](functions/excel-2024/TEXTAFTER.md)
- [TEXTBEFORE](functions/excel-2024/TEXTBEFORE.md)
- [TEXTSPLIT](functions/excel-2024/TEXTSPLIT.md)
- [TOCOL](functions/excel-2024/TOCOL.md)
- [TOROW](functions/excel-2024/TOROW.md)
- [VALUETOTEXT](functions/excel-2024/VALUETOTEXT.md)
- [VSTACK](functions/excel-2024/VSTACK.md)
- [WRAPCOLS](functions/excel-2024/WRAPCOLS.md)
- [WRAPROWS](functions/excel-2024/WRAPROWS.md)

### Microsoft 365 (9)

- [GROUPBY](functions/microsoft-365/GROUPBY.md)
- [IMPORTCSV](functions/microsoft-365/IMPORTCSV.md) — Microsoft 365 (2026 新增)
- [IMPORTTEXT](functions/microsoft-365/IMPORTTEXT.md) — Microsoft 365 (2026 新增)
- [PERCENTOF](functions/microsoft-365/PERCENTOF.md)
- [PIVOTBY](functions/microsoft-365/PIVOTBY.md)
- [REGEXEXTRACT](functions/microsoft-365/REGEXEXTRACT.md)
- [REGEXREPLACE](functions/microsoft-365/REGEXREPLACE.md)
- [REGEXTEST](functions/microsoft-365/REGEXTEST.md)
- [TRIMRANGE](functions/microsoft-365/TRIMRANGE.md)

## Google Sheets 相容函數 (7)

- [GOOGLETRANSLATE](functions/google/GOOGLETRANSLATE.md)
- [GTRS](functions/google/GTRS.md)
- [IMPORTDATA](functions/google/IMPORTDATA.md)
- [IMPORTFEED](functions/google/IMPORTFEED.md)
- [IMPORTHTML](functions/google/IMPORTHTML.md)
- [IMPORTRANGE](functions/google/IMPORTRANGE.md)
- [QUERY](functions/google/QUERY.md)

## EGTools 專用函數 (52)

### 查閱與參照 (3)

- [COMPARELIST](functions/eg-lookup/COMPARELIST.md)
- [ILOOKUP](functions/eg-lookup/ILOOKUP.md)
- [MVLOOKUP](functions/eg-lookup/MVLOOKUP.md)

### 資料與陣列 (4)

- [ADVANCEDFILTER](functions/eg-data/ADVANCEDFILTER.md)
- [EXPLODE](functions/eg-data/EXPLODE.md)
- [TEXTNUMSORT](functions/eg-data/TEXTNUMSORT.md)
- [UNPIVOT](functions/eg-data/UNPIVOT.md)

### 文字 (4)

- [STREXT](functions/eg-text/STREXT.md)
- [TEXTBETWEEN](functions/eg-text/TEXTBETWEEN.md)
- [TEXTREPLACE](functions/eg-text/TEXTREPLACE.md)
- [TRIMENDS](functions/eg-text/TRIMENDS.md)

### 日期與時間 (4)

- [MONTHBYWEEK](functions/eg-datetime/MONTHBYWEEK.md)
- [NETWORKHOUR](functions/eg-datetime/NETWORKHOUR.md)
- [TODATETIME](functions/eg-datetime/TODATETIME.md)
- [WEEKNUMOFMONTH](functions/eg-datetime/WEEKNUMOFMONTH.md)

### 資訊與公式 (4)

- [DIRFOLDER](functions/eg-info/DIRFOLDER.md)
- [EVAL](functions/eg-info/EVAL.md)
- [SHEETLIST](functions/eg-info/SHEETLIST.md)
- [TOTALPAGES](functions/eg-info/TOTALPAGES.md)

### 色彩與可見儲存格 (5)

- [COUNTIFCOLOR](functions/eg-visible/COUNTIFCOLOR.md)
- [DISPLAYCOLOR](functions/eg-visible/DISPLAYCOLOR.md)
- [ISVISIBLE](functions/eg-visible/ISVISIBLE.md)
- [SUMIFCOLOR](functions/eg-visible/SUMIFCOLOR.md)
- [VISIBLEAGGR](functions/eg-visible/VISIBLEAGGR.md)

### 圖像與條碼 (6)

- [BARCODE](functions/eg-image/BARCODE.md)
- [CODE128](functions/eg-image/CODE128.md)
- [DATAMATRIX](functions/eg-image/DATAMATRIX.md)
- [QRCODE](functions/eg-image/QRCODE.md)
- [READBARCODE](functions/eg-image/READBARCODE.md)
- [READGS1](functions/eg-image/READGS1.md)

### Web 與 API (3)

- [DECODEURL](functions/eg-web/DECODEURL.md)
- [JSONFILTER](functions/eg-web/JSONFILTER.md)
- [RESTAPI](functions/eg-web/RESTAPI.md)

### 韓國資料 (11)

- [ADDRESSMAP](functions/eg-korea/ADDRESSMAP.md)
- [BRNSTATUS](functions/eg-korea/BRNSTATUS.md)
- [EXRATE](functions/eg-korea/EXRATE.md)
- [GEOADDRESS](functions/eg-korea/GEOADDRESS.md)
- [GEOCODER](functions/eg-korea/GEOCODER.md)
- [GEOSEARCH](functions/eg-korea/GEOSEARCH.md)
- [HANTONUMBER](functions/eg-korea/HANTONUMBER.md)
- [KOREANHOLIDAYS](functions/eg-korea/KOREANHOLIDAYS.md)
- [SEARCHADDRESS](functions/eg-korea/SEARCHADDRESS.md)
- [TOLUNAR](functions/eg-korea/TOLUNAR.md)
- [TOSOLAR](functions/eg-korea/TOSOLAR.md)

### 品質與生產 (8)

- [BOMTREE](functions/eg-quality/BOMTREE.md)
- [CP](functions/eg-quality/CP.md)
- [CPK](functions/eg-quality/CPK.md)
- [SAMPLINGAC](functions/eg-quality/SAMPLINGAC.md)
- [SAMPLINGLABEL](functions/eg-quality/SAMPLINGLABEL.md)
- [SAMPLINGRE](functions/eg-quality/SAMPLINGRE.md)
- [SAMPLINGSIZE](functions/eg-quality/SAMPLINGSIZE.md)
- [SUMBYBOM](functions/eg-quality/SUMBYBOM.md)

## 功能區選單

功能區 `EGTools` 索引標籤提供的功能。完整配置請參閱[功能區選單總覽](ribbon/README.md)。

### 可見儲存格

- [可見儲存格 — 僅複製可見儲存格 · 全部複製 · 僅貼上值/貼上公式/全部貼上](ribbon/visible-cells.md)

### 合併/取消合併儲存格

- [合併/取消合併儲存格 — 合併連續相同值 · 保留值的合併 · 合併內容 · 逐列/逐欄合併 · 取消合併並填滿 · 拆分為列/欄](ribbon/merge-split.md)

### 圖片

- [圖片 — 插入圖片 · 依名稱插入 · 插入到表單 · 調整所選 · 全部調整 · 全部匯出](ribbon/pictures.md)

### 行事曆/日程表

- [行事曆/日程表 — 月曆 · 年曆A/B · 週 · 日A/B](ribbon/calendar.md)

### 表單

- [表單 — 製作標籤](ribbon/forms.md) · [郵件合併](ribbon/mailmerge.md)

### 工具

- [工具 — 刪除樣式 · 刪除名稱 · 整理錯誤 · 清理空文字 · 整理註解 · 變更框線色彩 · 描摹形狀 · UDF 固定為值 · 套用 CheckBox · 全部重算](ribbon/utilities.md)

### EGTools

- [套用舊版相容](ribbon/apply-compat.md)
- [還原為內建函數](ribbon/restore-native.md)
- [SMTP 設定](ribbon/smtp-settings.md)
- [管理 API 金鑰](ribbon/api-keys.md)

### 右鍵選單

- [儲存格右鍵選單 — 日期選擇 · 開啟連結](ribbon/context-menu.md)

