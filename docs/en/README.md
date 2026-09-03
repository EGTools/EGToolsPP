# EGTools++ Function Manual (English)

**EGTools++** is a single-file `.xll` add-in that brings modern Microsoft 365 functions to legacy Excel (2010 and later; Windows only, no installer or admin rights). It provides XLOOKUP, FILTER, TEXTSPLIT, GROUPBY, REGEX functions and more, identically on every version — registered under the original name as a drop-in where the host lacks the native function, or with an `EG.` prefix where it exists. A ribbon of practical utilities (visible-cells copy, mail merge, calendar builder) is included.

Download: [GitHub Releases](https://github.com/EGTools/EGToolsPP/releases) · Source & support: [GitHub repository](https://github.com/EGTools/EGToolsPP)

Per-function reference. Clicking a function name in the IntelliSense tooltip opens the matching page.

## MS Excel compatibility functions (69)

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
- [IMPORTCSV](functions/microsoft-365/IMPORTCSV.md) — Microsoft 365 (new in 2026)
- [IMPORTTEXT](functions/microsoft-365/IMPORTTEXT.md) — Microsoft 365 (new in 2026)
- [PERCENTOF](functions/microsoft-365/PERCENTOF.md)
- [PIVOTBY](functions/microsoft-365/PIVOTBY.md)
- [REGEXEXTRACT](functions/microsoft-365/REGEXEXTRACT.md)
- [REGEXREPLACE](functions/microsoft-365/REGEXREPLACE.md)
- [REGEXTEST](functions/microsoft-365/REGEXTEST.md)
- [TRIMRANGE](functions/microsoft-365/TRIMRANGE.md)

## Google Sheets compatibility functions (7)

- [GOOGLETRANSLATE](functions/google/GOOGLETRANSLATE.md)
- [GTRS](functions/google/GTRS.md)
- [IMPORTDATA](functions/google/IMPORTDATA.md)
- [IMPORTFEED](functions/google/IMPORTFEED.md)
- [IMPORTHTML](functions/google/IMPORTHTML.md)
- [IMPORTRANGE](functions/google/IMPORTRANGE.md)
- [QUERY](functions/google/QUERY.md)

## EGTools-only functions (52)

### Lookup & Reference (3)

- [COMPARELIST](functions/eg-lookup/COMPARELIST.md)
- [ILOOKUP](functions/eg-lookup/ILOOKUP.md)
- [MVLOOKUP](functions/eg-lookup/MVLOOKUP.md)

### Data & Arrays (4)

- [ADVANCEDFILTER](functions/eg-data/ADVANCEDFILTER.md)
- [EXPLODE](functions/eg-data/EXPLODE.md)
- [TEXTNUMSORT](functions/eg-data/TEXTNUMSORT.md)
- [UNPIVOT](functions/eg-data/UNPIVOT.md)

### Text (4)

- [STREXT](functions/eg-text/STREXT.md)
- [TEXTBETWEEN](functions/eg-text/TEXTBETWEEN.md)
- [TEXTREPLACE](functions/eg-text/TEXTREPLACE.md)
- [TRIMENDS](functions/eg-text/TRIMENDS.md)

### Date & Time (4)

- [MONTHBYWEEK](functions/eg-datetime/MONTHBYWEEK.md)
- [NETWORKHOUR](functions/eg-datetime/NETWORKHOUR.md)
- [TODATETIME](functions/eg-datetime/TODATETIME.md)
- [WEEKNUMOFMONTH](functions/eg-datetime/WEEKNUMOFMONTH.md)

### Information & Formulas (4)

- [DIRFOLDER](functions/eg-info/DIRFOLDER.md)
- [EVAL](functions/eg-info/EVAL.md)
- [SHEETLIST](functions/eg-info/SHEETLIST.md)
- [TOTALPAGES](functions/eg-info/TOTALPAGES.md)

### Color & Visible Cells (5)

- [COUNTIFCOLOR](functions/eg-visible/COUNTIFCOLOR.md)
- [DISPLAYCOLOR](functions/eg-visible/DISPLAYCOLOR.md)
- [ISVISIBLE](functions/eg-visible/ISVISIBLE.md)
- [SUMIFCOLOR](functions/eg-visible/SUMIFCOLOR.md)
- [VISIBLEAGGR](functions/eg-visible/VISIBLEAGGR.md)

### Images & Barcodes (6)

- [BARCODE](functions/eg-image/BARCODE.md)
- [CODE128](functions/eg-image/CODE128.md)
- [DATAMATRIX](functions/eg-image/DATAMATRIX.md)
- [QRCODE](functions/eg-image/QRCODE.md)
- [READBARCODE](functions/eg-image/READBARCODE.md)
- [READGS1](functions/eg-image/READGS1.md)

### Web & API (3)

- [DECODEURL](functions/eg-web/DECODEURL.md)
- [JSONFILTER](functions/eg-web/JSONFILTER.md)
- [RESTAPI](functions/eg-web/RESTAPI.md)

### Korean Data (11)

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

### Quality & Manufacturing (8)

- [BOMTREE](functions/eg-quality/BOMTREE.md)
- [CP](functions/eg-quality/CP.md)
- [CPK](functions/eg-quality/CPK.md)
- [SAMPLINGAC](functions/eg-quality/SAMPLINGAC.md)
- [SAMPLINGLABEL](functions/eg-quality/SAMPLINGLABEL.md)
- [SAMPLINGRE](functions/eg-quality/SAMPLINGRE.md)
- [SAMPLINGSIZE](functions/eg-quality/SAMPLINGSIZE.md)
- [SUMBYBOM](functions/eg-quality/SUMBYBOM.md)

## Ribbon Menu

Features on the `EGTools` ribbon tab. See the [ribbon overview](ribbon/README.md) for the full layout.

### Visible Cells

- [Visible Cells — Copy Visible Cells, Copy Range, Paste Values/Formulas/All](ribbon/visible-cells.md)

### Merge/Unmerge Cells

- [Merge/Unmerge Cells — Merge Same Values, Merge Keeping Values, Merge Contents, Merge Rows/Columns, Unmerge and Fill, Split to Rows/Columns](ribbon/merge-split.md)

### Image/Picture

- [Image/Picture — Insert Image, Insert by Name, Insert into Form, Fit Selected, Fit All, Export All](ribbon/pictures.md)

### Calendar

- [Calendar — Monthly, Annual A/B, Weekly, Daily A/B](ribbon/calendar.md)

### Forms

- [Forms — Make Labels](ribbon/forms.md) · [Mail Merge](ribbon/mailmerge.md)

### Utilities

- [Utilities — Delete Styles, Delete Names, Clear Errors, Clear Empty Text, Arrange Notes, Change Border Colors, Draw Free Form, Freeze UDF Values, Apply CheckBox, Recalculate All](ribbon/utilities.md)

### EGTools

- [Apply Legacy Compatibility](ribbon/apply-compat.md)
- [Restore Native Functions](ribbon/restore-native.md)
- [SMTP Settings](ribbon/smtp-settings.md)
- [Manage API Keys](ribbon/api-keys.md)

### Context Menu

- [Cell Right-Click Menu — Date Picker, Open URL](ribbon/context-menu.md)

