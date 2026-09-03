# EGTools++ 関数マニュアル (日本語)

**EGTools++** は、旧バージョンの Excel(2010 以降)で Microsoft 365 の最新関数を使えるようにする単一ファイルの `.xll` アドインです(Windows 専用、インストーラー・管理者権限不要)。XLOOKUP・FILTER・TEXTSPLIT・GROUPBY・正規表現関数などをどのバージョンでも同じように提供し、ホストにネイティブ関数がなければ元の名前のまま(ドロップイン)、あれば `EG.` 接頭辞付きで登録されます。可視セルのコピー、差し込みメール、カレンダー作成などの実務ユーティリティ リボンも含まれます。

ダウンロード: [GitHub Releases](https://github.com/EGTools/EGToolsPP/releases) · ソース・お問い合わせ: [GitHub リポジトリ](https://github.com/EGTools/EGToolsPP)

関数ごとの詳細説明です。IntelliSense ツールチップで関数名をクリックすると該当ページに移動します。

## MS Excel 互換関数 (69)

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
- [IMPORTCSV](functions/microsoft-365/IMPORTCSV.md) — Microsoft 365 (2026 新規)
- [IMPORTTEXT](functions/microsoft-365/IMPORTTEXT.md) — Microsoft 365 (2026 新規)
- [PERCENTOF](functions/microsoft-365/PERCENTOF.md)
- [PIVOTBY](functions/microsoft-365/PIVOTBY.md)
- [REGEXEXTRACT](functions/microsoft-365/REGEXEXTRACT.md)
- [REGEXREPLACE](functions/microsoft-365/REGEXREPLACE.md)
- [REGEXTEST](functions/microsoft-365/REGEXTEST.md)
- [TRIMRANGE](functions/microsoft-365/TRIMRANGE.md)

## Google スプレッドシート互換関数 (7)

- [GOOGLETRANSLATE](functions/google/GOOGLETRANSLATE.md)
- [GTRS](functions/google/GTRS.md)
- [IMPORTDATA](functions/google/IMPORTDATA.md)
- [IMPORTFEED](functions/google/IMPORTFEED.md)
- [IMPORTHTML](functions/google/IMPORTHTML.md)
- [IMPORTRANGE](functions/google/IMPORTRANGE.md)
- [QUERY](functions/google/QUERY.md)

## EGTools 専用関数 (52)

### 検索・参照 (3)

- [COMPARELIST](functions/eg-lookup/COMPARELIST.md)
- [ILOOKUP](functions/eg-lookup/ILOOKUP.md)
- [MVLOOKUP](functions/eg-lookup/MVLOOKUP.md)

### データ・配列 (4)

- [ADVANCEDFILTER](functions/eg-data/ADVANCEDFILTER.md)
- [EXPLODE](functions/eg-data/EXPLODE.md)
- [TEXTNUMSORT](functions/eg-data/TEXTNUMSORT.md)
- [UNPIVOT](functions/eg-data/UNPIVOT.md)

### テキスト (4)

- [STREXT](functions/eg-text/STREXT.md)
- [TEXTBETWEEN](functions/eg-text/TEXTBETWEEN.md)
- [TEXTREPLACE](functions/eg-text/TEXTREPLACE.md)
- [TRIMENDS](functions/eg-text/TRIMENDS.md)

### 日付・時刻 (4)

- [MONTHBYWEEK](functions/eg-datetime/MONTHBYWEEK.md)
- [NETWORKHOUR](functions/eg-datetime/NETWORKHOUR.md)
- [TODATETIME](functions/eg-datetime/TODATETIME.md)
- [WEEKNUMOFMONTH](functions/eg-datetime/WEEKNUMOFMONTH.md)

### 情報・数式 (4)

- [DIRFOLDER](functions/eg-info/DIRFOLDER.md)
- [EVAL](functions/eg-info/EVAL.md)
- [SHEETLIST](functions/eg-info/SHEETLIST.md)
- [TOTALPAGES](functions/eg-info/TOTALPAGES.md)

### 色・可視セル (5)

- [COUNTIFCOLOR](functions/eg-visible/COUNTIFCOLOR.md)
- [DISPLAYCOLOR](functions/eg-visible/DISPLAYCOLOR.md)
- [ISVISIBLE](functions/eg-visible/ISVISIBLE.md)
- [SUMIFCOLOR](functions/eg-visible/SUMIFCOLOR.md)
- [VISIBLEAGGR](functions/eg-visible/VISIBLEAGGR.md)

### 画像・バーコード (6)

- [BARCODE](functions/eg-image/BARCODE.md)
- [CODE128](functions/eg-image/CODE128.md)
- [DATAMATRIX](functions/eg-image/DATAMATRIX.md)
- [QRCODE](functions/eg-image/QRCODE.md)
- [READBARCODE](functions/eg-image/READBARCODE.md)
- [READGS1](functions/eg-image/READGS1.md)

### Web・API (3)

- [DECODEURL](functions/eg-web/DECODEURL.md)
- [JSONFILTER](functions/eg-web/JSONFILTER.md)
- [RESTAPI](functions/eg-web/RESTAPI.md)

### 韓国データ (11)

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

### 品質・生産管理 (8)

- [BOMTREE](functions/eg-quality/BOMTREE.md)
- [CP](functions/eg-quality/CP.md)
- [CPK](functions/eg-quality/CPK.md)
- [SAMPLINGAC](functions/eg-quality/SAMPLINGAC.md)
- [SAMPLINGLABEL](functions/eg-quality/SAMPLINGLABEL.md)
- [SAMPLINGRE](functions/eg-quality/SAMPLINGRE.md)
- [SAMPLINGSIZE](functions/eg-quality/SAMPLINGSIZE.md)
- [SUMBYBOM](functions/eg-quality/SUMBYBOM.md)

## リボンメニュー

リボンの `EGTools` タブで提供される機能です。全体構成は[リボンメニュー概要](ribbon/README.md)を参照してください。

### 表示セル

- [表示セル — 表示セルのみコピー · 全体コピー · 値のみ/数式/すべて貼り付け](ribbon/visible-cells.md)

### セル結合/分割

- [セル結合/分割 — 連続値結合 · 内容を残す結合 · 内容結合 · 行結合/列結合 · 分割して埋める · 行分割/列分割](ribbon/merge-split.md)

### 画像/写真

- [画像/写真 — 画像の挿入 · 名前で挿入 · フォームへ挿入 · 選択を調整 · すべて調整 · 抽出保存](ribbon/pictures.md)

### カレンダー/予定表

- [カレンダー/予定表 — 月間 · 年間A/B · 週間 · 日次A/B](ribbon/calendar.md)

### フォーム

- [フォーム — ラベル作成](ribbon/forms.md) · [差し込み](ribbon/mailmerge.md)

### ツール

- [ツール — スタイル削除 · 名前の削除 · エラー整理 · 空セル整理 · メモの整頓 · 罫線の色を変更 · 形状トレース · UDF 値固定 · CheckBox 適用 · 全体を再計算](ribbon/utilities.md)

### EGTools

- [旧バージョン互換に変換](ribbon/apply-compat.md)
- [組み込み関数に復元](ribbon/restore-native.md)
- [SMTP設定](ribbon/smtp-settings.md)
- [API キー管理](ribbon/api-keys.md)

### コンテキストメニュー

- [セル右クリックメニュー — 日付入力 · リンクを開く](ribbon/context-menu.md)

