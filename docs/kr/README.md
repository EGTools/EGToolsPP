# EGTools++ 함수 매뉴얼 (한국어)

**EGTools++**는 구형 Excel(2010 이상)에서 Microsoft 365의 최신 함수를 쓰게 해 주는 단일 `.xll` 추가기능입니다(Windows 전용, 설치 프로그램·관리자 권한 불필요). XLOOKUP·FILTER·TEXTSPLIT·GROUPBY·정규식 함수 등을 어느 버전에서든 동일하게 제공하며, 호스트에 네이티브 함수가 있으면 `EG.` 접두사로, 없으면 원래 이름 그대로 등록되어 드롭인(drop-in)으로 동작합니다. 보이는 셀 복사, 메일머지, 달력 생성 같은 실무 유틸리티 리본도 포함합니다.

다운로드: [GitHub Releases](https://github.com/EGTools/EGToolsPP/releases) · 소스·문의: [GitHub 저장소](https://github.com/EGTools/EGToolsPP)

함수별 상세 설명입니다. IntelliSense 툴팁에서 함수명을 클릭하면 해당 페이지로 이동합니다.

## MS Excel 호환 함수 (69)

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
- [IMPORTCSV](functions/microsoft-365/IMPORTCSV.md) — Microsoft 365 (2026 신규)
- [IMPORTTEXT](functions/microsoft-365/IMPORTTEXT.md) — Microsoft 365 (2026 신규)
- [PERCENTOF](functions/microsoft-365/PERCENTOF.md)
- [PIVOTBY](functions/microsoft-365/PIVOTBY.md)
- [REGEXEXTRACT](functions/microsoft-365/REGEXEXTRACT.md)
- [REGEXREPLACE](functions/microsoft-365/REGEXREPLACE.md)
- [REGEXTEST](functions/microsoft-365/REGEXTEST.md)
- [TRIMRANGE](functions/microsoft-365/TRIMRANGE.md)

## Google Sheets 호환 함수 (7)

- [GOOGLETRANSLATE](functions/google/GOOGLETRANSLATE.md)
- [GTRS](functions/google/GTRS.md)
- [IMPORTDATA](functions/google/IMPORTDATA.md)
- [IMPORTFEED](functions/google/IMPORTFEED.md)
- [IMPORTHTML](functions/google/IMPORTHTML.md)
- [IMPORTRANGE](functions/google/IMPORTRANGE.md)
- [QUERY](functions/google/QUERY.md)

## EGTools 전용 함수 (52)

### 조회·참조 (3)

- [COMPARELIST](functions/eg-lookup/COMPARELIST.md)
- [ILOOKUP](functions/eg-lookup/ILOOKUP.md)
- [MVLOOKUP](functions/eg-lookup/MVLOOKUP.md)

### 데이터·배열 (4)

- [ADVANCEDFILTER](functions/eg-data/ADVANCEDFILTER.md)
- [EXPLODE](functions/eg-data/EXPLODE.md)
- [TEXTNUMSORT](functions/eg-data/TEXTNUMSORT.md)
- [UNPIVOT](functions/eg-data/UNPIVOT.md)

### 텍스트 (4)

- [STREXT](functions/eg-text/STREXT.md)
- [TEXTBETWEEN](functions/eg-text/TEXTBETWEEN.md)
- [TEXTREPLACE](functions/eg-text/TEXTREPLACE.md)
- [TRIMENDS](functions/eg-text/TRIMENDS.md)

### 날짜·시간 (4)

- [MONTHBYWEEK](functions/eg-datetime/MONTHBYWEEK.md)
- [NETWORKHOUR](functions/eg-datetime/NETWORKHOUR.md)
- [TODATETIME](functions/eg-datetime/TODATETIME.md)
- [WEEKNUMOFMONTH](functions/eg-datetime/WEEKNUMOFMONTH.md)

### 정보·수식 (4)

- [DIRFOLDER](functions/eg-info/DIRFOLDER.md)
- [EVAL](functions/eg-info/EVAL.md)
- [SHEETLIST](functions/eg-info/SHEETLIST.md)
- [TOTALPAGES](functions/eg-info/TOTALPAGES.md)

### 색상·보이는 셀 (5)

- [COUNTIFCOLOR](functions/eg-visible/COUNTIFCOLOR.md)
- [DISPLAYCOLOR](functions/eg-visible/DISPLAYCOLOR.md)
- [ISVISIBLE](functions/eg-visible/ISVISIBLE.md)
- [SUMIFCOLOR](functions/eg-visible/SUMIFCOLOR.md)
- [VISIBLEAGGR](functions/eg-visible/VISIBLEAGGR.md)

### 이미지·바코드 (6)

- [BARCODE](functions/eg-image/BARCODE.md)
- [CODE128](functions/eg-image/CODE128.md)
- [DATAMATRIX](functions/eg-image/DATAMATRIX.md)
- [QRCODE](functions/eg-image/QRCODE.md)
- [READBARCODE](functions/eg-image/READBARCODE.md)
- [READGS1](functions/eg-image/READGS1.md)

### 웹·API (3)

- [DECODEURL](functions/eg-web/DECODEURL.md)
- [JSONFILTER](functions/eg-web/JSONFILTER.md)
- [RESTAPI](functions/eg-web/RESTAPI.md)

### 한국 데이터 (11)

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

### 품질·생산 (8)

- [BOMTREE](functions/eg-quality/BOMTREE.md)
- [CP](functions/eg-quality/CP.md)
- [CPK](functions/eg-quality/CPK.md)
- [SAMPLINGAC](functions/eg-quality/SAMPLINGAC.md)
- [SAMPLINGLABEL](functions/eg-quality/SAMPLINGLABEL.md)
- [SAMPLINGRE](functions/eg-quality/SAMPLINGRE.md)
- [SAMPLINGSIZE](functions/eg-quality/SAMPLINGSIZE.md)
- [SUMBYBOM](functions/eg-quality/SUMBYBOM.md)

## 리본 메뉴

리본 `EGTools` 탭에서 제공되는 기능입니다. 전체 구성은 [리본 메뉴 개요](ribbon/README.md) 참고.

### 보이는 셀

- [보이는 셀 — 보이는 셀만 복사 · 전체 복사 · 값만/수식/모두 붙여넣기](ribbon/visible-cells.md)

### 셀 병합/나누기

- [셀 병합/나누기 — 연속값 병합 · 내용감춘 병합 · 내용병합 · 행/열끼리 병합 · 나누고 채우기 · 행/열 나누기](ribbon/merge-split.md)

### 그림/사진

- [그림/사진 — 이미지 삽입 · 이름에 삽입 · 양식에 삽입 · 선택 맞춤 · 모두 맞춤 · 추출 저장](ribbon/pictures.md)

### 달력/일정표

- [달력/일정표 — 월간 · 연간A/B · 주간 · 일일A/B](ribbon/calendar.md)

### 양식

- [양식 — 라벨지 만들기](ribbon/forms.md) · [메일머지](ribbon/mailmerge.md)

### 도구

- [도구 — 스타일 삭제 · 이름 삭제 · 오류 정리 · 빈셀 정리 · 메모 정돈 · 테두리 색 변경 · 모양 뽑기 · UDF 값 고정 · CheckBox 적용 · 전체 재계산](ribbon/utilities.md)

### EGTools

- [구버전 호환 적용](ribbon/apply-compat.md)
- [내장 함수로 복원](ribbon/restore-native.md)
- [SMTP 설정](ribbon/smtp-settings.md)
- [API 키 관리](ribbon/api-keys.md)

### 컨텍스트 메뉴

- [셀 오른쪽 클릭 메뉴 — 날짜입력 · 링크열기](ribbon/context-menu.md)

