# make_sample_newfuncs.ps1 — 호환 변환 리본 테스트용 샘플 생성기.
#
# Excel 2021+ 신함수(= core::funcTable의 2021/2024/365 항목 전부)를 한 시트에
# 담은 sample_newfuncs.xlsx 를 "실제 설치된 모던 Excel"(COM 자동화)로 작성한다.
# 이렇게 만들면 저장 토큰(_xlfn.* / _xlfn._xlws.* / _xleta.SUM)이 실물과 동일해
#   * 구버전(2016 x86) 에서: 열면 #NAME? → [구버전 호환 적용] 전 함수 변환 실측
#   * 모던(365)      에서: [구버전 호환 적용] → [내장 함수로 복원] 왕복 실측
# 두 방향을 한 파일로 시험할 수 있다. IMPORTTEXT/IMPORTCSV 용 sample_import.csv
# 도 같은 폴더에 만든다(수식은 CELL("filename") 기반이라 드라이브 문자 무관).
#
# 실행: powershell -ExecutionPolicy Bypass -File tests\make_sample_newfuncs.ps1
# 요구: 동적배열 지원 Excel(2021/365)이 설치된 부팅 환경에서 실행할 것.

$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsxPath = Join-Path $here 'sample_newfuncs.xlsx'
$csvPath  = Join-Path $here 'sample_import.csv'

# ── IMPORT 함수용 CSV (UTF-8 BOM — IMPORTCSV는 콤마+UTF-8 고정) ──────────────
@'
품목,수량
연필,10
지우개,5
공책,3
'@ | Out-File -FilePath $csvPath -Encoding utf8

if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }

# ── 함수 목록: 이름, 구분, 수식(en-US 구문), 세로 점유 행수, 비고 ────────────
# Formula2 로 넣으므로 모던 Excel이 저장 시 스스로 _xlfn./_xleta. 토큰을 기록.
$entries = @(
    @{ N='XLOOKUP';     V='2021';     R=1;  F='=XLOOKUP(105,Data!A2:A11,Data!B2:B11)';            T='사번 105 → 이름' }
    @{ N='XMATCH';      V='2021';     R=1;  F='=XMATCH(106,Data!A2:A11)';                          T='사번 106의 위치(6)' }
    @{ N='FILTER';      V='2021';     R=4;  F='=FILTER(Data!B2:B11,Data!C2:C11="영업")';           T='호환 적용 시 xFILTER로 변환' }
    @{ N='SORT';        V='2021';     R=10; F='=SORT(Data!D2:D11,1,-1)';                           T='호환 적용 시 xSORT로 변환' }
    @{ N='SORTBY';      V='2021';     R=10; F='=SORTBY(Data!B2:B11,Data!D2:D11,-1)';               T='매출 내림차순 이름' }
    @{ N='UNIQUE';      V='2021';     R=3;  F='=UNIQUE(Data!C2:C11)';                              T='부서 고유값 3종' }
    @{ N='SEQUENCE';    V='2021';     R=5;  F='=SEQUENCE(5)';                                      T='1..5' }
    @{ N='RANDARRAY';   V='2021';     R=3;  F='=RANDARRAY(3,1,1,100,TRUE)';                        T='휘발성 — 열 때마다 값 변경' }
    @{ N='LET';         V='2021';     R=1;  F='=LET(x,Data!D2:D11,SUM(x)/COUNT(x))';               T='호환 적용 시 xLET로 변환(평균 매출)' }
    @{ N='TEXTSPLIT';   V='2024/365'; R=1;  F='=TEXTSPLIT("사과,배,포도",",")';                    T='가로 3칸 분할' }
    @{ N='TEXTBEFORE';  V='2024/365'; R=1;  F='=TEXTBEFORE("2026-08-03","-")';                     T='' }
    @{ N='TEXTAFTER';   V='2024/365'; R=1;  F='=TEXTAFTER("2026-08-03","-",2)';                    T='' }
    @{ N='VSTACK';      V='2024/365'; R=4;  F='=VSTACK(Data!B2:B3,Data!B9:B10)';                   T='' }
    @{ N='HSTACK';      V='2024/365'; R=2;  F='=HSTACK(Data!B2:B3,Data!C2:C3)';                    T='' }
    @{ N='TAKE';        V='2024/365'; R=3;  F='=TAKE(Data!B2:B11,3)';                              T='' }
    @{ N='DROP';        V='2024/365'; R=3;  F='=DROP(Data!B2:B11,7)';                              T='' }
    @{ N='TOROW';       V='2024/365'; R=1;  F='=TOROW(Data!C2:C4)';                                T='' }
    @{ N='TOCOL';       V='2024/365'; R=4;  F='=TOCOL(Data!B2:C3)';                                T='' }
    @{ N='CHOOSEROWS';  V='2024/365'; R=3;  F='=CHOOSEROWS(Data!B2:B11,1,3,5)';                    T='' }
    @{ N='CHOOSECOLS';  V='2024/365'; R=2;  F='=CHOOSECOLS(Data!A2:D3,2,4)';                       T='' }
    @{ N='EXPAND';      V='2024/365'; R=4;  F='=EXPAND(Data!B2:B3,4,2,"-")';                       T='부족분 "-" 채움' }
    @{ N='WRAPROWS';    V='2024/365'; R=2;  F='=WRAPROWS(Data!B2:B7,3)';                           T='' }
    @{ N='WRAPCOLS';    V='2024/365'; R=3;  F='=WRAPCOLS(Data!B2:B7,3)';                           T='' }
    @{ N='ARRAYTOTEXT'; V='2024/365'; R=1;  F='=ARRAYTOTEXT(Data!B2:B4)';                          T='' }
    @{ N='VALUETOTEXT'; V='2024/365'; R=1;  F='=VALUETOTEXT(Data!B2)';                             T='' }
    @{ N='REGEXTEST';   V='2024/365'; R=1;  F='=REGEXTEST("모델명 EG-2026","\d+")';                T='' }
    @{ N='REGEXREPLACE';V='2024/365'; R=1;  F='=REGEXREPLACE("010-1234-5678","\d{4}$","****")';    T='' }
    @{ N='REGEXEXTRACT';V='2024/365'; R=1;  F='=REGEXEXTRACT("주문번호 A-1234","\d+")';            T='' }
    @{ N='GROUPBY';     V='2024/365'; R=5;  F='=GROUPBY(Data!C2:C11,Data!D2:D11,SUM)';             T='호환 적용 시 집계자 SUM↔"SUM" 토글' }
    @{ N='PIVOTBY';     V='2024/365'; R=6;  F='=PIVOTBY(Data!C2:C11,Data!E2:E11,Data!D2:D11,SUM)'; T='호환 적용 시 집계자 SUM↔"SUM" 토글' }
    @{ N='PERCENTOF';   V='2024/365'; R=1;  F='=PERCENTOF(Data!D2:D5,Data!D2:D11)';                T='상위 4명 매출 비중' }
    @{ N='TRIMRANGE';   V='2024/365'; R=3;  F='=TRIMRANGE(Data!G1:G6)';                            T='끝쪽 빈 행 제거(G4:G6)' }
    @{ N='IMAGE';       V='2024/365'; R=1;  F='=IMAGE("https://www.google.com/favicon.ico")';      T='인터넷 필요 · 복원 시 EG.IMAGE 유지' }
    @{ N='IMPORTTEXT';  V='365(2026)';R=4;  F='=IMPORTTEXT(LEFT(CELL("filename",$A$1),FIND("[",CELL("filename",$A$1))-1)&"sample_import.csv",",")'; T='같은 폴더 sample_import.csv 필요' }
    @{ N='IMPORTCSV';   V='365(2026)';R=4;  F='=IMPORTCSV(LEFT(CELL("filename",$A$1),FIND("[",CELL("filename",$A$1))-1)&"sample_import.csv")';      T='같은 폴더 sample_import.csv 필요' }
)

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt 3) { [void]$wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    $shInfo = $wb.Worksheets.Item(1); $shInfo.Name = '안내'
    $shFn   = $wb.Worksheets.Item(2); $shFn.Name   = '새함수'
    $shData = $wb.Worksheets.Item(3); $shData.Name = 'Data'

    # ── Data 시트 ────────────────────────────────────────────────────────────
    $data = @(
        @('사번','이름','부서','매출','지역'),
        @(101,'김민준','영업',1200,'서울'), @(102,'이서연','개발', 950,'부산'),
        @(103,'박지훈','영업', 800,'서울'), @(104,'최수아','기획',1100,'부산'),
        @(105,'정우진','개발', 700,'서울'), @(106,'강하은','영업',1500,'부산'),
        @(107,'조현우','기획', 600,'서울'), @(108,'윤지민','개발',1300,'부산'),
        @(109,'임서준','영업', 900,'서울'), @(110,'한예은','기획',1000,'부산'))
    $arr = New-Object 'object[,]' $data.Count, 5
    for ($i = 0; $i -lt $data.Count; $i++) {
        for ($j = 0; $j -lt 5; $j++) { $arr[$i, $j] = $data[$i][$j] }
    }
    $shData.Range('A1:E' + $data.Count).Value2 = $arr
    # TRIMRANGE 시험 데이터: G1:G3 값, G4:G6 빈칸
    $g = New-Object 'object[,]' 3, 1
    $g[0,0] = 10; $g[1,0] = 20; $g[2,0] = 30
    $shData.Range('G1:G3').Value2 = $g

    # ── 새함수 시트 ──────────────────────────────────────────────────────────
    $shFn.Columns.Item(3).NumberFormat = '@'   # C열(수식 원문)은 텍스트로
    $hdr = @('함수','구분','원본 수식(참고)','비고','', '결과 →')
    for ($j = 0; $j -lt $hdr.Count; $j++) { $shFn.Cells.Item(1, $j + 1).Value2 = $hdr[$j] }
    $shFn.Rows.Item(1).Font.Bold = $true

    $r = 3
    foreach ($e in $entries) {
        $shFn.Cells.Item($r,1).Value2 = $e.N
        $shFn.Cells.Item($r,2).Value2 = $e.V
        $shFn.Cells.Item($r,3).Value2 = $e.F
        $shFn.Cells.Item($r,4).Value2 = $e.T
        $shFn.Cells.Item($r,6).Formula2 = $e.F
        $e.Row = $r
        $r += [int]$e.R + 1
    }

    # 정의된 이름(Names.RefersTo 변환 경로 시험) + 이를 쓰는 셀
    [void]$wb.Names.Add('매출상위3', '=TAKE(SORT(Data!$D$2:$D$11,1,-1),3)')
    $shFn.Cells.Item($r,1).Value2 = '(정의된 이름)'
    $shFn.Cells.Item($r,3).Value2 = '매출상위3 = TAKE(SORT(...)) / 셀 =SUM(매출상위3)'
    $shFn.Cells.Item($r,4).Value2 = '정의된 이름 RefersTo 변환 확인용'
    $shFn.Cells.Item($r,6).Formula2 = '=SUM(매출상위3)'
    $nameRow = $r

    $shFn.Columns.Item(1).ColumnWidth = 14
    $shFn.Columns.Item(2).ColumnWidth = 10
    $shFn.Columns.Item(3).ColumnWidth = 58
    $shFn.Columns.Item(4).ColumnWidth = 38

    # ── 안내 시트 ────────────────────────────────────────────────────────────
    $info = @(
        'EGTools++ 호환 변환 테스트 샘플 — Excel 2021+ 신함수 전체',
        '',
        '이 파일은 모던 Excel(동적배열)로 작성되어 신함수가 _xlfn./_xleta. 토큰으로 저장되어 있습니다.',
        '[새함수] 시트에 함수별 실제 수식(F열)과 원본 수식 텍스트(C열)가 있습니다.',
        '',
        '■ 구버전 Excel(예: 2016 x86 + EGTools++)에서:',
        '  1) 열면 F열이 _xlfn.* 표시와 함께 #NAME? 오류로 보입니다.',
        '  2) 리본 [EGTools++] > [구버전 호환 적용] 실행 → 전 함수가 EGTools 함수명(bare/x접두)으로 바뀌고 계산됩니다.',
        '  3) [내장 함수로 복원]은 이 버전엔 네이티브가 없으므로 "변환할 대상이 없습니다"가 정상입니다.',
        '',
        '■ 모던 Excel(2021/365 + EGTools++)에서:',
        '  1) 열면 전부 네이티브로 정상 계산됩니다(IMPORT 2종 제외 가능).',
        '  2) [구버전 호환 적용] → SORT/FILTER/LET은 xSORT/xFILTER/xLET로, GROUPBY/PIVOTBY 집계자는 "SUM"으로 변경.',
        '     (문서가 구버전용으로 바뀌므로 이 상태에선 일부 #NAME? 표시가 정상입니다)',
        '  3) [내장 함수로 복원] → 전 함수가 네이티브로 재토큰화되어 다시 계산됩니다(IMAGE는 EG.IMAGE 유지).',
        '',
        '※ RANDARRAY는 휘발성이라 열 때마다 값이 바뀝니다.',
        '※ IMAGE는 인터넷 연결이 필요합니다.',
        '※ IMPORTTEXT/IMPORTCSV는 같은 폴더의 sample_import.csv를 읽습니다(365 네이티브는 2026 롤아웃).')
    for ($i = 0; $i -lt $info.Count; $i++) { $shInfo.Cells.Item($i + 1, 1).Value2 = $info[$i] }
    $shInfo.Cells.Item(1,1).Font.Bold = $true

    # ── 저장(경로 확정) → CELL("filename") 재계산 → 재저장 ──────────────────
    $wb.SaveAs($xlsxPath, 51)   # xlOpenXMLWorkbook
    $xl.CalculateFullRebuild()
    $wb.Save()

    # ── 결과 검증 출력 ───────────────────────────────────────────────────────
    Write-Output "=== 결과 (앵커 셀 표시값) ==="
    foreach ($e in $entries) {
        $t = $shFn.Cells.Item($e.Row, 6).Text
        Write-Output ("{0,-14} {1}" -f $e.N, $t)
    }
    Write-Output ("{0,-14} {1}" -f '(이름)SUM', $shFn.Cells.Item($nameRow, 6).Text)
    Write-Output "저장: $xlsxPath"

    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
