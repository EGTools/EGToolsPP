# Regenerate manual/kr + manual/en with function category and native-intro version.
# Classification single source: src/core/Version.cpp funcTable (Excel-compat + intro year).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$utf8 = New-Object System.Text.UTF8Encoding $false

# 1) Excel-compat map: name -> intro year (parse funcTable)
$ver = [IO.File]::ReadAllText((Join-Path $root 'src\core\Version.cpp'))
$excelYear = @{}
foreach ($m in [regex]::Matches($ver, '\{\s*L"([A-Z0-9.]+)",\s*\{\s*(\d{4}),')) {
    $excelYear[$m.Groups[1].Value] = [int]$m.Groups[2].Value
}
Write-Output ("excel-compat entries: " + $excelYear.Count)

$googleFns = @('GOOGLETRANSLATE','GTRS','IMPORTDATA','IMPORTFEED','IMPORTRANGE','IMPORTHTML','QUERY')
$skip = @('HELLO','SEQ2D')

function YearLabel([string]$fn, [int]$y, [string]$lang) {
    if ($fn -eq 'IMPORTTEXT' -or $fn -eq 'IMPORTCSV') {
        if ($lang -eq 'kr') { return 'Excel 365 (2026 신규)' } else { return 'Excel 365 (new in 2026)' }
    }
    switch ($y) {
        2013 { return 'Excel 2013' }
        2019 { return 'Excel 2016/2019' }
        2021 { return 'Excel 2021' }
        2024 { return 'Excel 2024 / Microsoft 365' }
        default { return "Excel $y" }
    }
}

function CategoryBlock([string]$fn, [string]$lang) {
    if ($excelYear.ContainsKey($fn)) {
        $yl = YearLabel $fn $excelYear[$fn] $lang
        $xNote = ''
        if (@('SORT','FILTER','LET') -contains $fn) {
            $xNote = if ($lang -eq 'kr') { " 일부 구버전에서는 예약어 충돌로 ``x$fn`` 이름으로 등록됩니다." }
                     else { " On some legacy hosts it registers as ``x$fn`` due to a keyword conflict." }
        }
        if ($lang -eq 'kr') {
            return @("**종류**: MS Excel 호환 함수 · **네이티브 도입**: $yl",
                     ("- 지원: Excel 2010+. 네이티브가 없는 구버전에서는 ``$fn`` 그대로(드롭인), " +
                      "네이티브가 있는 최신 Excel에서는 ``EG.$fn``으로 등록됩니다." + $xNote))
        } else {
            return @("**Category**: MS Excel compatibility function · **Native since**: $yl",
                     ("- Supported: Excel 2010+. Registered as ``$fn`` (drop-in) on hosts without the native function, " +
                      "and as ``EG.$fn`` on modern Excel that has it." + $xNote))
        }
    }
    if ($googleFns -contains $fn) {
        if ($lang -eq 'kr') {
            return @('**종류**: Google Sheets 호환 함수',
                     "- 지원: Excel 2010+. 모든 Excel 버전에서 ``$fn`` 이름 그대로 등록됩니다.")
        } else {
            return @('**Category**: Google Sheets compatibility function',
                     "- Supported: Excel 2010+. Always registered as ``$fn`` on every Excel version.")
        }
    }
    if ($lang -eq 'kr') {
        return @('**종류**: EGTools 전용 함수',
                 "- 지원: Excel 2010+. 모든 Excel 버전에서 ``$fn`` 이름 그대로 등록됩니다.")
    } else {
        return @('**Category**: EGTools-only function',
                 "- Supported: Excel 2010+. Always registered as ``$fn`` on every Excel version.")
    }
}

function EscCell([string]$s) { if ($null -eq $s) { return '' } return $s.Replace('|', '\|').Replace("`r", '').Replace("`n", ' ') }

function Generate([string]$lang) {
    $isKr = $lang -eq 'kr'
    $catalog = Get-Content (Join-Path $root ("resources\i18n\" + $(if ($isKr) { 'ko' } else { 'en' }) + ".json")) -Raw -Encoding UTF8 | ConvertFrom-Json
    $en = Get-Content (Join-Path $root 'resources\i18n\en.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $contentFile = Join-Path $root "manual\_content\$lang.json"
    $content = $null
    if (Test-Path $contentFile) { $content = Get-Content $contentFile -Raw -Encoding UTF8 | ConvertFrom-Json }
    $outDir = Join-Path $root "manual\$lang"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory $outDir | Out-Null }

    $names = @()
    foreach ($prop in $catalog.func.PSObject.Properties) {
        $fn = $prop.Name
        if ($skip -contains $fn) { continue }
        $names += $fn
        $f = $prop.Value
        $ef = $en.func.$fn

        $sigParts = @()
        for ($i = 0; $i -lt $f.args.Count; $i++) {
            $opt = ($ef -and $i -lt $ef.args.Count -and $ef.args[$i].opt)
            $sigParts += $(if ($opt) { "[" + $f.args[$i].name + "]" } else { $f.args[$i].name })
        }
        $sig = "$fn(" + ($sigParts -join ', ') + $(if ($ef -and $ef.repeat) { ', ...' } else { '' }) + ')'

        $rows = @()
        for ($i = 0; $i -lt $f.args.Count; $i++) {
            $opt = ($ef -and $i -lt $ef.args.Count -and $ef.args[$i].opt)
            $req = if ($isKr) { $(if ($opt) { '선택' } else { '필수' }) } else { $(if ($opt) { 'Optional' } else { 'Required' }) }
            $rows += "| $($f.args[$i].name) | $req | $($f.args[$i].help) |"
        }
        if ($ef -and $ef.repeat) {
            $rows += $(if ($isKr) { '| … | 선택 | 위 인수(들)를 반복 지정할 수 있습니다 |' }
                       else { '| ... | Optional | the argument(s) above may be repeated |' })
        }
        $hdr = if ($isKr) { "| 인수 | 필수 | 설명 |" } else { "| Argument | Required | Description |" }
        $argTable = if ($rows.Count) { "$hdr`n|---|---|---|`n" + ($rows -join "`n") } else { $(if ($isKr) { '(인수 없음)' } else { '(no arguments)' }) }

        $cat = CategoryBlock $fn $lang
        $sSyntax  = if ($isKr) { '구문' } else { 'Syntax' }
        $sArgs    = if ($isKr) { '인수' } else { 'Arguments' }
        $sReturns = if ($isKr) { '반환' } else { 'Returns' }
        $sEx      = if ($isKr) { '예제' } else { 'Examples' }
        $sNotes   = if ($isKr) { '참고' } else { 'Notes' }
        $exHdr    = if ($isKr) { '| 수식 | 결과 | 설명 |' } else { '| Formula | Result | Description |' }
        $todoRet  = if ($isKr) { '<!-- TODO: 반환 형태(스칼라/배열·스필)와 의미, 오류 조건(#VALUE!, #N/A 등) -->' }
                    else { '<!-- TODO: return shape (scalar / spilled array), meaning, error conditions (#VALUE!, #N/A ...) -->' }
        $todoNote = if ($isKr) { '<!-- TODO: 원본(Excel/Google)과의 차이, 관련 함수 링크 -->' }
                    else { '<!-- TODO: differences from the Excel/Google original, related functions -->' }

        $c = $null
        if ($content) { $c = $content.$fn }

        $retBlock = if ($c -and $c.ret) { $c.ret } else { $todoRet }

        $exRows = @()
        if ($c -and $c.examples -and @($c.examples).Count) {
            foreach ($e in $c.examples) {
                $exRows += "| ``$($e[0])`` | " + (EscCell $e[1]) + " | " + (EscCell $e[2]) + " |"
            }
        } else {
            $exRows += "| ``=$fn(...)`` | | <!-- TODO --> |"
        }
        $exBlock = "$exHdr`n|---|---|---|`n" + ($exRows -join "`n")

        $noteLines = @()
        if ($c -and $c.notes -and @($c.notes).Count) {
            foreach ($n in $c.notes) { $noteLines += "- $n" }
        } elseif (-not $c) {
            $noteLines += $todoNote
        }
        $noteLines += $cat[1]
        $noteBlock = $noteLines -join "`n"

        $md = "# $fn`n`n" + $cat[0] + "`n`n" + $f.desc + "`n`n" +
              "## $sSyntax`n`n``````" + "`n=$sig`n" + "```````n`n" +
              "## $sArgs`n`n$argTable`n`n" +
              "## $sReturns`n`n$retBlock`n`n" +
              "## $sEx`n`n$exBlock`n`n" +
              "## $sNotes`n`n$noteBlock`n"
        [IO.File]::WriteAllText((Join-Path $outDir "$fn.md"), $md, $utf8)
    }

    # index grouped by category
    $names = $names | Sort-Object
    $excelList  = $names | Where-Object { $excelYear.ContainsKey($_) }
    $googleList = $names | Where-Object { $googleFns -contains $_ }
    $egList     = $names | Where-Object { -not $excelYear.ContainsKey($_) -and -not ($googleFns -contains $_) }
    if ($isKr) {
        $idx = "# EGTools++ 함수 매뉴얼 (한국어)`n`n함수별 상세 설명입니다. IntelliSense 툴팁에서 함수명을 클릭하면 해당 페이지로 이동합니다.`n`n"
        $sections = @(@('MS Excel 호환 함수', $excelList), @('Google Sheets 호환 함수', $googleList), @('EGTools 전용 함수', $egList))
    } else {
        $idx = "# EGTools++ Function Manual (English)`n`nPer-function reference. Clicking a function name in the IntelliSense tooltip opens the matching page.`n`n"
        $sections = @(@('MS Excel compatibility functions', $excelList), @('Google Sheets compatibility functions', $googleList), @('EGTools-only functions', $egList))
    }
    foreach ($sec in $sections) {
        $idx += "## " + $sec[0] + " (" + $sec[1].Count + ")`n`n"
        foreach ($n in $sec[1]) {
            if ($excelYear.ContainsKey($n)) { $idx += "- [$n]($n.md) — " + (YearLabel $n $excelYear[$n] $lang) + "`n" }
            else { $idx += "- [$n]($n.md)`n" }
        }
        $idx += "`n"
    }
    # ribbon section (hand-written pages under ribbon/, not regenerated here)
    if ($isKr) {
        $idx += "## 리본 메뉴`n`n리본 ``EGTools++`` 탭에서 제공되는 기능입니다. 전체 구성은 [리본 메뉴 개요](ribbon/README.md) 참고.`n`n" +
                "### EGTools (수식 변환)`n`n" +
                "- [구버전 호환 적용](ribbon/apply-compat.md)`n" +
                "- [내장 함수로 복원](ribbon/restore-native.md)`n`n"
    } else {
        $idx += "## Ribbon Menu`n`nFeatures on the ``EGTools++`` ribbon tab. See the [ribbon overview](ribbon/README.md) for the full layout.`n`n" +
                "### EGTools (formula conversion)`n`n" +
                "- [Apply Legacy Compatibility](ribbon/apply-compat.md)`n" +
                "- [Restore Native Functions](ribbon/restore-native.md)`n`n"
    }
    [IO.File]::WriteAllText((Join-Path $outDir 'README.md'), $idx, $utf8)
    Write-Output ("$lang -> " + $names.Count + " pages (excel=" + $excelList.Count + " google=" + $googleList.Count + " egtools=" + $egList.Count + ")")
}

Generate 'kr'
Generate 'en'
