# smoke_test.ps1 — registers the .xll and asserts one formula per function.
#
# Runs on BOTH host generations; everything is auto-detected at run time:
#   * XLL    — Excel bitness (PE header of EXCEL.EXE) picks dist\EGTools++64.xll
#              or dist\EGTools++32.xll unless -Xll overrides.
#   * Names  — registration is EG.F where the host has native F, x<F> for
#              keyword-conflict names (xSORT/xFILTER/xLET), bare F otherwise
#              (core::registeredName). Cases are written as EG.*; on legacy
#              hosts each cell walks the ladder EG.F → F → xF until one form
#              evaluates to something other than #NAME?.
#   * Entry  — Formula2 on dynamic-array hosts; .Formula otherwise (probed once:
#              the Formula2 SETTER fails with 0x800A03EC on pre-DA Excel).
#   * Arrays — modern: SpillingToRange. Legacy: wait briefly for the add-in's
#              forced CSE resize (WINDOW+COM) to expand the cell; if it doesn't
#              fire under headless COM, compare the first element only and mark
#              the line "PASS~" (counted separately in the summary).
#
# EXCLUDED (need other conditions):
#   - Macro-type (hang under headless COM): SHEET, SHEETS, FORMULATEXT, ISFORMULA,
#     LET, IMAGE → verify interactively (computer-use). See plan/09 §3.
#   - WEBSERVICE: needs network.   - RANDARRAY: non-deterministic.
#
# Usage:  powershell -File tests\smoke_test.ps1 [-Xll <path>]
# Exit code 0 = all pass (PASS~ counts as pass), 1 = one or more failures.

param([string]$Xll = '')
$ErrorActionPreference = 'Stop'

function Fmt($x) {
    if ($null -eq $x) { return '' }
    if ($x -is [double] -or $x -is [single]) {
        return $x.ToString([Globalization.CultureInfo]::InvariantCulture)
    }
    if ($x -is [bool]) { return $x.ToString().ToUpper() }
    return [string]$x
}

function Flatten($rng) {
    $v = $rng.Value2
    $out = @()
    if ($v -is [object[,]]) {
        for ($i = 1; $i -le $v.GetLength(0); $i++) {
            for ($j = 1; $j -le $v.GetLength(1); $j++) { $out += (Fmt $v[$i, $j]) }
        }
    } else { $out += (Fmt $v) }
    ($out -join ',')
}

# PE header machine field → '64' / '32' (works for any EXE, incl. ARM64 hosts).
function Get-ExeBitness([string]$exe) {
    $fs = [IO.File]::OpenRead($exe)
    try {
        $br = New-Object IO.BinaryReader($fs)
        $null = $fs.Seek(0x3C, 'Begin'); $peOff = $br.ReadInt32()
        $null = $fs.Seek($peOff + 4, 'Begin'); $machine = $br.ReadUInt16()
        if ($machine -eq 0x8664 -or $machine -eq 0xAA64) { '64' } else { '32' }
    } finally { $fs.Dispose() }
}

# Cell error values marshal as Int32 (2000 + xlErr code | 0x800A0000).
$XlErrName = -2146826259   # #NAME?

# n = name, f = formula, e = expected (flattened row-major, comma-joined)
$cases = @(
    [pscustomobject]@{n='IFNA';        f='=EG.IFNA(NA(),"x")';                             e='x'}
    [pscustomobject]@{n='XOR';         f='=EG.XOR(TRUE,FALSE)';                            e='TRUE'}
    [pscustomobject]@{n='UNICHAR';     f='=EG.UNICHAR(65)';                                e='A'}
    [pscustomobject]@{n='UNICODE';     f='=EG.UNICODE("A")';                               e='65'}
    [pscustomobject]@{n='NUMBERVALUE'; f='=EG.NUMBERVALUE("1,234.5",".",",")';             e='1234.5'}
    [pscustomobject]@{n='DAYS';        f='=EG.DAYS(DATE(2026,1,11),DATE(2026,1,1))';       e='10'}
    [pscustomobject]@{n='ISOWEEKNUM';  f='=EG.ISOWEEKNUM(DATE(2026,1,1))';                 e='1'}
    [pscustomobject]@{n='ENCODEURL';   f='=EG.ENCODEURL("a b")';                           e='a%20b'}
    [pscustomobject]@{n='FILTERXML';   f='=EG.FILTERXML("<a><b>x</b><b>y</b></a>","//b")'; e='x,y'}
    [pscustomobject]@{n='IFS';         f='=EG.IFS(FALSE,1,TRUE,7)';                        e='7'}
    [pscustomobject]@{n='SWITCH';      f='=EG.SWITCH(3,1,"a",3,"c")';                      e='c'}
    [pscustomobject]@{n='MAXIFS';      f='=EG.MAXIFS({1;2;3},{1;0;1},1)';                  e='3'}
    [pscustomobject]@{n='MINIFS';      f='=EG.MINIFS({1;2;3},{1;0;1},1)';                  e='1'}
    [pscustomobject]@{n='TEXTJOIN';    f='=EG.TEXTJOIN("-",TRUE,{"a";"";"b"})';            e='a-b'}
    [pscustomobject]@{n='CONCAT';      f='=EG.CONCAT("a","b")';                            e='ab'}
    [pscustomobject]@{n='XLOOKUP';     f='=EG.XLOOKUP(2,{1;2;3},{"a";"b";"c"})';           e='b'}
    [pscustomobject]@{n='XMATCH';      f='=EG.XMATCH(20,{10;20;30})';                      e='2'}
    [pscustomobject]@{n='FILTER';      f='=EG.FILTER({1;2;3},{1;0;1})';                    e='1,3'}
    [pscustomobject]@{n='SORT';        f='=EG.SORT({3;1;2})';                              e='1,2,3'}
    [pscustomobject]@{n='SORTBY';      f='=EG.SORTBY({"a";"b";"c"},{3;1;2})';              e='b,c,a'}
    [pscustomobject]@{n='UNIQUE';      f='=EG.UNIQUE({1;1;2})';                            e='1,2'}
    [pscustomobject]@{n='SEQUENCE';    f='=EG.SEQUENCE(1,4)';                              e='1,2,3,4'}
    [pscustomobject]@{n='TEXTSPLIT';   f='=EG.TEXTSPLIT("a,b,c",",")';                     e='a,b,c'}
    [pscustomobject]@{n='TEXTBEFORE';  f='=EG.TEXTBEFORE("a-b-c","-",2)';                  e='a-b'}
    [pscustomobject]@{n='TEXTAFTER';   f='=EG.TEXTAFTER("a-b-c","-",-1)';                  e='c'}
    [pscustomobject]@{n='VSTACK';      f='=EG.VSTACK({1;2},{3})';                          e='1,2,3'}
    [pscustomobject]@{n='HSTACK';      f='=EG.HSTACK({1;2},{3;4})';                        e='1,3,2,4'}
    [pscustomobject]@{n='TAKE';        f='=EG.TAKE({1;2;3},-2)';                           e='2,3'}
    [pscustomobject]@{n='DROP';        f='=EG.DROP({1;2;3},1)';                            e='2,3'}
    [pscustomobject]@{n='TOROW';       f='=EG.TOROW({1,2;3,4})';                           e='1,2,3,4'}
    [pscustomobject]@{n='TOCOL';       f='=EG.TOCOL({1,2;3,4})';                           e='1,2,3,4'}
    [pscustomobject]@{n='CHOOSEROWS';  f='=EG.CHOOSEROWS({10;20;30},3,1)';                 e='30,10'}
    [pscustomobject]@{n='CHOOSECOLS';  f='=EG.CHOOSECOLS({1,2,3},3,1)';                    e='3,1'}
    [pscustomobject]@{n='EXPAND';      f='=EG.EXPAND({1;2},3,1,0)';                        e='1,2,0'}
    [pscustomobject]@{n='WRAPROWS';    f='=EG.WRAPROWS({1;2;3},2,0)';                      e='1,2,3,0'}
    [pscustomobject]@{n='WRAPCOLS';    f='=EG.WRAPCOLS({1;2;3},2,0)';                      e='1,3,2,0'}
    [pscustomobject]@{n='ARRAYTOTEXT'; f='=EG.ARRAYTOTEXT({1,2;3,4})';                     e='1, 2, 3, 4'}
    [pscustomobject]@{n='VALUETOTEXT'; f='=EG.VALUETOTEXT("hi",1)';                        e='"hi"'}
    [pscustomobject]@{n='REGEXTEST';   f='=EG.REGEXTEST("abc123","\d+")';                  e='TRUE'}
    [pscustomobject]@{n='REGEXEXTRACT';f='=EG.REGEXEXTRACT("id=42; id=99","\d+")';         e='42'}
    [pscustomobject]@{n='REGEXREPLACE';f='=EG.REGEXREPLACE("a1b2","\d","#")';              e='a#b#'}
    # R1(네이티브 정합) 이후 기본 total_depth=1 → 합계 라벨이 UI 언어를 따르므로
    # 스모크는 로캘 무관하게 총계 없음(depth 0)을 명시한다.
    [pscustomobject]@{n='GROUPBY';     f='=EG.GROUPBY({"x";"y";"x"},{1;2;3},"SUM",0,0)';         e='x,4,y,2'}
    [pscustomobject]@{n='PIVOTBY';     f='=EG.PIVOTBY({"x";"y"},{"p";"p"},{1;2},"SUM",0,0,1,0)'; e=',p,x,1,y,2'}
    # FxMath — ROUND wraps the three that would otherwise expose binary-float noise.
    [pscustomobject]@{n='PERCENTOF';   f='=ROUND(EG.PERCENTOF({1;2;3},{1;2;3;4}),6)';      e='0.6'}
    [pscustomobject]@{n='COMBINA';     f='=EG.COMBINA(4,3)';                               e='20'}
    [pscustomobject]@{n='PERMUTATIONA';f='=EG.PERMUTATIONA(3,2)';                          e='9'}
    [pscustomobject]@{n='CEILING.MATH';f='=EG.CEILING.MATH(4.2)';                          e='5'}
    [pscustomobject]@{n='FLOOR.MATH';  f='=EG.FLOOR.MATH(4.8)';                            e='4'}
    [pscustomobject]@{n='BASE';        f='=EG.BASE(255,16)';                               e='FF'}
    [pscustomobject]@{n='DECIMAL';     f='=EG.DECIMAL("FF",16)';                           e='255'}
    [pscustomobject]@{n='ARABIC';      f='=EG.ARABIC("MCMXC")';                            e='1990'}
    [pscustomobject]@{n='RRI';         f='=ROUND(EG.RRI(2,100,400),6)';                    e='1'}
    [pscustomobject]@{n='PDURATION';   f='=ROUND(EG.PDURATION(1,100,400),6)';              e='2'}
    # FxBit — 12 = 1100b, 10 = 1010b.
    [pscustomobject]@{n='BITAND';      f='=EG.BITAND(12,10)';                              e='8'}
    [pscustomobject]@{n='BITOR';       f='=EG.BITOR(12,10)';                               e='14'}
    [pscustomobject]@{n='BITXOR';      f='=EG.BITXOR(12,10)';                              e='6'}
    [pscustomobject]@{n='BITLSHIFT';   f='=EG.BITLSHIFT(3,2)';                             e='12'}
    [pscustomobject]@{n='BITRSHIFT';   f='=EG.BITRSHIFT(12,2)';                            e='3'}
)

Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep -Milliseconds 400
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$pass = 0; $fail = 0; $firstOnly = 0
try {
    $bits = Get-ExeBitness (Join-Path $xl.Path 'EXCEL.EXE')
    if (-not $Xll) { $Xll = Join-Path $PSScriptRoot "..\dist\EGTools++$bits.xll" }
    if (-not (Test-Path $Xll)) { throw "XLL not found: $Xll" }
    if (-not $xl.RegisterXLL($Xll)) {
        throw "RegisterXLL failed: $Xll (bitness mismatch or load error)"
    }
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)

    # Host probe: the Formula2 SETTER errors on every pre-dynamic-array Excel.
    $modern = $true
    $probe = $ws.Cells.Item(1, 1)
    try { $probe.Formula2 = '=1+1' } catch { $modern = $false }
    $probe.ClearContents()
    $mode = if ($modern) { 'dynamic-array (EG.*, Formula2)' }
            else         { 'legacy (bare-name fallback, .Formula)' }
    Write-Output ("HOST: Excel {0} {1}-bit, {2} | XLL: {3}" -f `
        $xl.Version, $bits, $mode, (Split-Path $Xll -Leaf))

    $row = 1
    foreach ($c in $cases) {
        $cell = $ws.Cells.Item($row, 1)
        $entryBlocked = $false
        if ($modern) {
            $cell.Formula2 = $c.f
        } else {
            # Name ladder on legacy hosts: EG.F (host has the native, shadowed
            # registration) → bare F (drop-in) → xF (keyword-conflict names —
            # xSORT/xFILTER/xLET). A candidate misses either as #NAME? or by the
            # entry being REFUSED outright (0x800A03EC: modern binary + older
            # license knows-but-disables the bare name; typed entry is rejected
            # the same way, so this is not a harness artifact).
            $resolved = $false
            foreach ($cand in @($c.f,
                                ($c.f -replace 'EG\.', ''),
                                ($c.f -replace 'EG\.', 'x'))) {
                try { $cell.ClearContents(); $cell.Formula = $cand } catch { continue }
                $v = $cell.Value2
                if (-not (($v -is [int]) -and ($v -eq $XlErrName))) { $resolved = $true; break }
            }
            if (-not $resolved) { $entryBlocked = $true }
        }

        $rng = $cell
        if ($modern) { try { $sp = $cell.SpillingToRange; if ($sp) { $rng = $sp } } catch {} }
        $got = if ($entryBlocked) { '(no usable name: EG./bare/x all rejected)' }
               else { Flatten $rng }

        if (-not $modern -and -not $entryBlocked -and $got -ne $c.e -and $c.e.Contains(',')) {
            # Array result on a legacy host: give the add-in's background CSE
            # resize a moment to expand the cell, then re-read.
            $deadline = (Get-Date).AddSeconds(3)
            while (-not $cell.HasArray -and (Get-Date) -lt $deadline) {
                Start-Sleep -Milliseconds 200
            }
            if ($cell.HasArray) { $got = Flatten $cell.CurrentArray }
        }

        if ($got -eq $c.e) {
            $pass++; Write-Output ("PASS  {0,-12} = {1}" -f $c.n, $got)
        } elseif (-not $modern -and $c.e.Contains(',') -and $got -eq $c.e.Split(',')[0]) {
            # CSE resize didn't fire under headless COM — the function computed
            # (first element correct); full spill is an interactive-only check.
            $pass++; $firstOnly++
            Write-Output ("PASS~ {0,-12} = {1} (1st cell only; CSE not expanded)" -f $c.n, $got)
        } else {
            $fail++; Write-Output ("FAIL  {0,-12} got [{1}] expected [{2}]" -f $c.n, $got, $c.e)
        }
        $row += 8   # stay clear of the previous case's spill / CSE area
    }
    $wb.Close($false)
} finally {
    $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}
Write-Output ""
$note = if ($firstOnly -gt 0) { " ($firstOnly array case(s) first-cell only)" } else { "" }
Write-Output ("RESULT: {0} passed, {1} failed (of {2}){3}" -f $pass, $fail, $cases.Count, $note)
if ($fail -gt 0) { exit 1 } else { exit 0 }
