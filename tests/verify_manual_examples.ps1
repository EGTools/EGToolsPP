# verify_examples.ps1 - evaluate manual example formulas headless and compare to expected results.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent

# functions whose examples cannot run headless (macro-type / network / environment / volatile)
$skipFns = @('SHEET','SHEETS','FORMULATEXT','ISFORMULA','LET','IMAGE','RANDARRAY',
             'WEBSERVICE','GOOGLETRANSLATE','GTRS','IMPORTCSV','IMPORTDATA','IMPORTFEED',
             'IMPORTHTML','IMPORTRANGE','IMPORTTEXT','QUERY','EXRATE','RESTAPI',
             'DIRFOLDER','EVAL','SHEETLIST','TOTALPAGES')

# kr.json is the verification source: formulas match en.json, but localized labels
# (GROUPBY/PIVOTBY totals) follow the host UI language, which is Korean here.
$content = Get-Content (Join-Path $root 'manual\_content\kr.json') -Raw -Encoding UTF8 | ConvertFrom-Json

$errMap = @{
    -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!';
    -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!';
    -2146826246 = '#N/A';   -2146826245 = '#GETTING_DATA';
    -2146826241 = '#SPILL!'; -2146826238 = '#CALC!'
}

function Fmt($x) {
    if ($null -eq $x) { return '' }
    if ($x -is [double] -or $x -is [single]) { return $x.ToString([Globalization.CultureInfo]::InvariantCulture) }
    if ($x -is [bool]) { return $x.ToString().ToUpper() }
    if ($x -is [int] -and $errMap.ContainsKey([int]$x)) { return $errMap[[int]$x] }
    return [string]$x
}

# split an expected string like {"a",10;"b",20} or scalar into cell tokens (quote-aware)
function Tokens([string]$s) {
    $s = $s.Trim()
    if ($s.StartsWith('{') -and $s.EndsWith('}')) { $s = $s.Substring(1, $s.Length - 2) }
    $out = New-Object System.Collections.Generic.List[string]
    $cur = ''; $q = $false
    foreach ($ch in $s.ToCharArray()) {
        if ($q) { if ($ch -eq '"') { $q = $false } else { $cur += $ch }; continue }
        switch ($ch) {
            '"' { $q = $true }
            ',' { $out.Add($cur.Trim()); $cur = '' }
            ';' { $out.Add($cur.Trim()); $cur = '' }
            default { $cur += $ch }
        }
    }
    $out.Add($cur.Trim())
    return ,$out.ToArray()
}

function CellsEqual([string]$exp, [string]$got) {
    if ($exp -eq $got) { return $true }
    $de = 0.0; $dg = 0.0
    $ci = [Globalization.CultureInfo]::InvariantCulture
    if ([double]::TryParse($exp, [Globalization.NumberStyles]::Float, $ci, [ref]$de) -and
        [double]::TryParse($got, [Globalization.NumberStyles]::Float, $ci, [ref]$dg)) {
        # tolerate rounding in the expected value to its displayed precision
        $dec = 0; $dot = $exp.IndexOf('.')
        if ($dot -ge 0) { $dec = $exp.Length - $dot - 1 }
        if ($dec -gt 12) { $dec = 12 }
        return [math]::Round($dg, $dec) -eq [math]::Round($de, $dec)
    }
    return $false
}

$cases = @()
foreach ($p in $content.PSObject.Properties) {
    $fn = $p.Name
    if ($skipFns -contains $fn) { continue }
    $i = 0
    foreach ($e in @($p.Value.examples)) {
        $i++
        if ($null -eq $e -or [string]::IsNullOrEmpty([string]$e[1])) { continue }
        $cases += [pscustomobject]@{ fn = $fn; idx = $i; f = [string]$e[0]; exp = [string]$e[1] }
    }
}
Write-Output ("verifiable examples: " + $cases.Count)

Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep -Milliseconds 400
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$pass = 0; $fail = 0; $fails = @()
try {
    $xll = Join-Path $root 'dist\EGTools++64.xll'
    if (-not $xl.RegisterXLL($xll)) { throw "RegisterXLL failed: $xll" }
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)

    $row = 1
    foreach ($c in $cases) {
        $cell = $ws.Cells.Item($row, 1)
        # name ladder on this native host: EG.<FN>( first, then the formula as written
        $egForm = $c.f -replace ('^=(' + [regex]::Escape($c.fn) + ')\('), '=EG.$1('
        $got = $null
        foreach ($cand in @($egForm, $c.f)) {
            try { $cell.ClearContents(); $cell.Formula2 = $cand } catch { continue }
            $v = $cell.Value2
            if (($v -is [int]) -and ($v -eq -2146826259)) { continue }   # #NAME?
            $rng = $cell
            try { $sp = $cell.SpillingToRange; if ($sp) { $rng = $sp } } catch {}
            $vals = $rng.Value2
            $list = New-Object System.Collections.Generic.List[string]
            if ($vals -is [object[,]]) {
                for ($i = 1; $i -le $vals.GetLength(0); $i++) {
                    for ($j = 1; $j -le $vals.GetLength(1); $j++) { $list.Add((Fmt $vals[$i, $j])) }
                }
            } else { $list.Add((Fmt $vals)) }
            $got = $list
            break
        }
        if ($null -eq $got) {
            $fail++; $fails += ("{0} ex{1}: could not enter formula {2}" -f $c.fn, $c.idx, $c.f)
            $row += 12; continue
        }
        # scalar-first: a single-cell result may legally contain , ; or even braces
        # (ARRAYTOTEXT strict). Fall back to token-wise array comparison for spills.
        $ok = ($got.Count -eq 1 -and (CellsEqual $c.exp $got[0]))
        $expTok = @([string]$c.exp)
        if (-not $ok -and $c.exp.Trim().StartsWith('{')) {
            $expTok = Tokens $c.exp
            $ok = ($expTok.Count -eq $got.Count)
            if ($ok) { for ($k = 0; $k -lt $expTok.Count; $k++) { if (-not (CellsEqual $expTok[$k] $got[$k])) { $ok = $false; break } } }
        }
        if ($ok) { $pass++ }
        else {
            $fail++
            $fails += ("{0} ex{1}: {2}`n    expected [{3}] got [{4}]" -f $c.fn, $c.idx, $c.f, ($expTok -join '|'), ($got -join '|'))
        }
        $row += 12
        if ($row -gt 900) { $row = 1; $ws = $wb.Worksheets.Add() }
    }
    $wb.Close($false)
} finally {
    $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}
Write-Output ("RESULT: {0} passed, {1} failed" -f $pass, $fail)
foreach ($f in $fails) { Write-Output ("FAIL " + $f) }
