# extract_stored_names.ps1 — derive the internal (_xlfn.) names that modern
# Excel writes into .xlsx for the functions we shadow. Run on Excel 365/2021
# (where these functions are native). Output is the authoritative source for
# Version.cpp::storedName() — the legacy-compat alias table.
#
# Usage:  powershell -File tests\extract_stored_names.ps1
# Output: one "bareToken -> stored token" line per function.

$ErrorActionPreference = 'Stop'
$tmp = Join-Path $env:TEMP ("egtools_names_" + [IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Force -Path $tmp | Out-Null
$xlsx = Join-Path $tmp 'names.xlsx'

# Native modern formulas (minimal valid args; results don't matter, only the
# stored function token does).
$funcs = @(
 '=XLOOKUP(1,{1;2},{3;4})','=XMATCH(1,{1;2})','=FILTER({1;2},{1;0})','=SORT({2;1})',
 '=SORTBY({1;2},{2;1})','=UNIQUE({1;1;2})','=SEQUENCE(2)','=RANDARRAY(2)',
 '=IFS(TRUE,1)','=SWITCH(1,1,"a")','=MAXIFS({1;2},{1;1},1)','=MINIFS({1;2},{1;1},1)',
 '=TEXTJOIN(",",1,"a")','=CONCAT("a","b")','=TEXTSPLIT("a,b",",")','=TEXTBEFORE("a-b","-")',
 '=TEXTAFTER("a-b","-")','=VSTACK({1},{2})','=HSTACK({1},{2})','=TAKE({1;2},1)',
 '=DROP({1;2},1)','=TOROW({1,2})','=TOCOL({1,2})','=CHOOSEROWS({1;2},1)',
 '=CHOOSECOLS({1,2},1)','=EXPAND({1},2,1,0)','=WRAPROWS({1;2;3},2)','=WRAPCOLS({1;2;3},2)',
 '=ARRAYTOTEXT({1,2})','=VALUETOTEXT(1)','=GROUPBY({"x";"y"},{1;2},SUM)','=PIVOTBY({"x"},{"p"},{1},SUM)',
 '=REGEXTEST("a","a")','=REGEXEXTRACT("a1","\d")','=REGEXREPLACE("a1","\d","#")',
 '=IFNA(1,2)','=XOR(TRUE,FALSE)','=FORMULATEXT(A1)','=ISFORMULA(A1)','=SHEET()','=SHEETS()',
 '=UNICHAR(65)','=UNICODE("A")','=NUMBERVALUE("1")','=DAYS(2,1)','=ISOWEEKNUM(1)',
 '=ENCODEURL("a b")','=FILTERXML("<a/>","//a")'
)

Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep -Milliseconds 400
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1)
    $r = 2
    foreach ($f in $funcs) {
        try { $ws.Cells.Item($r, 2).Formula2 = $f }
        catch { try { $ws.Cells.Item($r, 2).Formula = $f } catch {} }
        $r++
    }
    $wb.SaveAs($xlsx, 51)   # xlOpenXMLWorkbook (.xlsx)
    $wb.Close($false)
} finally {
    $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}

$zip = Join-Path $tmp 'names.zip'
Copy-Item $xlsx $zip -Force
$ext = Join-Path $tmp 'unz'
Expand-Archive -Path $zip -DestinationPath $ext -Force
$xmlText = Get-Content (Join-Path $ext 'xl\worksheets\sheet1.xml') -Raw -Encoding UTF8

[regex]::Matches($xmlText, '<f[^>]*>(.*?)</f>') |
    ForEach-Object { $_.Groups[1].Value } |
    ForEach-Object { if ($_ -match '^(.+?)\(') { $matches[1] } else { $_ } } |
    Sort-Object -Unique

Remove-Item $tmp -Recurse -Force -EA SilentlyContinue
