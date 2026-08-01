# Generate Korean lunisolar table (918..2050) from .NET KoreanLunisolarCalendar.
# Output: C++ header rows { newYearSerial, monthMask(bit=30days), leapMonth, monthsInYear }
$ErrorActionPreference = 'Stop'
$cal = New-Object System.Globalization.KoreanLunisolarCalendar
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine("// KoreanLunarTable.h - generated from .NET KoreanLunisolarCalendar (918..2050).")
[void]$sb.AppendLine("// Row: lunar year 1/1 solar OADate serial, month-length mask (bit i = month i+1 has 30 days),")
[void]$sb.AppendLine("//      leap month index in 1..13 numbering (0 = none), months in year (12 or 13).")
[void]$sb.AppendLine("#pragma once")
[void]$sb.AppendLine("namespace egtools::functions {")
[void]$sb.AppendLine("struct LunarYear { int serial; unsigned short mask; unsigned char leap; unsigned char months; };")
[void]$sb.AppendLine("constexpr int kLunarFirstYear = 918, kLunarLastYear = 2050;")
[void]$sb.AppendLine("static const LunarYear kLunarYears[] = {")
for ($y = 918; $y -le 2050; $y++) {
    $leap = $cal.GetLeapMonth($y)
    $months = $cal.GetMonthsInYear($y)
    $mask = 0
    for ($m = 1; $m -le $months; $m++) {
        if ($cal.GetDaysInMonth($y, $m) -eq 30) { $mask = $mask -bor (1 -shl ($m - 1)) }
    }
    $ny = $cal.ToDateTime($y, 1, 1, 0, 0, 0, 0)
    $serial = [int]$ny.ToOADate()
    [void]$sb.AppendLine("    { $serial, $mask, $leap, $months },")
}
[void]$sb.AppendLine("};")
[void]$sb.AppendLine("}")
[IO.File]::WriteAllText("D:\EGWorkSpace\EGTools++\src\functions\KoreanLunarTable.h", $sb.ToString(), (New-Object System.Text.UTF8Encoding $false))
Write-Output "generated. sample checks:"
Write-Output ("2025 lunar new year: " + $cal.ToDateTime(2025,1,1,0,0,0,0).ToString("yyyy-MM-dd"))
Write-Output ("2025 leap month: " + $cal.GetLeapMonth(2025))
Write-Output ("2024-09-17 -> lunar: " + $cal.GetYear([datetime]"2024-09-17") + "-" + $cal.GetMonth([datetime]"2024-09-17") + "-" + $cal.GetDayOfMonth([datetime]"2024-09-17"))
