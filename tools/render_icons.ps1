# render_icons.ps1 — resources/icons/svg/*.svg → resources/icons/png/<name>_32.png
#
# 외부 SVG 렌더러(inkscape/resvg 등) 의존 없이 Windows 내장 WPF(PresentationCore)로
# 렌더링한다(멀티부팅 양쪽 OS에서 재현 가능 — plan/12 원칙). 지원하는 SVG 부분집합:
#   <path d>(WPF 경로 미니언어와 호환), <rect x y width height rx>, <circle cx cy r>,
#   <line x1 y1 x2 y2>  + 속성 fill / stroke / stroke-width / stroke-linecap /
#   stroke-linejoin. 아이콘 원본은 이 부분집합 안에서만 작성한다(디자인 SoT는 SVG).
#
# 예외: about_64.png는 SVG가 아니라 EG 로고 래스터(resources/icons/EGlogo.png,
# 256px 원본)를 64px로 축소한 정적 자산이다 — 이 스크립트가 생성하지 않는다.
#
# 사용: powershell -ExecutionPolicy Bypass -File tools\render_icons.ps1 [-Sizes 64,32]
# 기본 64px — 고DPI 리본에서 Office가 축소 방향으로 스케일링해 선명하다
# (32px 임베드는 125~150% 배율에서 확대되며 계단 현상, 2026-08-04 사용자 피드백).
param([int[]]$Sizes = @(64))

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationCore, WindowsBase

$root   = Split-Path -Parent $PSScriptRoot
$svgDir = Join-Path $root 'resources\icons\svg'
$pngDir = Join-Path $root 'resources\icons\png'
New-Item -ItemType Directory -Force $pngDir | Out-Null

function New-Brush([string]$spec) {
    if (-not $spec -or $spec -eq 'none') { return $null }
    $b = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($spec))
    $b.Freeze(); return $b
}

function New-Pen($el) {
    $stroke = $el.GetAttribute('stroke')
    if (-not $stroke -or $stroke -eq 'none') { return $null }
    $w = 1.0; $ws = $el.GetAttribute('stroke-width'); if ($ws) { $w = [double]$ws }
    $pen = New-Object System.Windows.Media.Pen ((New-Brush $stroke), $w)
    $cap = $el.GetAttribute('stroke-linecap')
    if ($cap -eq 'round') {
        $pen.StartLineCap = 'Round'; $pen.EndLineCap = 'Round'
    }
    $join = $el.GetAttribute('stroke-linejoin')
    if ($join -eq 'round') { $pen.LineJoin = 'Round' }
    $pen.Freeze(); return $pen
}

function Get-Geometry($el) {
    switch ($el.LocalName) {
        'path'   { return [System.Windows.Media.Geometry]::Parse($el.GetAttribute('d')) }
        'line'   { return New-Object System.Windows.Media.LineGeometry (
                       (New-Object System.Windows.Point ([double]$el.GetAttribute('x1'), [double]$el.GetAttribute('y1'))),
                       (New-Object System.Windows.Point ([double]$el.GetAttribute('x2'), [double]$el.GetAttribute('y2')))) }
        'circle' { return New-Object System.Windows.Media.EllipseGeometry (
                       (New-Object System.Windows.Point ([double]$el.GetAttribute('cx'), [double]$el.GetAttribute('cy'))),
                       [double]$el.GetAttribute('r'), [double]$el.GetAttribute('r')) }
        'rect'   { $rx = 0.0; if ($el.GetAttribute('rx')) { $rx = [double]$el.GetAttribute('rx') }
                   $r = New-Object System.Windows.Rect ([double]$el.GetAttribute('x'), [double]$el.GetAttribute('y'),
                                                        [double]$el.GetAttribute('width'), [double]$el.GetAttribute('height'))
                   return New-Object System.Windows.Media.RectangleGeometry ($r, $rx, $rx) }
        default  { return $null }
    }
}

foreach ($svgFile in Get-ChildItem $svgDir -Filter *.svg) {
    [xml]$doc = Get-Content $svgFile.FullName -Raw -Encoding UTF8
    $vb = $doc.DocumentElement.GetAttribute('viewBox') -split '\s+'
    $vbW = [double]$vb[2]; $vbH = [double]$vb[3]

    foreach ($size in $Sizes) {
        $visual = New-Object System.Windows.Media.DrawingVisual
        $dc = $visual.RenderOpen()
        $sx = $size / $vbW; $sy = $size / $vbH
        $scale = New-Object System.Windows.Media.ScaleTransform ($sx, $sy)
        $dc.PushTransform($scale)
        foreach ($el in $doc.DocumentElement.ChildNodes) {
            if ($el.NodeType -ne 'Element') { continue }
            $geo = Get-Geometry $el
            if ($null -eq $geo) { continue }
            $fillSpec = $el.GetAttribute('fill')
            $fill = $null
            # SVG 기본 fill은 black이지만 이 파이프라인에선 명시된 fill만 칠한다(모노라인 아이콘 전제)
            if ($fillSpec -and $fillSpec -ne 'none') { $fill = New-Brush $fillSpec }
            $dc.DrawGeometry($fill, (New-Pen $el), $geo)
        }
        $dc.Pop(); $dc.Close()

        $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap ($size, $size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
        $rtb.Render($visual)
        $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $out = Join-Path $pngDir ("{0}_{1}.png" -f $svgFile.BaseName, $size)
        $fs = [System.IO.File]::Create($out)
        $enc.Save($fs); $fs.Close()
        Write-Output ("rendered {0} ({1}x{1})" -f (Split-Path $out -Leaf), $size)
    }
}
