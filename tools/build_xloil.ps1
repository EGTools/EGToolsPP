<#
.SYNOPSIS
  Fetch and statically build the xlOil libraries EGTools++ links against.

  *** DEVELOPER BUILD-TIME SCRIPT — NOT part of the distributed add-in. ***
  It downloads dependencies (git) and compiles them (MSBuild). Some antivirus
  products flag this "download → build" chain heuristically; on a dev machine,
  add this script / the project folder to the AV exclusions. The shipped .xll is
  unrelated to this script.

.DESCRIPTION
  EGTools++ links xlOil's static libraries (config: ReleaseStatic). xlOil is an
  MSBuild project (not CMake) and pulls several dependencies not vendored in its
  repo. This makes the whole thing reproducible:

    1. Clone xlOil (pinned) into external/xloil.
    2. Clone the dependency subset a static XLL needs (spdlog, tomlplusplus,
       winreg, asmjit). Python/SQL deps are skipped.
    3. Apply the EGTools++ patch (Collection<> explicit instantiation) needed
       for static linking (patches/).
    4. Mirror the Win32 xlcall32 import lib to the SDK LIB root (xlOil quirk).
    5. Write Directory.Build.props forcing the static CRT (/MT) so the single
       .xll needs no VC++ redistributable (plan O1).
    6. Build with MSBuild (two passes to avoid a parallel-build race in
       xlOil-COM's lib-combine step).

  Notes:
    * 32-bit uses the MSBuild platform name "Win32".

.PARAMETER Platform
  x64 (default) or Win32.

.EXAMPLE
  pwsh tools/build_xloil.ps1 -Platform x64
  pwsh tools/build_xloil.ps1 -Platform Win32
#>
[CmdletBinding()]
param(
  [ValidateSet('x64','Win32')] [string]$Platform = 'x64'
)
$ErrorActionPreference = 'Stop'

$XLOIL_COMMIT = 'a527de1742cf38f9721425aa9e1e42b18a8797bb'  # v0.22.1
$RepoRoot  = Split-Path -Parent $PSScriptRoot
$ExtDir    = Join-Path $RepoRoot 'external'
$XlOilDir  = Join-Path $ExtDir   'xloil'
$Patches   = @(
  (Join-Path $RepoRoot 'patches\xloil-0.22.1-collection-instantiation.patch')      # static-link Collection<> instantiation
  (Join-Path $RepoRoot 'patches\xloil-0.22.1-xladdinmanagerinfo-no-excel12.patch')  # x86/2016 Add-in Manager crash (Excel12 xlCoerce)
  (Join-Path $RepoRoot 'patches\xloil-0.22.1-no-spdlog.patch')                      # build with XLOIL_NO_SPDLOG (drop unused spdlog, ~-144KB)
  (Join-Path $RepoRoot 'patches\xloil-0.22.1-reload-safety.patch')                  # add-in untick/re-tick: Messenger class unregister + COM reconnect null-deref/CoUninit fix
  (Join-Path $RepoRoot 'patches\xloil-0.22.1-win32-thunk-args.patch')               # Win32 hand-rolled thunk: lifts the 16-arg UDF limit (variadic fns on x86)
)

function Find-MSBuild {
  $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
  $vs = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath
  $msb = Join-Path $vs 'MSBuild\Current\Bin\MSBuild.exe'
  if (-not (Test-Path $msb)) { throw "MSBuild not found under $vs" }
  return $msb
}

# --- 1. xlOil source --------------------------------------------------------
New-Item -ItemType Directory -Force $ExtDir | Out-Null
if (-not (Test-Path (Join-Path $XlOilDir '.git'))) {
  Write-Host "Cloning xlOil ($XLOIL_COMMIT)..."
  git clone https://gitlab.com/stevecu/xloil.git $XlOilDir
  Push-Location $XlOilDir; git checkout $XLOIL_COMMIT; Pop-Location
}

# --- 2. Dependencies (static-XLL subset) ------------------------------------
$deps = @(
  @{ name='spdlog';       url='https://github.com/gabime/spdlog.git';          br='v1.8.5' },
  @{ name='tomlplusplus'; url='https://github.com/marzer/tomlplusplus.git';    br='v3.4.0' },
  @{ name='winreg';       url='https://github.com/GiovanniDicanio/WinReg.git'; br='v6.1.0' }
)
$depRoot = Join-Path $XlOilDir 'external'
foreach ($d in $deps) {
  $dst = Join-Path $depRoot $d.name
  if (Test-Path $dst) { continue }
  Write-Host "Cloning $($d.name) $($d.br)..."
  git clone --single-branch --depth=1 --branch=$($d.br) $d.url $dst
}
# asmjit has no versioning; xlOil pins a 2019 commit.
$asmjit = Join-Path $depRoot 'asmjit'
if (-not (Test-Path $asmjit)) {
  Write-Host "Cloning asmjit (pinned)..."
  git clone https://github.com/asmjit/asmjit.git $asmjit
  Push-Location $asmjit; git reset --hard ac77dfcd75f043e2fe317133a971040e5b999916; Pop-Location
}

# --- 3. Apply EGTools++ patches (idempotent) --------------------------------
Push-Location $XlOilDir
foreach ($Patch in $Patches) {
  git apply --reverse --check $Patch 2>$null
  if ($LASTEXITCODE -eq 0) {
    Write-Host "Patch already applied: $(Split-Path -Leaf $Patch)"
  } else {
    Write-Host "Applying patch: $(Split-Path -Leaf $Patch)"
    git apply $Patch
    if ($LASTEXITCODE) { Pop-Location; throw "git apply failed: $Patch" }
  }
}
Pop-Location

# --- 3b. SDK lib layout fix (Win32) -----------------------------------------
# xlOil's Win32 static props look for xlcall32.lib in the SDK LIB root, but the
# bundled SDK only ships it under LIB\WIn32. Mirror it so xlOil's own projects
# (e.g. the example TestAddin) link. (x64 resolves fine via LIB\x64.)
$sdkLib = Join-Path $XlOilDir 'external\Excel2013SDK\LIB'
foreach ($f in 'XLCALL32.LIB','XLCALL32.DLL') {
  $src = Join-Path $sdkLib "WIn32\$f"; $dst = Join-Path $sdkLib $f
  if ((Test-Path $src) -and -not (Test-Path $dst)) { Copy-Item $src $dst }
}

# --- 3c. Force static CRT (/MT) + size optimization (plan O1 / size goal) ----
# /MT: single .xll, no VC++ redist. MinSpace+/Os: the .xll is a distribution
# artifact and UDF hot paths are dominated by Excel marshalling, so optimize the
# xlOil framework (incl. spdlog/asmjit) for size. Always rewritten so changes
# here propagate; MSBuild rebuilds when the .props timestamp changes.
$dbp = Join-Path $XlOilDir 'Directory.Build.props'
$dbpContent = @'
<?xml version="1.0" encoding="utf-8"?>
<!-- [EGTools++] Force static CRT (/MT) and size-optimized codegen (/O1-style)
     for Release/ReleaseStatic. Auto-imported by all xlOil projects. -->
<Project>
  <ItemDefinitionGroup Condition="'$(Configuration)'=='Release' or '$(Configuration)'=='ReleaseStatic'">
    <ClCompile>
      <RuntimeLibrary>MultiThreaded</RuntimeLibrary>
      <Optimization>MinSpace</Optimization>
      <FavorSizeOrSpeed>Size</FavorSizeOrSpeed>
      <InlineFunctionExpansion>OnlyExplicitInline</InlineFunctionExpansion>
      <FunctionLevelLinking>true</FunctionLevelLinking>
      <!-- Logging is never initialised by EGTools++ (no file/sink), so spdlog
           was dead weight (~144KB). Requires the no-spdlog patch. Remove this
           line (and relink spdlog.lib in cmake/xlOil.cmake) to restore logs
           temporarily for debugging. -->
      <PreprocessorDefinitions>XLOIL_NO_SPDLOG;%(PreprocessorDefinitions)</PreprocessorDefinitions>
    </ClCompile>
  </ItemDefinitionGroup>
</Project>
'@
if (-not (Test-Path $dbp) -or (Get-Content $dbp -Raw).TrimEnd() -ne $dbpContent.TrimEnd()) {
  Write-Host "Writing Directory.Build.props (static CRT /MT + size opt + no spdlog)..."
  $dbpContent | Set-Content -Path $dbp -Encoding UTF8
}

# --- 4. Build the solution in ReleaseStatic ---------------------------------
# The ReleaseStatic solution config builds exactly the static-XLL project set
# (Python/SQL excluded). xlOil-COM's pre-link lib-combine consumes the external
# libs without a project-reference edge, so a parallel first pass can race; a
# second pass always succeeds once the deps exist.
$msb = Find-MSBuild
$sln = Join-Path $XlOilDir 'xlOil.sln'
$common = @("/p:Configuration=ReleaseStatic", "/p:Platform=$Platform", '/m', '/v:minimal',
            '/clp:ErrorsOnly;Summary')

for ($attempt = 1; $attempt -le 2; $attempt++) {
  Write-Host "`n[Build pass $attempt] msbuild xlOil.sln (ReleaseStatic|$Platform)..."
  & $msb $sln @common
  if (-not $LASTEXITCODE) { break }
  if ($attempt -eq 2) { throw "xlOil build failed ($LASTEXITCODE)" }
  Write-Host "First pass hit the dependency race; retrying..."
}

$libDir = Join-Path $XlOilDir "build\$Platform\ReleaseStatic"
Write-Host "`nDone. Static libs in: $libDir"
Get-ChildItem $libDir -Filter *.lib | Select-Object Name,@{n='KB';e={[int]($_.Length/1KB)}} | Format-Table -AutoSize
