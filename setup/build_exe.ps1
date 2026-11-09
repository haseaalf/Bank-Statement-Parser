# Builds BSP.exe (program folder) from app\BSP.cs with the C# compiler that ships with
# Windows (.NET Framework 4.x). Nothing has to be installed, and the exe is created
# locally, so Windows SmartScreen does not treat it as a downloaded file.
#
# Called by Install_BSP.bat; can also be run on its own:  Build_BSP_exe.bat
param([switch]$Quiet)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$src  = Join-Path $root 'app\BSP.cs'
$exe  = Join-Path $root 'BSP.exe'

$csc = @("$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
         "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $csc) { throw "Compiler .NET Framework 4 (csc.exe) tidak ditemukan di Windows ini." }
if (-not (Test-Path $src)) { throw "Kode sumber $src tidak ditemukan." }

$refs = 'System.dll', 'System.Core.dll', 'System.Drawing.dll', 'System.Windows.Forms.dll'
$cargs = @('/nologo', '/target:winexe', '/optimize+', "/out:$exe") + ($refs | ForEach-Object { "/reference:$_" })
$icon = Join-Path $root 'app\BSP.ico'
if (Test-Path $icon) { $cargs += "/win32icon:$icon" }
$cargs += $src

$out = & $csc @cargs 2>&1
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $exe)) {
    $out | ForEach-Object { Write-Host "      $_" }
    throw "Kompilasi BSP.exe gagal."
}
if (-not $Quiet) { Write-Host "BSP.exe dibuat: $exe" }
$exe
