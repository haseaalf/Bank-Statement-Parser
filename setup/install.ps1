# One-click setup for Bank Statement Parser (BSP). Started by Install_BSP.bat.
#
#  1. Checks Windows / Excel
#  2. Finds R; installs R (default 4.2.2, the tested version) for the current user if missing
#     - no administrator rights needed
#  3. Installs the R packages into BSP_modular_engine\r-library\<R version> (online or offline)
#  4. Removes the "downloaded from the internet" block from the program files (so Excel
#     does not block the macros of BSP_v1.xlsm)
#  5. Writes bsp_config.txt (paths used by the workbook) and a desktop shortcut
#  6. Self-check: runs the engine once
#
# Options:
#   -RVersion 4.2.2     R version to install when R is missing
#   -PrepareOffline     only download the R installer + packages into setup\offline
#                       (copy the whole folder to a PC without internet, then run Install_BSP.bat there)
#   -NoShortcut         do not create the desktop shortcut
param([string]$RVersion = "4.2.2", [switch]$PrepareOffline, [switch]$NoShortcut)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$root    = Split-Path -Parent $PSScriptRoot
$engine  = Join-Path $root 'BSP_modular_engine'
$offline = Join-Path $PSScriptRoot 'offline'
$log     = Join-Path $PSScriptRoot 'install_log.txt'
$cran    = 'https://cloud.r-project.org'
$packages = @('pdftools', 'openxlsx', 'zip')

try { Start-Transcript -Path $log -Force | Out-Null } catch { }

function Step($n, $text) { Write-Host ""; Write-Host "[$n/6] $text" -ForegroundColor Cyan }
function Ok($text)   { Write-Host "      OK  $text" -ForegroundColor Green }
function Warn($text) { Write-Host "      !!  $text" -ForegroundColor Yellow }
function Fail($text) { Write-Host "      GAGAL  $text" -ForegroundColor Red; throw $text }

function Find-Rscript {
    $cands = @()
    foreach ($hive in 'HKCU:', 'HKLM:') {
        foreach ($k in 'Software\R-core\R', 'Software\R-core\R64') {
            $key = "$hive\$k"
            if (Test-Path $key) {
                Get-ChildItem $key -ErrorAction SilentlyContinue | ForEach-Object {
                    $p = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).InstallPath
                    if ($p) { $cands += (Join-Path $p 'bin\Rscript.exe') }
                }
            }
        }
    }
    foreach ($base in "$env:ProgramFiles\R", "$env:LOCALAPPDATA\Programs\R", "$env:USERPROFILE\Documents\R") {
        if (Test-Path $base) { Get-ChildItem $base -Directory -Filter 'R-*' | ForEach-Object { $cands += (Join-Path $_.FullName 'bin\Rscript.exe') } }
    }
    $found = $cands | Where-Object { Test-Path $_ } | Select-Object -Unique | ForEach-Object {
        $v = [regex]::Match($_, 'R-(\d+)\.(\d+)\.(\d+)')
        [pscustomobject]@{ Path = $_; Version = if ($v.Success) { [version]"$($v.Groups[1]).$($v.Groups[2]).$($v.Groups[3])" } else { [version]'0.0.0' } }
    } | Where-Object { $_.Version -ge [version]'4.1.0' } | Sort-Object Version -Descending
    return $found | Select-Object -First 1
}

function R-Installer-Url($ver) {
    # latest release lives in base/, older ones in base/old/<ver>/
    $latest = (Invoke-WebRequest -UseBasicParsing "$cran/bin/windows/base/release.html").Content
    if ($latest -match "R-$([regex]::Escape($ver))-win\.exe") { return "$cran/bin/windows/base/R-$ver-win.exe" }
    return "$cran/bin/windows/base/old/$ver/R-$ver-win.exe"
}

function Run-R($rscript, $code) {
    $ErrorActionPreference = 'Continue'      # R writes progress/warnings to stderr
    $tmp = Join-Path $env:TEMP ("bsp_setup_" + [guid]::NewGuid().ToString('N') + '.R')
    [IO.File]::WriteAllText($tmp, $code, (New-Object Text.UTF8Encoding $false))
    try {
        & $rscript --no-save --no-restore $tmp 2>&1 | ForEach-Object {
            $line = if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { "$_" }
            if ($line -and $line -ne 'System.Management.Automation.RemoteException') { Write-Host "      $line" }
        }
        return $LASTEXITCODE
    }
    finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

function R-Version-MM($rscript) {
    # "4.2" - written to a temp file because PowerShell 5 mangles quotes in native arguments
    $ErrorActionPreference = 'Continue'
    $tmp = Join-Path $env:TEMP ("bsp_setup_" + [guid]::NewGuid().ToString('N') + '.R')
    [IO.File]::WriteAllText($tmp, "cat(R.version`$major, sub('[.].*', '', R.version`$minor), sep = '.')")
    try { return ((& $rscript --no-save --no-restore $tmp 2>$null) -join '').Trim() }
    finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

Write-Host "============================================================"
Write-Host "  Instalasi Bank Statement Parser (BSP)"
Write-Host "  Folder program: $root"
Write-Host "============================================================"

try {
    # ---------------------------------------------------------------- offline bundle
    if ($PrepareOffline) {
        Step 1 "Menyiapkan paket offline di $offline"
        New-Item -ItemType Directory -Force $offline | Out-Null
        $url = R-Installer-Url $RVersion
        $exe = Join-Path $offline "R-$RVersion-win.exe"
        if (-not (Test-Path $exe)) { Write-Host "      Mengunduh $url"; Invoke-WebRequest -UseBasicParsing $url -OutFile $exe }
        Ok "Installer R: $exe"
        $r = Find-Rscript
        if (-not $r) { Fail "Untuk mengunduh paket R, PC ini perlu R terpasang. Jalankan Install_BSP.bat biasa dulu." }
        $mm = $RVersion.Substring(0, $RVersion.LastIndexOf('.'))
        $code = @"
pk <- c($(($packages | ForEach-Object { "'$_'" }) -join ', '))
repo <- '$cran'
ap <- available.packages(contriburl = contrib.url(repo, type = 'win.binary'), type = 'win.binary')
deps <- unique(c(pk, unlist(tools::package_dependencies(pk, db = ap, which = c('Depends','Imports','LinkingTo'), recursive = TRUE))))
deps <- setdiff(deps, rownames(installed.packages(priority = 'base')))
dir.create('$($offline -replace '\\','/')/packages-$mm', showWarnings = FALSE)
download.packages(deps, destdir = '$($offline -replace '\\','/')/packages-$mm', contriburl = '$cran/bin/windows/contrib/$mm', type = 'win.binary')
"@
        if ((Run-R $r.Path $code) -ne 0) { Fail "Gagal mengunduh paket R" }
        Ok "Paket R untuk R $mm disimpan di $offline\packages-$mm"
        Write-Host ""
        Write-Host "Salin seluruh folder program ke PC tujuan lalu jalankan Install_BSP.bat di sana." -ForegroundColor Green
        return
    }

    # ---------------------------------------------------------------- 1. Windows / Excel
    Step 1 "Memeriksa Windows dan Microsoft Excel"
    Ok ("Windows " + [Environment]::OSVersion.Version + ", PowerShell " + $PSVersionTable.PSVersion)
    $excel = @("$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
               "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE",
               "$env:ProgramFiles\Microsoft Office\Office16\EXCEL.EXE",
               "${env:ProgramFiles(x86)}\Microsoft Office\Office16\EXCEL.EXE") | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $excel) { $excel = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe' -ErrorAction SilentlyContinue).'(default)' }
    if ($excel) { Ok "Excel: $excel" } else { Warn "Excel tidak ditemukan - BSP_v1.xlsm membutuhkan Microsoft Excel (Windows)" }
    if (-not (Test-Path (Join-Path $engine 'BSP_engine.R'))) { Fail "Folder BSP_modular_engine tidak lengkap. Ekstrak ulang seluruh isi program." }

    # ---------------------------------------------------------------- 2. R
    Step 2 "Memeriksa R"
    $r = Find-Rscript
    if ($r) {
        Ok "R $($r.Version) sudah terpasang: $($r.Path)"
    } else {
        Warn "R belum terpasang - memasang R $RVersion untuk user ini (tanpa hak administrator)"
        $exe = Join-Path $offline "R-$RVersion-win.exe"
        if (-not (Test-Path $exe)) {
            $dl = Join-Path $env:TEMP 'BSP_setup'; New-Item -ItemType Directory -Force $dl | Out-Null
            $exe = Join-Path $dl "R-$RVersion-win.exe"
            $url = R-Installer-Url $RVersion
            Write-Host "      Mengunduh $url (sekitar 75 MB)..."
            Invoke-WebRequest -UseBasicParsing $url -OutFile $exe
        } else { Write-Host "      Memakai installer offline $exe" }
        $dir = Join-Path $env:LOCALAPPDATA "Programs\R\R-$RVersion"
        Write-Host "      Memasang ke $dir ..."
        $p = Start-Process -FilePath $exe -Wait -PassThru -ArgumentList @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/CURRENTUSER', "/DIR=`"$dir`"")
        if ($p.ExitCode -ne 0) { Fail "Installer R berhenti dengan kode $($p.ExitCode)" }
        $r = Find-Rscript
        if (-not $r) { Fail "R sudah dipasang tetapi Rscript.exe tidak ditemukan" }
        Ok "R $($r.Version) terpasang: $($r.Path)"
    }
    $rscript = $r.Path

    # ---------------------------------------------------------------- 3. packages
    Step 3 "Memasang paket R ($($packages -join ', '))"
    $mm = R-Version-MM $rscript
    if ($mm -notmatch '^\d+\.\d+$') { Fail "Versi R tidak terbaca dari $rscript" }
    $lib = Join-Path $engine "r-library\$mm"
    New-Item -ItemType Directory -Force $lib | Out-Null
    $libR = $lib -replace '\\', '/'
    $pkgR = "c(" + (($packages | ForEach-Object { "'$_'" }) -join ', ') + ")"
    $local = Join-Path $offline "packages-$mm"
    if (Test-Path $local) {
        Write-Host "      Memakai paket offline: $local"
        $code = "zips <- list.files('$($local -replace '\\','/')', pattern = '[.]zip`$', full.names = TRUE)`n" +
                "install.packages(zips, lib = '$libR', repos = NULL, type = 'win.binary')`n"
    } else {
        $code = "pk <- $pkgR`n.libPaths(c('$libR', .libPaths()))`n" +
                "miss <- pk[!vapply(pk, requireNamespace, logical(1), quietly = TRUE)]`n" +
                "if (length(miss)) install.packages(miss, lib = '$libR', repos = '$cran', type = 'win.binary', dependencies = c('Depends','Imports','LinkingTo'))`n"
    }
    $code += ".libPaths(c('$libR', .libPaths()))`nfor (p in $pkgR) suppressPackageStartupMessages(library(p, character.only = TRUE))`n" +
             "cat('Versi:', paste(sapply($pkgR, function(p) paste(p, packageVersion(p))), collapse = ', '), '\n')`n"
    if ((Run-R $rscript $code) -ne 0) { Fail "Paket R gagal dipasang (cek koneksi internet / proxy kantor, atau pakai paket offline)" }
    Ok "Paket R siap di $lib"

    # ---------------------------------------------------------------- 4. unblock
    Step 4 "Melepas blokir 'file dari internet' pada file program"
    $n = 0
    $targets = @('BSP_v1.xlsm', 'Install_BSP.bat', 'BSP_modular_engine', 'setup', 'excel', 'tools', 'docs') |
               ForEach-Object { Join-Path $root $_ } | Where-Object { Test-Path $_ }
    Get-ChildItem $targets -Recurse -File | ForEach-Object {
        if (Get-Item $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue) { Unblock-File $_.FullName; $n++ }
    }
    Ok "$n file dilepas blokirnya (macro BSP_v1.xlsm tidak lagi diblokir Excel)"

    # ---------------------------------------------------------------- 5. config + shortcut
    Step 5 "Menyimpan pengaturan dan membuat shortcut"
    $cfg = Join-Path $root 'bsp_config.txt'
    @("# Dibuat oleh Install_BSP.bat - dibaca BSP_v1.xlsm saat dibuka",
      "RSCRIPT_DIR=$(Split-Path -Parent $rscript)",
      "ENGINE_DIR=$engine",
      "R_VERSION=$($r.Version)",
      "INSTALLED=$(Get-Date -Format 'yyyy-MM-dd HH:mm')") | Set-Content -Path $cfg -Encoding ASCII
    Ok "Pengaturan: $cfg"
    $xlsm = Join-Path $root 'BSP_v1.xlsm'
    if (-not $NoShortcut -and (Test-Path $xlsm)) {
        $lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Bank Statement Parser.lnk'
        $sh = (New-Object -ComObject WScript.Shell).CreateShortcut($lnk)
        $sh.TargetPath = $xlsm; $sh.WorkingDirectory = $root; $sh.Description = 'Bank Statement Parser (BSP)'
        $sh.Save()
        Ok "Shortcut desktop: $lnk"
    }

    # ---------------------------------------------------------------- 6. self-check
    Step 6 "Uji coba engine"
    $work = Join-Path $env:TEMP 'BSP_setup_check'
    New-Item -ItemType Directory -Force $work | Out-Null
    $done = Join-Path $work 'engine_done.txt'
    if (Test-Path $done) { Remove-Item $done -Force }
    $ErrorActionPreference = 'Continue'
    & $rscript (Join-Path $engine 'BSP_engine.R') $work $engine 'list_banks' 2>&1 | Out-Null
    $ErrorActionPreference = 'Stop'
    $status = if (Test-Path $done) { (Get-Content $done)[0] } else { 'TIDAK SELESAI' }
    if ($status -ne 'OK') { Fail "Engine gagal berjalan ($status). Lihat $work\engine_log.txt" }
    $banks = (Import-Csv (Join-Path $work 'bank_registry.csv')).Count
    Ok "Engine berjalan, $banks format bank tersedia"

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Green
    Write-Host "  INSTALASI SELESAI" -ForegroundColor Green
    Write-Host "  Buka 'Bank Statement Parser' di desktop (atau BSP_v1.xlsm)," -ForegroundColor Green
    Write-Host "  klik 'Enable Content' bila diminta, lalu 'Perbarui daftar bank'." -ForegroundColor Green
    Write-Host "============================================================" -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "INSTALASI GAGAL: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Log lengkap: $log" -ForegroundColor Red
    $global:LASTEXITCODE = 1
}
finally {
    try { Stop-Transcript | Out-Null } catch { }
}
exit $LASTEXITCODE
