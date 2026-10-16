# Builds BSP_v1.xlsm (repository root) from the VBA modules in excel\src.
#
# Requirement (one-time, done by the user in Excel):
#   File > Options > Trust Center > Trust Center Settings > Macro Settings >
#   [x] Trust access to the VBA project object model
#
# Usage (PowerShell, from the repository folder):
#   powershell -ExecutionPolicy Bypass -File excel\build_xlsm.ps1
#   powershell -ExecutionPolicy Bypass -File excel\build_xlsm.ps1 -Test   # build + run the self-test
param([switch]$Test)

$ErrorActionPreference = 'Stop'
$repo   = Split-Path -Parent $PSScriptRoot
$src    = Join-Path $PSScriptRoot 'src'
$target = Join-Path $repo 'BSP_v1.xlsm'
$tmp    = Join-Path $env:TEMP 'BSP_build'
New-Item -ItemType Directory -Force $tmp | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# A VBA compile/runtime error opens a modal dialog in the hidden Excel and the COM
# call never returns. Every macro call runs under a watchdog that kills this Excel
# instance (only this one) after a timeout and then shows the VBA trace.
Add-Type -Namespace BspWin -Name User32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);'
$xlPid = [uint32]0
[void][BspWin.User32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$xlPid)
$trace = Join-Path $env:TEMP 'BSP_work\vba_trace.txt'
if (Test-Path $trace) { Remove-Item $trace -Force }

function Invoke-Macro([string]$name, [int]$seconds, [object[]]$macroArgs = @()) {
    $dog = Start-Process -FilePath powershell.exe -WindowStyle Hidden -PassThru -ArgumentList @(
        '-NoProfile', '-Command', "Start-Sleep -Seconds $seconds; Stop-Process -Id $xlPid -Force")
    try {
        switch ($macroArgs.Count) {
            0 { return $excel.Run($name) }
            3 { return $excel.Run($name, $macroArgs[0], $macroArgs[1], $macroArgs[2]) }
            4 { return $excel.Run($name, $macroArgs[0], $macroArgs[1], $macroArgs[2], $macroArgs[3]) }
        }
    } catch {
        $msg = "Macro $name tidak selesai dalam $seconds detik atau Excel berhenti (kemungkinan compile/runtime error VBA)."
        if (Test-Path $trace) { $msg += "`nLangkah terakhir VBA:`n" + ((Get-Content $trace | Select-Object -Last 8) -join "`n") }
        throw $msg
    } finally {
        Stop-Process -Id $dog.Id -Force -ErrorAction SilentlyContinue
    }
}

try {
    $wb = $excel.Workbooks.Add()
    $vbp = $null
    try { $vbp = $wb.VBProject; $null = $vbp.VBComponents.Count } catch { $vbp = $null }
    if ($null -eq $vbp) {
        $wb.Close($false)
        throw "Excel menolak akses ke VBA project. Aktifkan: File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model', lalu jalankan lagi."
    }

    foreach ($f in Get-ChildItem $src -Filter *.bas | Sort-Object Name) {
        # VBE imports ANSI text with CRLF line endings
        $text = [IO.File]::ReadAllText($f.FullName) -replace "`r?`n", "`r`n"
        $copy = Join-Path $tmp $f.Name
        [IO.File]::WriteAllText($copy, $text, [Text.Encoding]::GetEncoding(1252))
        $null = $vbp.VBComponents.Import($copy)
        Write-Host "Imported $($f.Name)"
    }

    if (Test-Path $target) { Remove-Item $target -Force }
    $wb.SaveAs($target, 52)          # 52 = xlOpenXMLWorkbookMacroEnabled
    Invoke-Macro "BSP_Setup_UI" 120 | Out-Null
    $wb.Save()
    Write-Host "Saved $target"

    if ($Test) {
        $pdfs = @(Get-ChildItem (Join-Path $repo 'input output sample'), (Join-Path $repo 'other bank statement sample') `
                  -Filter *.pdf -ErrorAction SilentlyContinue | ForEach-Object FullName)
        $out = Join-Path $tmp 'selftest_out'
        if (Test-Path $out) { Remove-Item $out -Recurse -Force }
        New-Item -ItemType Directory -Force $out | Out-Null
        Write-Host "`n--- Self-test per file ($($pdfs.Count) PDF) ---"
        Write-Host (Invoke-Macro "BSP_Selftest" 900 @(($pdfs -join '|'), $out, 'per_file'))
        Write-Host "--- Output files ---"
        Get-ChildItem $out | ForEach-Object { Write-Host $_.Name }
        Remove-Item (Join-Path $out '*') -Force
        Write-Host "`n--- Self-test gabung per rekening (nama selalu standar) ---"
        Write-Host (Invoke-Macro "BSP_Selftest" 900 @(($pdfs -join '|'), $out, 'consolidated', 'original'))
        Write-Host "--- Output files ---"
        Get-ChildItem $out | ForEach-Object { Write-Host $_.Name }
        $wb.Save()
    }
    $wb.Close($true)
}
finally {
    try { $excel.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    Stop-Process -Id $xlPid -Force -ErrorAction SilentlyContinue
}
