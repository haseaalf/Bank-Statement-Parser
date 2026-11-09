@echo off
rem One-click setup for Bank Statement Parser (BSP): double-click this file.
rem Installs R (if missing, without admin rights), the R packages, and prepares BSP_v1.xlsm.
rem Extra options can be passed on, e.g.  Install_BSP.bat -PrepareOffline
title Instalasi Bank Statement Parser
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup\install.ps1" %*
echo.
pause
