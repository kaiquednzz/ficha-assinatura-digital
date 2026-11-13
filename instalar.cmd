@echo off
rem Instala o suplemento Coleta (guia "Assinaturas" no Excel) neste PC. Rode com o Excel fechado.
tasklist /fi "imagename eq excel.exe" | find /i "excel.exe" >nul && (
  echo Feche o Excel e rode de novo.
  pause
  exit /b 1
)
copy /y "%~dp0Coleta.xlam" "%APPDATA%\Microsoft\AddIns\Coleta.xlam" >nul || (echo Falha ao copiar o Coleta.xlam. & pause & exit /b 1)
powershell -NoProfile -Command "$x=New-Object -ComObject Excel.Application; $w=$x.Workbooks.Add(); $a=$x.AddIns.Add($env:APPDATA+'\Microsoft\AddIns\Coleta.xlam',$true); $a.Installed=$true; $w.Close($false); $x.Quit(); Start-Sleep 2; Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force"
echo.
echo Pronto. Abra o Excel: vai aparecer a guia "Assinaturas".
pause
