@echo off
title Solucion Controles VB6 - OSBusiness
echo ==============================================
echo  SOLUCION: Registro de controles VB6 (26/07/2026)
echo  EJECUTAR COMO ADMINISTRADOR
echo ==============================================
echo.

cd /d "C:\OSBusiness FE"

echo 1. Copiando OCX original v6.1 a SysWOW64...
copy /Y "Support\crystl32.ocx" "C:\Windows\SysWOW64\crystl32.ocx"
echo    OK

echo.
echo 2. Desregistrando y registrando OCX...
C:\Windows\SysWOW64\regsvr32.exe /s /u "C:\Windows\SysWOW64\crystl32.ocx"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\crystl32.ocx"
echo    OK

echo.
echo 3. Registrando CLSID y ProgID...
reg add "HKCR\WOW6432Node\Crystal.CrystalReport" /ve /d "Crystal Report Control" /f >nul
reg add "HKCR\WOW6432Node\Crystal.CrystalReport\Clsid" /ve /d "{00025600-0000-0000-C000-000000000046}" /f >nul

reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}" /ve /d "Crystal Report Control" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\InprocServer32" /ve /d "C:\Windows\SysWow64\crystl32.ocx" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\ProgID" /ve /d "Crystal.CrystalReport" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Version" /ve /d "5.2" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\TypeLib" /ve /d "{00025600-0000-0000-C000-000000000046}" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Control" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Programmable" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\MiscStatus" /ve /d "0" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\ToolboxBitmap32" /ve /d "C:\Windows\SysWow64\crystl32.ocx, 1" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Implemented Categories\{0DE86A52-2BAA-11CF-A229-00AA003D7352}" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Implemented Categories\{0DE86A53-2BAA-11CF-A229-00AA003D7352}" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Implemented Categories\{0DE86A57-2BAA-11CF-A229-00AA003D7352}" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}" /f >nul
reg add "HKCR\WOW6432Node\CLSID\{00025600-0000-0000-C000-000000000046}\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}" /f >nul
echo    OK

echo.
echo 4. Configurando DSN ODBC (OSBusinessCR)...
reg add "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\OSBusinessCR" /ve /d "" /f >nul
reg add "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\OSBusinessCR" /v "Database" /t REG_SZ /d "BDATOS" /f >nul
reg add "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\OSBusinessCR" /v "Driver" /t REG_SZ /d "C:\WINDOWS\system32\SQLSRV32.dll" /f >nul
reg add "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\OSBusinessCR" /v "LastUser" /t REG_SZ /d "SYSTEM" /f >nul
reg add "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\OSBusinessCR" /v "Server" /t REG_SZ /d "SERVER" /f >nul
reg add "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\OSBusinessCR" /v "Trusted_Connection" /t REG_SZ /d "No" /f >nul
reg add "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources" /v "OSBusinessCR" /t REG_SZ /d "SQL Server" /f >nul
echo    OK

echo.
echo 5. Registrando controles VB6 adicionales...
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\MSFLXGRD.OCX"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\MSCOMCTL.OCX"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\MSCOMCT2.OCX"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\Comdlg32.ocx"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\RICHTX32.OCX"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\MSMASK32.OCX"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\TABCTL32.OCX"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\MSCHRT20.OCX"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\comctl32.ocx"
C:\Windows\SysWOW64\regsvr32.exe /s "C:\Windows\SysWOW64\MSRDO20.DLL"
echo    OK

echo.
echo 6. Eliminando cache .oca...
if exist "Support\crystl32.oca" del /f /q "Support\crystl32.oca"
echo    OK

echo.
echo ==============================================
echo  LISTO. Cierra VB6, vuelve a abrir primero.vbp
echo  y prueba imprimir desde frmDocu
echo ==============================================
pause
