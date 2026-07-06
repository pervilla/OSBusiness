@echo off
echo ============================================
echo   REGISTRANDO TODOS LOS CONTROLES VB6
echo   Ejecutar como ADMINISTRADOR
echo ============================================
echo.

cd /d "C:\Windows\SysWOW64"

echo 1. Registrando OCX estandar...
regsvr32 /s MSCOMCTL.OCX
echo    MSCOMCTL.OCX
regsvr32 /s MSCOMCT2.OCX
echo    MSCOMCT2.OCX
regsvr32 /s MSFLXGRD.OCX
echo    MSFLXGRD.OCX
regsvr32 /s Comdlg32.ocx
echo    Comdlg32.OCX
regsvr32 /s RICHTX32.OCX
echo    RICHTX32.OCX
regsvr32 /s MSMASK32.OCX
echo    MSMASK32.OCX
regsvr32 /s TABCTL32.OCX
echo    TABCTL32.OCX
regsvr32 /s MSCHRT20.OCX
echo    MSCHRT20.OCX
regsvr32 /s crystl32.ocx
echo    crystl32.OCX
regsvr32 /s comctl32.ocx
echo    comctl32.OCX

echo.
echo 2. Registrando DLL de soporte...
regsvr32 /s MSRDO20.DLL
echo    MSRDO20.DLL

echo.
echo 3. Creando ProgID faltantes en registro 32-bit...
reg add "HKCR\WOW6432Node\Crystal.CrystalReport" /ve /d "Crystal Report Control" /f >nul
reg add "HKCR\WOW6432Node\Crystal.CrystalReport\Clsid" /ve /d "{00025601-0000-0000-C000-000000000046}" /f >nul
reg add "HKCR\WOW6432Node\RichTextLib.RichTextBox" /ve /d "Rich TextBox Control" /f >nul
reg add "HKCR\WOW6432Node\RichTextLib.RichTextBox\Clsid" /ve /d "{3B7C8860-D78F-101B-B9B5-04021C009402}" /f >nul
reg add "HKCR\WOW6432Node\TabDlg.SSTab" /ve /d "Microsoft Tabbed Dialog Control 6.0 (SP5)" /f >nul
reg add "HKCR\WOW6432Node\TabDlg.SSTab\Clsid" /ve /d "{BDC217C5-ED16-11CD-956C-0000C04E4C0A}" /f >nul
echo    Crystal.CrystalReport, RichTextLib.RichTextBox, TabDlg.SSTab

echo.
echo 4. Eliminando InprocServer32 corrupto de MSCOMCTL...
reg delete "HKCR\WOW6432Node\CLSID\{1EFB6596-857C-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{2C247F23-8591-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{35053A22-8589-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{66833FE6-8583-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{8E3867A3-8586-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{BDD1F04B-858B-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{C74190B6-8589-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{DD9DA666-8594-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
reg delete "HKCR\WOW6432Node\CLSID\{F08DF954-8592-11D1-B16A-00C0F0283628}\InprocServer32" /v InprocServer32 /f >nul 2>&1
echo    9 CLSID corruptos limpiados

echo.
echo ============================================
echo   LISTO. Cierra VB6 y abre primero.vbp
echo ============================================
pause
