@echo off
echo Registrando todas las librerias de OSBusiness FE...
echo.

cd /d "E:\DOCUMENTOS\OSBusiness FE\Support"

echo Registrando DLLs...
for %%f in (*.dll) do (
    echo Registrando %%f...
    C:\Windows\SysWOW64\regsvr32.exe /s "%%f"
)

echo.
echo Registrando OCXs...
for %%f in (*.ocx) do (
    echo Registrando %%f...
    C:\Windows\SysWOW64\regsvr32.exe /s "%%f"
)

echo.
echo Proceso completado.
pause