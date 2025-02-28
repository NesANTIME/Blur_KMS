@echo off
setlocal enabledelayedexpansion
title Blur_KMS v1.0 by NesAnTime
:inicio
cls
echo =======================================
echo       MENU DE OPCIONES (Blur_KMS)     
echo =======================================
echo.
echo [1] Instalar Python (Automatico - Via CMD)
echo [2] Instalar Python (Manual)
echo [3] Ejecutar Activador Blur_KMS
echo [4] Salir
echo ==========================
echo.
set /p opcion= Seleccione una opcion: 

if "%opcion%"=="1" goto instalar_python
if "%opcion%"=="2" goto instalar_py
if "%opcion%"=="3" goto ejecutar
if "%opcion%"=="4" goto salir

echo Opcion invalida. Elija Nuevamente.
pause
goto inicio

:instalar_python
echo Instalando Python...
echo.
timeout /t 3 >nul
python --version >nul 2>&1
if %errorlevel%==0 (
    echo Python ya se Encuentra instalado en tu sistema, Obmite este Paso!
    pause
    goto inicio
)

set "python_url=https://www.python.org/ftp/python/3.13.1/python-3.13.1-amd64.exe"
set "installer=python_installer.exe"

echo Descargando Python...
powershell -Command "(New-Object System.Net.WebClient).DownloadFile('%python_url%', '%installer%')"

if exist "%installer%" (
    echo Instalando Python...
    start /wait %installer% /quiet InstallAllUsers=1 PrependPath=1
    echo Instalación completada.
    del %installer%
) else (
    echo Error: No se pudo descargar el instalador.
)

python --version >nul 2>&1
if %errorlevel%==0 (
    echo Python se ha instalado correctamente.
) else (
    echo Hubo un problema con la instalación.
)

pause
goto inicio

:instalar_py
echo Abriendo Pagina Oficial de Python
timeout /t 3 >nul
start https://www.python.org/downloads/
pause
goto inicio

:ejecutar_script
if exist "Blur.py" (
    echo Ejecutando script.py...
    python Blur.py
) else (
    echo No se encontró script.py
)
pause
exit

:salir
echo Saliendo del programa...
exit
