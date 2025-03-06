@echo off
setlocal enabledelayedexpansion
title Blur_KMS v2.0 by NesAnTime

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo        ! Es Nesesario Ejecutar Como Administrador...
    echo.
    echo    Solicitando Permiso! Acepte para Continuar...
    timeout /t 2 >nul
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
    exit /b
)

:inicio
cls
echo =======================================
echo       MENU DE OPCIONES (Blur_KMS Execute)     
echo =======================================
echo.
echo [1] Ejecutar Blur_KMS
echo [2] Instalar Python (Automatico - Via CMD)
echo [3] Instalar Python (Manual)
echo [4] Salir
echo ==========================
echo.
set /p opcion= Seleccione una opcion: 

if "%opcion%"=="1" goto ejecutar_script
if "%opcion%"=="2" goto instalar_python
if "%opcion%"=="3" goto instalar_py
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
    echo Python ya se Encuentra instalado en tu sistema.
    timeout /t 2 >nul
    echo [!] Ejecutando Blur_KMS...
    timeout /t 1 >nul
    goto ejecutar_script
    pause
)

set "python_url=https://www.python.org/ftp/python/3.13.1/python-3.13.2-amd64.exe"
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
cd /d "%~dp0"
if exist "Blur.py" (
    echo Verificando si colorama está instalado...
    python -c "import colorama" 2>nul
    if %errorlevel% neq 0 (
        echo Colorama no está instalado. Instalándolo ahora...
        python -m pip install colorama
    ) 
    echo Ejecutando Blur.py...
    python Blur.py
) else (
    echo No se encontró Blur.py
)
pause

:salir
echo Saliendo del programa...
exit
