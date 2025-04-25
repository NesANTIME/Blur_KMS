@echo off
setlocal enabledelayedexpansion
title Blur_KMS v3.4 Launcher by NesAnTimes

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo                     Lanzador Blur_KMS v3.4
    echo.
    echo             Solicitando Permiso de Administrador
    echo                   Acepte para Continuar...
    timeout /t 3 >nul
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

:Main
cls
echo.
echo.
echo =====================================================================================================
echo                                 MENU DE OPCIONES (Blur_KMS Lanzador v3.4)     
echo =====================================================================================================
echo.
echo       [1] Ejecutar Blur_KMS
echo       [2] Instalador Python (Automatico)
echo       [3] Instalar Librerias (Automatico)
echo       [4] Salir
echo.
echo -----------------------------------------------------------------------------------------------------
echo.
set /p opcion= -    -- Seleccione una opcion:  

if "%opcion%"=="1" goto ejecutar_blurkms
if "%opcion%"=="2" goto instalar_python
if "%opcion%"=="3" goto instalar_librerias
if "%opcion%"=="4" goto salir

echo.
echo Opción inválida. Intenta de nuevo.
pause
goto Main

:instalar_python
echo.
echo.
echo      [#] Verificando si Python esta instalado...
echo.
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo      [##] Python no esta instalado. Descargando e instalando...
    powershell -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.12.2/python-3.12.2-amd64.exe -OutFile python-installer.exe"
    if exist python-installer.exe (
        start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
        del python-installer.exe
        echo      [+] Python instalado correctamente.
    ) else (
        echo      [X] Error al descargar el instalador de Python.
    )
) else (
    echo      [+] Python ya esta instalado.
)
echo.
pause
goto Main

:instalar_librerias
echo.
echo      [#] Instalando librerias desde requirements.txt...
echo.

python -m pip install --upgrade pip
python -m pip install -r colorama pillow
echo      [+] Librerias instaladas correctamente (si no viste errores arriba).
echo.
pause
goto Main

:ejecutar_blurkms
cls
cd /d "%~dp0"
if not exist Blur.py (
    echo      [x] El archivo Blur.py no se encontro en la carpeta actual.
    pause
    goto Main
)

echo      [#] Verificando dependencias necesarias...

python -c "import colorama" 2>nul
if %errorlevel% neq 0 (
    echo       [###] - Colorama no esta instalado. Instalandolo...
    python -m pip install colorama
)

python -c "import PIL" 2>nul
if %errorlevel% neq 0 (
    echo       [###] - Pillow no esta instalado. Instalandolo...
    python -m pip install pillow
)

echo.
echo       [#] Ejecutando Blur_KMS...
start "Blur_KMS v3.4 by NesAnTimes" python Blur.py
exit /b

:salir
echo       [##] Cerrando el launcher...
timeout /t 2 >nul
exit
