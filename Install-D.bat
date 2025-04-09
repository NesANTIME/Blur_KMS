@echo off
setlocal enabledelayedexpansion
title Lanzador Blur_KMS v3.0 by NesAnTimes

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo                     Lanzador Blur_KMS v3.0
    echo.
    echo             Solicitando Permiso de Administrador
    echo                   Acepte para Continuar...
    timeout /t 3 >nul
    powershell -Command "Start-Process cmd -ArgumentList '/c %~s0' -Verb RunAs"
    exit /b
)

:Main
cls
echo.
echo.
echo =====================================================================================================
echo                                 MENU DE OPCIONES (Blur_KMS Lanzador v3.0)     
echo =====================================================================================================
echo.
echo       [1] Ejecutar Blur_KMS
echo       [2] Instalar Python (Manual)
echo       [3] Salir
echo.
echo -----------------------------------------------------------------------------------------------------
echo.
set /p opcion= -    -- Seleccione una opcion: 

if "%opcion%"=="1" goto ejecutar_blurkms
if "%opcion%"=="2" goto asistencia_py
if "%opcion%"=="3" goto salir

echo.
echo Opcion invalida. Elija Nuevamente.
pause
goto Main

:asistencia_py
echo.
echo    Abriendo Pagina Oficial de Python
timeout /t 3 >nul
start https://www.python.org/downloads/
pause
goto inicio

:ejecutar_blurkms
cd /d "%~dp0"
if exist "Blur.py" (
    echo _-_ Verificando si colorama & pillow están instalados...
    python -c "import colorama" 2>nul
    if %errorlevel% neq 0 (
        echo Colorama no está instalado. Instalándolo ahora...
        python -m pip install colorama pillow
    )
    python -c "import pillow" 2>nul
    if %errorlevel% neq 0 (
        echo Pillow no está instalado. Instalándolo ahora...
        python -m pip install pillow
    )
    echo.
    echo _-_ Ejecutando BlurKMS...
    start "BlurKMS v3.0 by NesAnTime" python Blur.py
    exit /b
) else (
    echo No se encontró Blur.py
    pause
    goto Main
)
pause

:salir
echo Saliendo del programa...
exit
