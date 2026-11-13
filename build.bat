@echo off
setlocal

echo ============================================
echo   Markdown Converter - Build Script
echo ============================================
echo.

echo [1/3] Limpiando artefactos anteriores...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
echo       OK

echo.
echo [2/3] Generando ejecutable con PyInstaller...
if exist .venv\Scripts\pyinstaller.exe (
    set "PYINSTALLER_CMD=.venv\Scripts\pyinstaller"
) else (
    set "PYINSTALLER_CMD=pyinstaller"
)
%PYINSTALLER_CMD% markdown-converter.spec

if %ERRORLEVEL% neq 0 (
    echo.
    echo ============================================
    echo   ERROR: El build fallo. Revisa la salida
    echo   de PyInstaller arriba para mas detalles.
    echo ============================================
    exit /b 1
)

echo.
echo [3/3] Verificando resultado...
if exist dist\markdown-converter.exe (
    echo       Exito: dist\markdown-converter.exe
    for %%I in (dist\markdown-converter.exe) do echo       Tamano: %%~zI bytes
) else (
    echo       ERROR: No se encontro dist\markdown-converter.exe
    exit /b 1
)

echo.
echo ============================================
echo   Build completado exitosamente!
echo   Ejecutable: dist\markdown-converter.exe
echo ============================================
endlocal
