@echo off
chcp 65001 >nul
title Conversor REGATI TXT para XLSX
cd /d "%~dp0"

where python >nul 2>&1
if %errorlevel% equ 0 (
    set PY=python
) else (
    where py >nul 2>&1
    if %errorlevel% equ 0 (
        set PY=py
    ) else (
        echo Python nao encontrado. Instale o Python ou adicione ao PATH.
        pause
        exit /b 1
    )
)

echo Instalando dependencias (se necessario)...
%PY% -m pip install -r requirements.txt -q
if errorlevel 1 (
    echo AVISO: Falha ao instalar dependencias. Tentando executar mesmo assim...
)
echo.

if "%~1"=="" (
    echo Abrindo interface para selecionar o arquivo TXT...
    %PY% "txt_para_xlsx.py"
) else (
    echo Convertendo: %~1
    %PY% "txt_para_xlsx.py" "%~1"
)

if errorlevel 1 (
    echo.
    echo Ocorreu um erro. Veja a mensagem acima.
    pause
    exit /b 1
)
echo.
echo Concluido. Pode fechar esta janela.
pause
