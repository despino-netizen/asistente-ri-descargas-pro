@echo off
setlocal

cd /d "%~dp0"

set "SCRIPT=DescargadorAutomatico.py"

if not exist "%SCRIPT%" (
    echo [ERROR] No se encontro "%SCRIPT%" en esta carpeta.
    echo Copia este .bat junto al archivo Python y vuelve a intentar.
    pause
    exit /b 1
)

set "PY_CMD="
where py >nul 2>&1
if %errorlevel%==0 set "PY_CMD=py -3"

if not defined PY_CMD (
    where python >nul 2>&1
    if %errorlevel%==0 set "PY_CMD=python"
)

if not defined PY_CMD (
    echo [ERROR] Python no esta instalado o no esta en PATH.
    echo Instala Python 3 y marca la opcion "Add Python to PATH".
    pause
    exit /b 1
)

echo Verificando dependencias...
%PY_CMD% -c "import customtkinter, selenium, webdriver_manager" >nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias necesarias...
    %PY_CMD% -m pip install --upgrade pip
    if errorlevel 1 (
        echo [ERROR] No se pudo actualizar pip.
        pause
        exit /b 1
    )

    %PY_CMD% -m pip install customtkinter selenium webdriver-manager
    if errorlevel 1 (
        echo [ERROR] No se pudieron instalar las dependencias.
        pause
        exit /b 1
    )
)

echo Iniciando %SCRIPT%...
%PY_CMD% "%SCRIPT%"
if errorlevel 1 (
    echo [ERROR] El script termino con error.
    pause
    exit /b 1
)

endlocal
