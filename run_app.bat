@echo off
setlocal enableextensions enabledelayedexpansion

REM === Ir a la carpeta donde está este .bat ===
cd /d %~dp0

echo.
echo ================== Iniciando ==================
echo Carpeta: %CD%
echo.

REM === Detectar Python (py -3 o python) ===
where py >nul 2>&1
if %ERRORLEVEL%==0 (
    set "PYEXE=py -3"
) else (
    where python >nul 2>&1
    if %ERRORLEVEL%==0 (
        set "PYEXE=python"
    ) else (
        echo [ERROR] No se encontro Python en el sistema.
        echo Instala Python 3.9+ desde https://www.python.org y vuelve a intentar.
        pause
        exit /b 1
    )
)

REM === Crear entorno virtual si no existe ===
if not exist ".venv\Scripts\python.exe" (
    echo [+] Creando entorno virtual .venv ...
    %PYEXE% -m venv .venv
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
) else (
    echo [+] Entorno virtual .venv ya existe.
)

REM === Activar entorno virtual ===
call ".venv\Scripts\activate"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] No se pudo activar el entorno virtual.
    pause
    exit /b 1
)

REM === Crear requirements.txt si no existe ===
if not exist "requirements.txt" (
    echo [+] Creando requirements.txt ...
    > requirements.txt echo streamlit
    >> requirements.txt echo pandas
    >> requirements.txt echo openpyxl
) else (
    echo [+] Usando requirements.txt existente.
)

REM === Actualizar pip e instalar dependencias ===
echo.
echo [+] Actualizando pip e instalando dependencias ...
python -m pip install --upgrade pip
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Fallo la instalacion de dependencias.
    pause
    exit /b 1
)

REM === Lanzar la app de Streamlit ===
echo.
echo ================== Ejecutando la app ==================
echo (Cierra esta ventana para detener la app)
echo.
REM Si el puerto 8501 esta ocupado, cambia --server.port 8502
streamlit run app.py
REM streamlit run app.py --server.port 8502

endlocal
pause
