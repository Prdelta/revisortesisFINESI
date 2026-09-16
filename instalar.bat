@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Instalando dependencias del revisor...
python --version >nul 2>&1
if errorlevel 1 (
  echo.
  echo No se encontro Python. Instalalo desde https://www.python.org/downloads/
  echo IMPORTANTE: marca la casilla "Add Python to PATH" durante la instalacion.
  pause
  exit /b 1
)
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
echo.
java -version >nul 2>&1
if errorlevel 1 (
  echo AVISO: no se encontro Java. Sin Java el revisor solo revisa ortografia con diccionario, sin gramatica.
  echo Para gramatica instala Java 17 o superior: https://adoptium.net
) else (
  echo Java encontrado. La primera revision descargara LanguageTool, unos 260 MB. Solo pasa una vez.
)
echo.
echo Listo.
pause
