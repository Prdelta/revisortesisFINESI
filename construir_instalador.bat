@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================================
echo  Instalador  -  Revisor de Tesis FINESI
echo ==========================================================
echo.
set "ISCC="
if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC (
  echo No se encontro Inno Setup 6. Descargalo de https://jrsoftware.org/isdl.php
  pause
  exit /b 1
)
echo [1/2] Construyendo el ejecutable...
rem Sin borrar dist\: ahi puede haber datos del revisor junto al .exe portable.
python -m PyInstaller RevisorTesisFINESI.spec --noconfirm
if errorlevel 1 (
  echo.
  echo La construccion del ejecutable fallo. Revisa los mensajes de arriba.
  pause
  exit /b 1
)
echo.
echo [2/2] Armando el instalador...
"%ISCC%" /Q "instalador\RevisorTesisFINESI.iss"
if errorlevel 1 (
  echo.
  echo Inno Setup no pudo armar el instalador. Revisa los mensajes de arriba.
  pause
  exit /b 1
)
echo.
echo Listo. El instalador esta en:  instalador\salida\
echo Ese archivo es el que se entrega: instala el programa con acceso directo y desinstalador.
echo.
start "" "%~dp0instalador\salida"
pause
