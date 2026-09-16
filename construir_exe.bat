@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ==========================================================
echo  Construccion del ejecutable  -  Revisor de Tesis FINESI
echo ==========================================================
echo.
python --version >nul 2>&1
if errorlevel 1 (
  echo No se encontro Python. Instalalo desde python.org y marca "Add Python to PATH".
  pause
  exit /b 1
)
echo [1/3] Instalando dependencias...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller
echo.
echo [2/3] Limpiando construcciones anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
echo.
echo [3/3] Construyendo. Esto tarda varios minutos...
python -m PyInstaller RevisorTesisFINESI.spec --noconfirm
if errorlevel 1 (
  echo.
  echo La construccion fallo. Revisa los mensajes de arriba.
  pause
  exit /b 1
)
echo.
echo Listo. El ejecutable esta en:  dist\RevisorTesisFINESI.exe
echo Ese es el programa. Se puede copiar a otra PC tal cual; no necesita Python instalado.
echo.
if exist dist start "" "%~dp0dist"
pause
