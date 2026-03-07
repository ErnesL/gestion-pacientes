@echo off
setlocal

cd /d "%~dp0.."

if not exist "dist\GestionPacientes\GestionPacientes.exe" (
  echo No existe el build de la app. Ejecuta primero scripts\build_windows_app.bat
  exit /b 1
)

set "ISCC_EXE=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC_EXE%" set "ISCC_EXE=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if not exist "%ISCC_EXE%" (
  echo No se encontro Inno Setup 6. Instala Inno Setup y vuelve a ejecutar este script.
  exit /b 1
)

"%ISCC_EXE%" "packaging\windows\installer.iss"
if errorlevel 1 exit /b 1

echo Instalador listo en dist\installer
endlocal
