@echo off
setlocal

cd /d "%~dp0.."

if not exist ".venv\Scripts\python.exe" (
  py -3 -m venv .venv
)

call ".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 exit /b 1

call ".venv\Scripts\python.exe" -m pip install -r requirements-windows.txt
if errorlevel 1 exit /b 1

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

call ".venv\Scripts\pyinstaller.exe" ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onedir ^
  --name GestionPacientes ^
  src\windows_gui.py
if errorlevel 1 exit /b 1

if not exist "dist\GestionPacientes\templates" mkdir "dist\GestionPacientes\templates"
copy /Y "src-material\Plan de Alimentación Base.pptx" "dist\GestionPacientes\templates\" >nul
if errorlevel 1 exit /b 1
copy /Y "src-material\Informe Antropométrico base.pptx" "dist\GestionPacientes\templates\" >nul
if errorlevel 1 exit /b 1

echo Build listo en dist\GestionPacientes
endlocal
