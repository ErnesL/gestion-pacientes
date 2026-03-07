@echo off
setlocal

set "EXCEL_PATH=%~1"
if "%EXCEL_PATH%"=="" set "EXCEL_PATH=src-material\\test.xlsx"

set "TEMPLATE_PATH=%~2"
if "%TEMPLATE_PATH%"=="" set "TEMPLATE_PATH=src-material\\Informe Antropométrico base.pptx"

set "OUTPUT_PATH=%~3"
if "%OUTPUT_PATH%"=="" set "OUTPUT_PATH=output\\Informe Antropometrico - output.pptx"

set "TODAY=%~4"
if "%TODAY%"=="" (
  python src\\generate_anthro_pptx.py "%EXCEL_PATH%" --template "%TEMPLATE_PATH%" --output "%OUTPUT_PATH%"
) else (
  python src\\generate_anthro_pptx.py "%EXCEL_PATH%" --template "%TEMPLATE_PATH%" --output "%OUTPUT_PATH%" --today "%TODAY%"
)
endlocal
