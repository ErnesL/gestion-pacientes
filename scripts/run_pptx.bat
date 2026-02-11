@echo off
setlocal

set "EXCEL_PATH=%~1"
if "%EXCEL_PATH%"=="" set "EXCEL_PATH=src-material\\test.xlsx"

set "TEMPLATE_PATH=%~2"
if "%TEMPLATE_PATH%"=="" set "TEMPLATE_PATH=src-material\\Plan de Alimentación Base.pptx"

set "OUTPUT_PATH=%~3"
if "%OUTPUT_PATH%"=="" set "OUTPUT_PATH=output\\Plan Alimentacion - prueba.pptx"

python src\\generate_pptx.py "%EXCEL_PATH%" --template "%TEMPLATE_PATH%" --output "%OUTPUT_PATH%"
endlocal
