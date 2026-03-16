# Gestion de Pacientes

Genera 2 documentos desde un Excel del nutricionista:
- `Plan de Alimentacion`
- `Informe Antropometrico`

En desarrollo puede ejecutarse por CLI. En Windows tiene una GUI que genera 4 archivos:
- `Plan Alimentacion - {Paciente}.pptx`
- `Plan Alimentacion - {Paciente}.pdf`
- `Informe Antropometrico - {Paciente}.pptx`
- `Informe Antropometrico - {Paciente}.pdf`

## Requisitos
- Python 3.10+
- Para PDF en Windows: Microsoft PowerPoint instalado

## Templates
Las plantillas versionadas viven en `templates/`:
- `templates/plan-de-alimentacion-base.pptx`
- `templates/informe-antropometrico-base.pptx`

## Desarrollo
Instala dependencias:

```bash
pip install -r requirements.txt
```

Genera el plan:

```bash
python src/generate_pptx.py "src/ayuda.xlsx" \
  --template "templates/plan-de-alimentacion-base.pptx" \
  --output "output/Plan Alimentacion.pptx"
```

Genera el informe antropometrico:

```bash
python src/generate_anthro_pptx.py "ayuda.xlsx" \
  --template "templates/informe-antropometrico-base.pptx" \
  --output "output/Informe Antropometrico.pptx"
```

Prueba rapida con scripts:

```bash
./scripts/run_pptx.sh
bash ./scripts/run_anthro_pptx.sh
```

## Hojas obligatorias
El parser usa `HISTORIA` para los datos del paciente y exige nuestras hojas para plan y antropometría. Si estas hojas no existen o están incompletas, la generación falla:

- `PLAN_ALIMENTACION_TEMPLATE`
  - Columnas requeridas en la fila 1: `COMIDA`, `LACTEOS`, `VEGETALES`, `FRUTAS`, `ALMIDONES`, `PROTEINAS`, `GRASAS`
  - Comidas permitidas: `PRE`, `DES`, `MAM`, `ALM`, `MTP`, `CEN`
- `ANTROPOMETRIA_TEMPLATE`
  - Columnas requeridas en la fila 1: `SECCION`, `ETIQUETA`, `VALOR`
  - Secciones permitidas: `RESUMEN`, `MEDIDAS`

Hay un ejemplo listo para copiar en:

```bash
examples/ejemplo-template-codex.xlsx
```

Ese archivo es una referencia para copiar las hojas dentro del Excel real del cliente.

No hay compatibilidad con layouts viejos del cliente en `REQUERIMIENTOS` ni `RESUMEN ANTROPOMETRICO`. Si quieren generar documentos, deben llenar nuestras hojas.

## Windows
Instala dependencias de build:

```bat
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements-windows.txt
```

Ejecuta la GUI en desarrollo:

```bat
.venv\Scripts\python.exe src\windows_gui.py
```

Genera la app:

```bat
scripts\build_windows_app.bat
```

Genera el instalador:

```bat
scripts\build_windows_installer.bat
```

## Notas
- La GUI usa las plantillas en `templates/`.
- Si el PDF falla, los PPTX se conservan y se reporta advertencia.
- La exportacion a PDF solo esta soportada en Windows.
