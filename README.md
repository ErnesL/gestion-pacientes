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
python src/generate_pptx.py "ruta/al/archivo.xlsx" \
  --template "templates/plan-de-alimentacion-base.pptx" \
  --output "output/Plan Alimentacion.pptx"
```

Genera el informe antropometrico:

```bash
python src/generate_anthro_pptx.py "ruta/al/archivo.xlsx" \
  --template "templates/informe-antropometrico-base.pptx" \
  --output "output/Informe Antropometrico.pptx"
```

Prueba rapida con scripts:

```bash
./scripts/run_pptx.sh
bash ./scripts/run_anthro_pptx.sh
```

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
