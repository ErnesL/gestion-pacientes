# Generador de informes (Excel -> PPTX)

Este proyecto genera el **Plan de Alimentación** y el **Informe Antropométrico** en formato PPTX a partir del Excel del nutricionista, usando plantillas base con placeholders.

## Requisitos
- Python 3.10+
- Dependencias: ver `requirements.txt`

## Uso - Plan de Alimentación
1) Instala dependencias:

```bash
pip install -r requirements.txt
```

2) Genera el PPTX:

```bash
python src/generate_pptx.py "ruta/al/archivo.xlsx" \
  --template "src-material/Plan de Alimentación Base.pptx" \
  --output "output/Plan Alimentacion.pptx"
```

Alternativa rápida:

```bash
./scripts/run_pptx.sh
```

En Windows:

```bat
scripts\\run_pptx.bat
```

El PPTX se guarda en `output/`.

## Uso - Informe Antropométrico
```bash
.venv/bin/python src/generate_anthro_pptx.py "ruta/al/archivo.xlsx" \
  --template "src-material/Informe Antropométrico base.pptx" \
  --output "output/Informe Antropometrico.pptx"
```

Opcional para pruebas deterministas:

```bash
.venv/bin/python src/generate_anthro_pptx.py "ruta/al/archivo.xlsx" \
  --today "2026-03-02"
```

Alternativa rápida:

```bash
bash ./scripts/run_anthro_pptx.sh
```

En Windows:

```bat
scripts\\run_anthro_pptx.bat
```

## Plantilla - Informe Antropométrico
- La plantilla base debe vivir en `src-material/Informe Antropométrico base.pptx`.
- Los textos variables se reemplazan con placeholders `{{...}}`.
- La diapositiva 3 usa una tabla de `13 x 2`.
- La tabla de la diapositiva 3 se llena desde `RESUMEN ANTROPOMETRICO!D4:D16` y `F4:F16`.
- La columna `E` del resumen se ignora.
- La diapositiva 4 usa una tabla de `26 x 2`.
- La tabla de la diapositiva 4 se llena desde `RESUMEN ANTROPOMETRICO!E36:F61`.
- En la diapositiva 4, las filas `5`, `15` y `24` pueden estar mergeadas horizontalmente; el script solo escribe el placeholder que exista en cada celda.
- Placeholders esperados para tabla resumen: `{{R1C1}}` ... `{{R13C2}}`.
- Placeholders esperados para tabla medidas: `{{M1C1}}` ... `{{M26C2}}`.
- Por compatibilidad, en la diapositiva 4 tambien se aceptan placeholders `{{R...}}`, pero el formato recomendado es `{{M...}}`.

## Notas
- En el informe antropométrico, `Metas Nutricionales`, `Diagnóstico`, `Observaciones` y `Meta Final` quedan para edición manual.
- En el informe antropométrico, `Objetivo` se fija automáticamente en `PERDER GRASA`.
- El próximo control se calcula automáticamente como fecha de hoy + 6 semanas (42 días).
- Para exportar a PDF, abre el PPTX en PowerPoint y usa **Archivo → Exportar → Crear PDF**.
