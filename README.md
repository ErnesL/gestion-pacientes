# Generador de informes (Excel → PPTX)

Este proyecto genera el **Plan de Alimentación** en formato PPTX a partir del Excel del nutricionista, usando una plantilla base con placeholders.

## Requisitos
- Python 3.10+
- Dependencias: ver `requirements.txt`

## Uso
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

## Notas
- Los campos de Objetivo/Metas/Observaciones quedan en blanco para que el nutricionista los complete manualmente.
- El próximo control se calcula automáticamente como fecha de hoy + 6 semanas.
- Para exportar a PDF, abre el PPTX en PowerPoint y usa **Archivo → Exportar → Crear PDF**.
