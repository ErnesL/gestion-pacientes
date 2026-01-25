# Generador de informes (Excel → PDF)

Este proyecto genera 2 PDFs a partir del Excel del nutricionista:
- Plan de Alimentación
- Informe Antropométrico

## Requisitos
- Python 3.10+
- Dependencias: ver `requirements.txt`
- WeasyPrint requiere dependencias del sistema (Windows) para renderizar PDF.

## Uso
1) Instala dependencias:

```bash
pip install -r requirements.txt
```

2) Genera los PDFs:

```bash
python src/generate_reports.py "ruta/al/archivo.xlsx"
```

Los PDFs se guardan en `output/`.

## Notas
- Los campos de Objetivo/Metas/Observaciones quedan en blanco para que el nutricionista los complete manualmente.
- El próximo control se calcula automáticamente como fecha de hoy + 6 semanas.
