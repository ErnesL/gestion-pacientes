from __future__ import annotations

import platform
from pathlib import Path


PDF_FORMAT_TYPE = 32


class PdfExportError(RuntimeError):
    pass


def ensure_windows_powerpoint_support() -> None:
    if platform.system() != "Windows":
        raise PdfExportError(
            "La exportacion a PDF solo esta disponible en Windows.")
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise PdfExportError(
            "No se encontro pywin32. Instala las dependencias de Windows."
        ) from exc


def export_pptx_to_pdf(pptx_path: Path | str, pdf_path: Path | str) -> Path:
    ensure_windows_powerpoint_support()

    import pythoncom
    import win32com.client

    pptx_path = Path(pptx_path)
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    app = None
    presentation = None
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = 1
        presentation = app.Presentations.Open(str(pptx_path), WithWindow=False)
        presentation.SaveAs(str(pdf_path), PDF_FORMAT_TYPE)
    except Exception as exc:
        raise PdfExportError(
            f"PowerPoint no pudo exportar a PDF: {exc}") from exc
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()

    return pdf_path
