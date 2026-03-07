from __future__ import annotations

import queue
import threading
import traceback
from dataclasses import dataclass
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from app_support import GenerationResult, generate_all_documents


@dataclass
class WorkerMessage:
    kind: str
    text: str = ""
    result: GenerationResult | None = None


class GestionPacientesApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Gestion de Pacientes")
        self.root.geometry("760x520")
        self.root.minsize(700, 480)

        self.queue: queue.Queue[WorkerMessage] = queue.Queue()
        self.worker_running = False

        self.excel_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.status_var = tk.StringVar(
            value="Selecciona un Excel y una carpeta destino.")

        self._build_ui()
        self._refresh_generate_button()
        self.root.after(150, self._poll_queue)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

        title = ttk.Label(
            frame,
            text="Generacion de PPTX y PDF",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        subtitle = ttk.Label(
            frame,
            text="La aplicacion genera el plan de alimentacion y el informe antropometrico en PPTX y PDF.",
            wraplength=700,
        )
        subtitle.grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 18))

        ttk.Label(frame, text="Archivo Excel").grid(
            row=2, column=0, sticky="w", pady=(0, 8))
        excel_entry = ttk.Entry(frame, textvariable=self.excel_var)
        excel_entry.grid(row=2, column=1, sticky="ew",
                         pady=(0, 8), padx=(12, 12))
        ttk.Button(frame, text="Examinar...", command=self._choose_excel).grid(
            row=2, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Carpeta destino").grid(
            row=3, column=0, sticky="w", pady=(0, 12))
        output_entry = ttk.Entry(frame, textvariable=self.output_dir_var)
        output_entry.grid(row=3, column=1, sticky="ew",
                          pady=(0, 12), padx=(12, 12))
        ttk.Button(frame, text="Examinar...", command=self._choose_output_dir).grid(
            row=3, column=2, sticky="ew", pady=(0, 12))

        status_frame = ttk.LabelFrame(frame, text="Estado", padding=12)
        status_frame.grid(row=4, column=0, columnspan=3, sticky="nsew")
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(1, weight=1)

        ttk.Label(status_frame, textvariable=self.status_var).grid(
            row=0, column=0, sticky="w", pady=(0, 10))

        self.log_widget = scrolledtext.ScrolledText(
            status_frame,
            wrap="word",
            height=16,
            font=("Consolas", 10),
            state="disabled",
        )
        self.log_widget.grid(row=1, column=0, sticky="nsew")

        self.generate_button = ttk.Button(
            frame,
            text="Generar PPTXs",
            command=self._start_generation,
        )
        self.generate_button.grid(row=5, column=2, sticky="e", pady=(14, 0))

        self.excel_var.trace_add(
            "write", lambda *_: self._refresh_generate_button())
        self.output_dir_var.trace_add(
            "write", lambda *_: self._refresh_generate_button())

    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx"),
                       ("Todos los archivos", "*.*")],
        )
        if path:
            self.excel_var.set(path)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Selecciona la carpeta destino")
        if path:
            self.output_dir_var.set(path)

    def _refresh_generate_button(self) -> None:
        can_generate = (
            not self.worker_running
            and bool(self.excel_var.get().strip())
            and bool(self.output_dir_var.get().strip())
        )
        self.generate_button.state(
            ["!disabled"] if can_generate else ["disabled"])

    def _append_log(self, text: str) -> None:
        self.log_widget.configure(state="normal")
        self.log_widget.insert("end", text + "\n")
        self.log_widget.see("end")
        self.log_widget.configure(state="disabled")

    def _set_idle_status(self, text: str) -> None:
        self.worker_running = False
        self.status_var.set(text)
        self._refresh_generate_button()

    def _start_generation(self) -> None:
        excel_path = Path(self.excel_var.get().strip())
        output_dir = Path(self.output_dir_var.get().strip())

        self.worker_running = True
        self._refresh_generate_button()
        self.status_var.set("Iniciando generacion...")
        self._append_log("")
        self._append_log("=== Nueva ejecucion ===")
        self._append_log(f"Excel: {excel_path}")
        self._append_log(f"Destino: {output_dir}")

        worker = threading.Thread(
            target=self._run_generation_worker,
            args=(excel_path, output_dir),
            daemon=True,
        )
        worker.start()

    def _run_generation_worker(self, excel_path: Path, output_dir: Path) -> None:
        def log(text: str) -> None:
            self.queue.put(WorkerMessage(kind="log", text=text))

        try:
            result = generate_all_documents(
                excel_path=excel_path,
                output_dir=output_dir,
                log=log,
                today=date.today(),
            )
        except Exception as exc:
            self.queue.put(
                WorkerMessage(
                    kind="error",
                    text=f"{exc}\n\n{traceback.format_exc()}",
                )
            )
            return

        self.queue.put(WorkerMessage(kind="done", result=result))

    def _poll_queue(self) -> None:
        try:
            while True:
                message = self.queue.get_nowait()
                self._handle_worker_message(message)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    def _handle_worker_message(self, message: WorkerMessage) -> None:
        if message.kind == "log":
            self.status_var.set(message.text)
            self._append_log(message.text)
            return

        if message.kind == "error":
            self._append_log("ERROR")
            self._append_log(message.text)
            self._set_idle_status("Error durante la generacion.")
            messagebox.showerror("Generacion fallida",
                                 message.text.splitlines()[0])
            return

        if message.kind == "done" and message.result is not None:
            result = message.result
            self._render_result(result)

    def _render_result(self, result: GenerationResult) -> None:
        success_lines: list[str] = []
        warning_lines: list[str] = []

        for document in result.documents:
            if document.pptx_path is not None:
                success_lines.append(
                    f"{document.label} PPTX: {document.pptx_path}")
            if document.pdf_path is not None:
                success_lines.append(
                    f"{document.label} PDF: {document.pdf_path}")
            warning_lines.extend(document.errors)

        if success_lines:
            self._append_log("Archivos generados:")
            for line in success_lines:
                self._append_log(f"- {line}")

        if warning_lines:
            self._append_log("Advertencias:")
            for warning in warning_lines:
                self._append_log(f"- {warning}")
            self._set_idle_status("Completado con advertencias.")
            messagebox.showwarning(
                "Generacion completada con advertencias",
                "\n".join(warning_lines),
            )
            return

        self._set_idle_status("Completado correctamente.")
        messagebox.showinfo(
            "Generacion completada",
            "Se generaron los PPTX y PDF correctamente.",
        )


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = GestionPacientesApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
