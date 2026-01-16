import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional

# Reutilizamos la lógica del CLI existente
from main import convert_file


class Pdf2WordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF → Word (DOCX)")
        self.geometry("560x280")
        self.resizable(False, False)

        # Variables de estado
        self.var_input = tk.StringVar()
        self.var_output = tk.StringVar()
        self.var_start = tk.StringVar()
        self.var_end = tk.StringVar()
        self.var_overwrite = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=True)

        # Entrada PDF
        ttk.Label(frm, text="Archivo PDF:").grid(row=0, column=0, sticky=tk.W, **pad)
        ent_in = ttk.Entry(frm, textvariable=self.var_input, width=52)
        ent_in.grid(row=0, column=1, sticky=tk.W, **pad)
        ttk.Button(frm, text="Examinar…", command=self.on_browse_pdf).grid(row=0, column=2, **pad)

        # Salida DOCX
        ttk.Label(frm, text="Salida DOCX (opcional):").grid(row=1, column=0, sticky=tk.W, **pad)
        ent_out = ttk.Entry(frm, textvariable=self.var_output, width=52)
        ent_out.grid(row=1, column=1, sticky=tk.W, **pad)
        ttk.Button(frm, text="Guardar como…", command=self.on_browse_docx).grid(row=1, column=2, **pad)

        # Rango de páginas
        ttk.Label(frm, text="Página inicio (1-basado):").grid(row=2, column=0, sticky=tk.W, **pad)
        ttk.Entry(frm, textvariable=self.var_start, width=10).grid(row=2, column=1, sticky=tk.W, **pad)
        ttk.Label(frm, text="Página fin (1-basado):").grid(row=3, column=0, sticky=tk.W, **pad)
        ttk.Entry(frm, textvariable=self.var_end, width=10).grid(row=3, column=1, sticky=tk.W, **pad)

        # Overwrite
        ttk.Checkbutton(frm, text="Sobrescribir si existe", variable=self.var_overwrite).grid(row=4, column=1, sticky=tk.W, **pad)

        # Barra de estado y acciones
        self.lbl_status = ttk.Label(frm, text="Listo", foreground="#444")
        self.lbl_status.grid(row=5, column=0, columnspan=3, sticky=tk.W, **pad)

        btn_convertir = ttk.Button(frm, text="Convertir", command=self.on_convert_click)
        btn_convertir.grid(row=6, column=2, sticky=tk.E, **pad)

        for i in range(3):
            frm.grid_columnconfigure(i, weight=1 if i == 1 else 0)

    def on_browse_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos", "*.*")],
        )
        if path:
            self.var_input.set(path)
            # Proponer salida por defecto
            stem = Path(path).with_suffix(".docx")
            if not self.var_output.get():
                self.var_output.set(str(stem))

    def on_browse_docx(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Guardar como DOCX",
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")],
        )
        if path:
            self.var_output.set(path)

    def on_convert_click(self) -> None:
        in_path = self.var_input.get().strip()
        out_path = self.var_output.get().strip()
        start_s = self.var_start.get().strip()
        end_s = self.var_end.get().strip()

        if not in_path:
            messagebox.showwarning("Falta archivo", "Selecciona un archivo PDF de entrada.")
            return

        try:
            start_i: Optional[int] = int(start_s) if start_s else None
            end_i: Optional[int] = int(end_s) if end_s else None
        except ValueError:
            messagebox.showerror("Rango inválido", "Las páginas de inicio/fin deben ser números enteros.")
            return

        input_pdf = Path(in_path)
        output_docx = Path(out_path) if out_path else input_pdf.with_suffix(".docx")
        overwrite = bool(self.var_overwrite.get())

        # Ejecutar en hilo aparte para no congelar la UI
        th = threading.Thread(
            target=self._convert_task,
            args=(input_pdf, output_docx, start_i, end_i, overwrite),
            daemon=True,
        )
        th.start()

    def _convert_task(self, input_pdf: Path, output_docx: Path, start_i: Optional[int], end_i: Optional[int], overwrite: bool) -> None:
        try:
            self._set_status("Convirtiendo…")
            convert_file(input_pdf, output_docx, start_i, end_i, overwrite)
        except Exception as e:
            self._set_status("Error en la conversión")
            messagebox.showerror("Error", str(e))
            return

        self._set_status("Conversión completada")
        messagebox.showinfo("Listo", f"Archivo creado:\n{output_docx}")

    def _set_status(self, text: str) -> None:
        self.lbl_status.config(text=text)
        self.lbl_status.update_idletasks()


if __name__ == "__main__":
    app = Pdf2WordApp()
    app.mainloop()
