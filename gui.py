import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional

from tools import pdf_to_docx, docx_to_pdf, compress_pdf, compress_docx_images


class Pdf2WordApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF → Word (DOCX)")
        self.geometry("760x520")
        self.resizable(False, False)

        # Tema más amable
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except Exception:
            style.theme_use("clam")
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TEntry", padding=2)

        # Variables de estado
        self.var_input = tk.StringVar()
        self.var_output = tk.StringVar()
        self.var_start = tk.StringVar()
        self.var_end = tk.StringVar()
        self.var_overwrite = tk.BooleanVar(value=False)

        # Para DOCX->PDF
        self.var_docx_in = tk.StringVar()
        self.var_pdf_out = tk.StringVar()
        self.var_docx_overwrite = tk.BooleanVar(value=False)

        # Para compresión
        self.var_pdf_comp_in = tk.StringVar()
        self.var_pdf_comp_out = tk.StringVar()
        self.var_docx_comp_in = tk.StringVar()
        self.var_docx_comp_out = tk.StringVar()
        self.var_quality = tk.IntVar(value=75)
        self.var_max_w = tk.IntVar(value=0)
        self.var_max_h = tk.IntVar(value=0)

        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        # --- Tab PDF->DOCX ---
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text="PDF → DOCX")
        ttk.Label(tab1, text="Archivo PDF:").grid(row=0, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab1, textvariable=self.var_input, width=56).grid(row=0, column=1, sticky=tk.W, **pad)
        ttk.Button(tab1, text="Examinar…", command=self.on_browse_pdf).grid(row=0, column=2, **pad)
        ttk.Label(tab1, text="Salida DOCX (opcional):").grid(row=1, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab1, textvariable=self.var_output, width=56).grid(row=1, column=1, sticky=tk.W, **pad)
        ttk.Button(tab1, text="Guardar como…", command=self.on_browse_docx).grid(row=1, column=2, **pad)
        ttk.Label(tab1, text="Página inicio (1-basado):").grid(row=2, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab1, textvariable=self.var_start, width=12).grid(row=2, column=1, sticky=tk.W, **pad)
        ttk.Label(tab1, text="Página fin (1-basado):").grid(row=3, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab1, textvariable=self.var_end, width=12).grid(row=3, column=1, sticky=tk.W, **pad)
        ttk.Checkbutton(tab1, text="Sobrescribir si existe", variable=self.var_overwrite).grid(row=4, column=1, sticky=tk.W, **pad)
        ttk.Button(tab1, text="Convertir (editable)", command=self.on_convert_pdf2docx).grid(row=5, column=2, sticky=tk.E, **pad)

        # Modo fidelidad exacta (raster)
        self.var_raster_on = tk.BooleanVar(value=False)
        ttk.Checkbutton(tab1, text="Fidelidad exacta (imagen)", variable=self.var_raster_on).grid(row=6, column=1, sticky=tk.W, **pad)
        ttk.Label(tab1, text="DPI:").grid(row=6, column=0, sticky=tk.W, **pad)
        self.var_raster_dpi = tk.IntVar(value=200)
        ttk.Entry(tab1, textvariable=self.var_raster_dpi, width=8).grid(row=6, column=1, sticky=tk.W, padx=120)
        ttk.Button(tab1, text="Convertir (imagen)", command=self.on_convert_pdf2docx_raster).grid(row=7, column=2, sticky=tk.E, **pad)

        # --- Tab DOCX->PDF ---
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text="DOCX → PDF")
        ttk.Label(tab2, text="Archivo DOCX:").grid(row=0, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab2, textvariable=self.var_docx_in, width=56).grid(row=0, column=1, sticky=tk.W, **pad)
        ttk.Button(tab2, text="Examinar…", command=self.on_browse_docx_in).grid(row=0, column=2, **pad)
        ttk.Label(tab2, text="Salida PDF (opcional):").grid(row=1, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab2, textvariable=self.var_pdf_out, width=56).grid(row=1, column=1, sticky=tk.W, **pad)
        ttk.Button(tab2, text="Guardar como…", command=self.on_browse_pdf_out).grid(row=1, column=2, **pad)
        ttk.Checkbutton(tab2, text="Sobrescribir si existe", variable=self.var_docx_overwrite).grid(row=2, column=1, sticky=tk.W, **pad)
        ttk.Button(tab2, text="Convertir", command=self.on_convert_docx2pdf).grid(row=3, column=2, sticky=tk.E, **pad)

        # --- Tab Compresión ---
        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text="Compresión")
        ttk.Label(tab3, text="PDF a optimizar:").grid(row=0, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab3, textvariable=self.var_pdf_comp_in, width=56).grid(row=0, column=1, sticky=tk.W, **pad)
        ttk.Button(tab3, text="Examinar…", command=self.on_browse_pdf_comp_in).grid(row=0, column=2, **pad)
        ttk.Label(tab3, text="PDF optimizado:").grid(row=1, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab3, textvariable=self.var_pdf_comp_out, width=56).grid(row=1, column=1, sticky=tk.W, **pad)
        ttk.Button(tab3, text="Guardar como…", command=self.on_browse_pdf_comp_out).grid(row=1, column=2, **pad)
        ttk.Button(tab3, text="Optimizar PDF", command=self.on_compress_pdf).grid(row=2, column=2, sticky=tk.E, **pad)

        ttk.Separator(tab3).grid(row=3, column=0, columnspan=3, sticky="ew", **pad)

        ttk.Label(tab3, text="DOCX a comprimir:").grid(row=4, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab3, textvariable=self.var_docx_comp_in, width=56).grid(row=4, column=1, sticky=tk.W, **pad)
        ttk.Button(tab3, text="Examinar…", command=self.on_browse_docx_comp_in).grid(row=4, column=2, **pad)
        ttk.Label(tab3, text="DOCX comprimido:").grid(row=5, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab3, textvariable=self.var_docx_comp_out, width=56).grid(row=5, column=1, sticky=tk.W, **pad)
        ttk.Button(tab3, text="Guardar como…", command=self.on_browse_docx_comp_out).grid(row=5, column=2, **pad)
        ttk.Label(tab3, text="Calidad JPEG (1-95):").grid(row=6, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab3, textvariable=self.var_quality, width=12).grid(row=6, column=1, sticky=tk.W, **pad)
        ttk.Label(tab3, text="Ancho máx:").grid(row=7, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab3, textvariable=self.var_max_w, width=12).grid(row=7, column=1, sticky=tk.W, **pad)
        ttk.Label(tab3, text="Alto máx:").grid(row=8, column=0, sticky=tk.W, **pad)
        ttk.Entry(tab3, textvariable=self.var_max_h, width=12).grid(row=8, column=1, sticky=tk.W, **pad)
        ttk.Button(tab3, text="Comprimir DOCX", command=self.on_compress_docx).grid(row=9, column=2, sticky=tk.E, **pad)

        # Barra de estado
        status_frame = ttk.Frame(self)
        status_frame.pack(fill=tk.X, padx=10, pady=6)
        self.progress = ttk.Progressbar(status_frame, length=200, mode="determinate")
        self.progress.pack(side=tk.RIGHT)
        self.lbl_status = ttk.Label(status_frame, text="Listo", foreground="#444")
        self.lbl_status.pack(side=tk.LEFT)

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

    def on_convert_pdf2docx(self) -> None:
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
            target=self._convert_pdf2docx_task,
            args=(input_pdf, output_docx, start_i, end_i, overwrite),
            daemon=True,
        )
        th.start()

    def _convert_pdf2docx_task(self, input_pdf: Path, output_docx: Path, start_i: Optional[int], end_i: Optional[int], overwrite: bool) -> None:
        try:
            self._set_status("Convirtiendo…")
            pdf_to_docx(input_pdf, output_docx, start_i, end_i, overwrite)
        except Exception as e:
            self._set_status("Error en la conversión")
            messagebox.showerror("Error", str(e))
            return

        self._set_status("Conversión completada")
        messagebox.showinfo("Listo", f"Archivo creado:\n{output_docx}")

    def on_convert_pdf2docx_raster(self) -> None:
        in_path = self.var_input.get().strip()
        out_path = self.var_output.get().strip()
        if not in_path:
            messagebox.showwarning("Falta archivo", "Selecciona un archivo PDF de entrada.")
            return
        input_pdf = Path(in_path)
        output_docx = Path(out_path) if out_path else input_pdf.with_suffix(".docx")
        dpi = int(self.var_raster_dpi.get()) if str(self.var_raster_dpi.get()).strip() else 200
        th = threading.Thread(target=self._convert_pdf2docx_raster_task, args=(input_pdf, output_docx, dpi), daemon=True)
        th.start()

    def _convert_pdf2docx_raster_task(self, input_pdf: Path, output_docx: Path, dpi: int) -> None:
        try:
            self._set_status("Convirtiendo (imagen)…")
            pdf_to_docx_raster(input_pdf, output_docx, dpi=dpi)
        except Exception as e:
            self._set_status("Error en conversión por imagen")
            messagebox.showerror("Error", str(e))
            return
        self._set_status("Conversión completada")
        messagebox.showinfo("Listo", f"Archivo creado:\n{output_docx}")

    # DOCX -> PDF
    def on_browse_docx_in(self) -> None:
        path = filedialog.askopenfilename(title="Seleccionar DOCX", filetypes=[("Documento Word", "*.docx")])
        if path:
            self.var_docx_in.set(path)
            if not self.var_pdf_out.get():
                self.var_pdf_out.set(str(Path(path).with_suffix('.pdf')))

    def on_browse_pdf_out(self) -> None:
        path = filedialog.asksaveasfilename(title="Guardar como PDF", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if path:
            self.var_pdf_out.set(path)

    def on_convert_docx2pdf(self) -> None:
        inp = self.var_docx_in.get().strip()
        out = self.var_pdf_out.get().strip()
        overwrite = bool(self.var_docx_overwrite.get())
        if not inp:
            messagebox.showwarning("Falta archivo", "Selecciona un archivo DOCX de entrada.")
            return
        input_docx = Path(inp)
        output_pdf = Path(out) if out else input_docx.with_suffix('.pdf')
        th = threading.Thread(target=self._convert_docx2pdf_task, args=(input_docx, output_pdf, overwrite), daemon=True)
        th.start()

    def _convert_docx2pdf_task(self, input_docx: Path, output_pdf: Path, overwrite: bool) -> None:
        try:
            self._set_status("Convirtiendo DOCX → PDF…")
            docx_to_pdf(input_docx, output_pdf, overwrite)
        except Exception as e:
            self._set_status("Error en DOCX→PDF")
            messagebox.showerror("Error", str(e))
            return
        self._set_status("Conversión completada")
        messagebox.showinfo("Listo", f"Archivo creado:\n{output_pdf}")

    # Compresión PDF
    def on_browse_pdf_comp_in(self) -> None:
        path = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("PDF", "*.pdf")])
        if path:
            self.var_pdf_comp_in.set(path)
            if not self.var_pdf_comp_out.get():
                self.var_pdf_comp_out.set(str(Path(path).with_name(Path(path).stem + '_optimized.pdf')))

    def on_browse_pdf_comp_out(self) -> None:
        path = filedialog.asksaveasfilename(title="Guardar PDF optimizado", defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if path:
            self.var_pdf_comp_out.set(path)

    def on_compress_pdf(self) -> None:
        inp = self.var_pdf_comp_in.get().strip()
        out = self.var_pdf_comp_out.get().strip()
        if not inp or not out:
            messagebox.showwarning("Faltan rutas", "Selecciona PDF de entrada y salida.")
            return
        th = threading.Thread(target=self._compress_pdf_task, args=(Path(inp), Path(out)), daemon=True)
        th.start()

    def _compress_pdf_task(self, inp: Path, out: Path) -> None:
        try:
            self._set_status("Optimizando PDF…")
            compress_pdf(inp, out)
        except Exception as e:
            self._set_status("Error al optimizar PDF")
            messagebox.showerror("Error", str(e))
            return
        self._set_status("PDF optimizado")
        messagebox.showinfo("Listo", f"PDF optimizado:\n{out}")

    # Compresión DOCX
    def on_browse_docx_comp_in(self) -> None:
        path = filedialog.askopenfilename(title="Seleccionar DOCX", filetypes=[("DOCX", "*.docx")])
        if path:
            self.var_docx_comp_in.set(path)
            if not self.var_docx_comp_out.get():
                self.var_docx_comp_out.set(str(Path(path).with_name(Path(path).stem + '_compressed.docx')))

    def on_browse_docx_comp_out(self) -> None:
        path = filedialog.asksaveasfilename(title="Guardar DOCX comprimido", defaultextension=".docx", filetypes=[("DOCX", "*.docx")])
        if path:
            self.var_docx_comp_out.set(path)

    def on_compress_docx(self) -> None:
        inp = self.var_docx_comp_in.get().strip()
        out = self.var_docx_comp_out.get().strip()
        q = int(self.var_quality.get()) if str(self.var_quality.get()).strip() else 75
        q = max(1, min(95, q))
        mw = int(self.var_max_w.get()) if str(self.var_max_w.get()).strip() else 0
        mh = int(self.var_max_h.get()) if str(self.var_max_h.get()).strip() else 0
        max_w = mw if mw > 0 else None
        max_h = mh if mh > 0 else None
        if not inp or not out:
            messagebox.showwarning("Faltan rutas", "Selecciona DOCX de entrada y salida.")
            return
        th = threading.Thread(target=self._compress_docx_task, args=(Path(inp), Path(out), q, max_w, max_h), daemon=True)
        th.start()

    def _compress_docx_task(self, inp: Path, out: Path, q: int, max_w: Optional[int], max_h: Optional[int]) -> None:
        try:
            self._set_status("Comprimiendo imágenes DOCX…")
            compress_docx_images(inp, out, quality=q, max_width=max_w, max_height=max_h)
        except Exception as e:
            self._set_status("Error al comprimir DOCX")
            messagebox.showerror("Error", str(e))
            return
        self._set_status("DOCX comprimido")
        messagebox.showinfo("Listo", f"DOCX comprimido:\n{out}")

    def _set_status(self, text: str) -> None:
        self.lbl_status.config(text=text)
        self.lbl_status.update_idletasks()
        self.progress.stop()
        self.progress.config(value=0)

        
        # --- Tab Lotes ---
        tab4 = ttk.Frame(notebook)
        notebook.add(tab4, text="Lotes")
        pad = {"padx": 10, "pady": 6}
        self.file_list = tk.Listbox(tab4, height=10, selectmode=tk.EXTENDED)
        self.file_list.grid(row=0, column=0, columnspan=3, sticky="nsew", **pad)
        ttk.Button(tab4, text="Agregar PDFs", command=self.on_add_pdfs).grid(row=1, column=0, **pad)
        ttk.Button(tab4, text="Agregar DOCXs", command=self.on_add_docxs).grid(row=1, column=1, **pad)
        ttk.Button(tab4, text="Quitar seleccionados", command=self.on_remove_selected).grid(row=1, column=2, **pad)
        ttk.Label(tab4, text="Carpeta salida:").grid(row=2, column=0, sticky=tk.W, **pad)
        self.var_outdir_batch = tk.StringVar()
        ttk.Entry(tab4, textvariable=self.var_outdir_batch, width=56).grid(row=2, column=1, sticky=tk.W, **pad)
        ttk.Button(tab4, text="Elegir…", command=self.on_choose_outdir_batch).grid(row=2, column=2, **pad)
        self.var_batch_pdf2docx = tk.BooleanVar(value=True)
        self.var_batch_raster = tk.BooleanVar(value=False)
        self.var_batch_docx2pdf = tk.BooleanVar(value=False)
        ttk.Checkbutton(tab4, text="PDF → DOCX", variable=self.var_batch_pdf2docx).grid(row=3, column=0, sticky=tk.W, **pad)
        ttk.Checkbutton(tab4, text="Fidelidad exacta (imagen)", variable=self.var_batch_raster).grid(row=3, column=1, sticky=tk.W, **pad)
        ttk.Checkbutton(tab4, text="DOCX → PDF", variable=self.var_batch_docx2pdf).grid(row=3, column=2, sticky=tk.W, **pad)
        ttk.Label(tab4, text="DPI:").grid(row=4, column=0, sticky=tk.W, **pad)
        self.var_batch_dpi = tk.IntVar(value=200)
        ttk.Entry(tab4, textvariable=self.var_batch_dpi, width=10).grid(row=4, column=1, sticky=tk.W, **pad)
        ttk.Button(tab4, text="Convertir lote", command=self.on_run_batch).grid(row=5, column=2, sticky=tk.E, **pad)
        for i in range(3):
            tab4.grid_columnconfigure(i, weight=1 if i == 1 else 0)

    # --- Handlers de pestaña Lotes ---
    def on_add_pdfs(self):
        paths = filedialog.askopenfilenames(title="Seleccionar PDFs", filetypes=[("PDF", "*.pdf")])
        for p in paths:
            self.file_list.insert(tk.END, p)

    def on_add_docxs(self):
        paths = filedialog.askopenfilenames(title="Seleccionar DOCXs", filetypes=[("DOCX", "*.docx")])
        for p in paths:
            self.file_list.insert(tk.END, p)

    def on_remove_selected(self):
        sel = list(self.file_list.curselection())
        for idx in reversed(sel):
            self.file_list.delete(idx)

    def on_choose_outdir_batch(self):
        path = filedialog.askdirectory(title="Elegir carpeta de salida")
        if path:
            self.var_outdir_batch.set(path)

    def on_run_batch(self):
        items = [Path(self.file_list.get(i)) for i in range(self.file_list.size())]
        outdir = Path(self.var_outdir_batch.get()) if self.var_outdir_batch.get().strip() else Path.cwd()
        do_pdf2docx = bool(self.var_batch_pdf2docx.get())
        do_raster = bool(self.var_batch_raster.get())
        do_docx2pdf = bool(self.var_batch_docx2pdf.get())
        dpi = int(self.var_batch_dpi.get()) if str(self.var_batch_dpi.get()).strip() else 200
        th = threading.Thread(target=self._run_batch_task, args=(items, outdir, do_pdf2docx, do_raster, do_docx2pdf, dpi), daemon=True)
        th.start()

    def _run_batch_task(self, items: list[Path], outdir: Path, do_pdf2docx: bool, do_raster: bool, do_docx2pdf: bool, dpi: int):
        try:
            self._set_status("Procesando lote…")
            self.progress.config(mode="determinate", maximum=max(1, len(items)))
            done = 0
            pdfs = [p for p in items if p.suffix.lower() == ".pdf"]
            docxs = [p for p in items if p.suffix.lower() == ".docx"]
            if do_pdf2docx:
                mode = "raster" if do_raster else "editable"
                for p in pdfs:
                    tgt = outdir / (p.stem + ".docx")
                    if mode == "raster":
                        pdf_to_docx_raster(p, tgt, dpi=dpi)
                    else:
                        pdf_to_docx(p, tgt, None, None, True)
                    done += 1
                    self.progress.config(value=done)
            if do_docx2pdf:
                for d in docxs:
                    tgt = outdir / (d.stem + ".pdf")
                    docx_to_pdf(d, tgt, True)
                    done += 1
                    self.progress.config(value=done)
        except Exception as e:
            self._set_status("Error en lote")
            messagebox.showerror("Error", str(e))
            return
        self._set_status("Lote completado")
        messagebox.showinfo("Listo", f"Lote completado en:\n{outdir}")


if __name__ == "__main__":
    app = Pdf2WordApp()
    app.mainloop()
