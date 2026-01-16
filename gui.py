import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import Optional

from tools import (
    pdf_to_docx, docx_to_pdf, compress_pdf, compress_docx_images,
    pdf_to_docx_raster, ocr_pdf_to_docx
)

# Configurar apariencia
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class Pdf2WordApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF Converter Pro")
        self.geometry("900x650")
        self.minsize(800, 600)

        # Variables de estado - PDF->DOCX
        self.var_input = ctk.StringVar()
        self.var_output = ctk.StringVar()
        self.var_start = ctk.StringVar()
        self.var_end = ctk.StringVar()
        self.var_overwrite = ctk.BooleanVar(value=False)
        self.var_raster_on = ctk.BooleanVar(value=False)
        self.var_raster_dpi = ctk.IntVar(value=200)
        self.var_ocr_lang = ctk.StringVar(value="spa")
        self.var_ocr_dpi = ctk.IntVar(value=300)

        # Variables - DOCX->PDF
        self.var_docx_in = ctk.StringVar()
        self.var_pdf_out = ctk.StringVar()
        self.var_docx_overwrite = ctk.BooleanVar(value=False)

        # Variables - Compresion
        self.var_pdf_comp_in = ctk.StringVar()
        self.var_pdf_comp_out = ctk.StringVar()
        self.var_docx_comp_in = ctk.StringVar()
        self.var_docx_comp_out = ctk.StringVar()
        self.var_quality = ctk.IntVar(value=75)
        self.var_max_w = ctk.IntVar(value=0)
        self.var_max_h = ctk.IntVar(value=0)

        # Variables - Lotes
        self.var_outdir_batch = ctk.StringVar()
        self.var_batch_pdf2docx = ctk.BooleanVar(value=True)
        self.var_batch_raster = ctk.BooleanVar(value=False)
        self.var_batch_docx2pdf = ctk.BooleanVar(value=False)
        self.var_batch_overwrite = ctk.BooleanVar(value=True)
        self.var_batch_dpi = ctk.IntVar(value=200)

        self._build_ui()

    def _build_ui(self) -> None:
        # Header
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(20, 10))

        title_label = ctk.CTkLabel(
            header_frame,
            text="PDF Converter Pro",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title_label.pack(side="left")

        # Theme toggle
        self.theme_switch = ctk.CTkSwitch(
            header_frame,
            text="Modo Oscuro",
            command=self._toggle_theme,
            onvalue="dark",
            offvalue="light"
        )
        self.theme_switch.select()
        self.theme_switch.pack(side="right", padx=10)

        # Tabview principal
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=10)

        # Crear tabs
        tab1 = self.tabview.add("PDF → DOCX")
        tab2 = self.tabview.add("DOCX → PDF")
        tab3 = self.tabview.add("Compresion")
        tab4 = self.tabview.add("Lotes")

        self._build_pdf2docx_tab(tab1)
        self._build_docx2pdf_tab(tab2)
        self._build_compression_tab(tab3)
        self._build_batch_tab(tab4)

        # Status bar
        self._build_status_bar()

    def _build_pdf2docx_tab(self, parent) -> None:
        # Frame principal con scroll
        main_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Seccion: Archivo de entrada
        self._create_section_label(main_frame, "Archivo de Entrada")

        input_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        input_frame.pack(fill="x", pady=(0, 15))
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="PDF:", width=100, anchor="w").grid(row=0, column=0, padx=15, pady=12)
        ctk.CTkEntry(input_frame, textvariable=self.var_input, placeholder_text="Selecciona un archivo PDF...").grid(row=0, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(input_frame, text="Examinar", width=100, command=self.on_browse_pdf).grid(row=0, column=2, padx=15, pady=12)

        ctk.CTkLabel(input_frame, text="Salida DOCX:", width=100, anchor="w").grid(row=1, column=0, padx=15, pady=12)
        ctk.CTkEntry(input_frame, textvariable=self.var_output, placeholder_text="Opcional - se genera automaticamente").grid(row=1, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(input_frame, text="Guardar como", width=100, command=self.on_browse_docx).grid(row=1, column=2, padx=15, pady=12)

        # Seccion: Opciones de paginas
        self._create_section_label(main_frame, "Rango de Paginas (opcional)")

        pages_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        pages_frame.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(pages_frame, text="Pagina inicio:", width=100).pack(side="left", padx=15, pady=12)
        ctk.CTkEntry(pages_frame, textvariable=self.var_start, width=80, placeholder_text="1").pack(side="left", padx=5, pady=12)
        ctk.CTkLabel(pages_frame, text="Pagina fin:", width=100).pack(side="left", padx=15, pady=12)
        ctk.CTkEntry(pages_frame, textvariable=self.var_end, width=80, placeholder_text="Ultima").pack(side="left", padx=5, pady=12)
        ctk.CTkCheckBox(pages_frame, text="Sobrescribir si existe", variable=self.var_overwrite).pack(side="right", padx=15, pady=12)

        # Seccion: Modo de conversion
        self._create_section_label(main_frame, "Modo de Conversion")

        mode_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        mode_frame.pack(fill="x", pady=(0, 15))
        mode_frame.grid_columnconfigure((0, 1, 2), weight=1)

        # Conversion editable
        edit_card = self._create_card(mode_frame, "Editable", "Texto editable, puede perder formato")
        edit_card.grid(row=0, column=0, padx=10, pady=15, sticky="nsew")
        ctk.CTkButton(edit_card, text="Convertir", fg_color="#2196F3", hover_color="#1976D2", command=self.on_convert_pdf2docx).pack(pady=(10, 15), padx=15, fill="x")

        # Conversion imagen (fidelidad exacta)
        raster_card = self._create_card(mode_frame, "Imagen", "Fidelidad exacta, no editable")
        raster_card.grid(row=0, column=1, padx=10, pady=15, sticky="nsew")
        dpi_frame = ctk.CTkFrame(raster_card, fg_color="transparent")
        dpi_frame.pack(pady=5, padx=15, fill="x")
        ctk.CTkLabel(dpi_frame, text="DPI:").pack(side="left")
        ctk.CTkEntry(dpi_frame, textvariable=self.var_raster_dpi, width=60).pack(side="left", padx=5)
        ctk.CTkButton(raster_card, text="Convertir", fg_color="#4CAF50", hover_color="#388E3C", command=self.on_convert_pdf2docx_raster).pack(pady=(5, 15), padx=15, fill="x")

        # Conversion OCR
        ocr_card = self._create_card(mode_frame, "OCR", "Extrae texto de imagenes/escaneos")
        ocr_card.grid(row=0, column=2, padx=10, pady=15, sticky="nsew")
        lang_frame = ctk.CTkFrame(ocr_card, fg_color="transparent")
        lang_frame.pack(pady=5, padx=15, fill="x")
        ctk.CTkLabel(lang_frame, text="Idioma:").pack(side="left")
        ctk.CTkEntry(lang_frame, textvariable=self.var_ocr_lang, width=80, placeholder_text="spa").pack(side="left", padx=5)
        dpi_ocr_frame = ctk.CTkFrame(ocr_card, fg_color="transparent")
        dpi_ocr_frame.pack(pady=5, padx=15, fill="x")
        ctk.CTkLabel(dpi_ocr_frame, text="DPI:").pack(side="left")
        ctk.CTkEntry(dpi_ocr_frame, textvariable=self.var_ocr_dpi, width=60).pack(side="left", padx=5)
        ctk.CTkButton(ocr_card, text="Convertir", fg_color="#FF9800", hover_color="#F57C00", command=self.on_convert_pdf2docx_ocr).pack(pady=(5, 15), padx=15, fill="x")

    def _build_docx2pdf_tab(self, parent) -> None:
        main_frame = ctk.CTkFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self._create_section_label(main_frame, "Convertir Word a PDF")

        input_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        input_frame.pack(fill="x", pady=(0, 15))
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="DOCX:", width=100, anchor="w").grid(row=0, column=0, padx=15, pady=12)
        ctk.CTkEntry(input_frame, textvariable=self.var_docx_in, placeholder_text="Selecciona un archivo Word...").grid(row=0, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(input_frame, text="Examinar", width=100, command=self.on_browse_docx_in).grid(row=0, column=2, padx=15, pady=12)

        ctk.CTkLabel(input_frame, text="Salida PDF:", width=100, anchor="w").grid(row=1, column=0, padx=15, pady=12)
        ctk.CTkEntry(input_frame, textvariable=self.var_pdf_out, placeholder_text="Opcional - se genera automaticamente").grid(row=1, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(input_frame, text="Guardar como", width=100, command=self.on_browse_pdf_out).grid(row=1, column=2, padx=15, pady=12)

        options_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        options_frame.pack(fill="x", pady=(0, 15))
        ctk.CTkCheckBox(options_frame, text="Sobrescribir si existe", variable=self.var_docx_overwrite).pack(side="left", padx=15, pady=12)
        ctk.CTkButton(options_frame, text="Convertir a PDF", width=150, fg_color="#E91E63", hover_color="#C2185B", command=self.on_convert_docx2pdf).pack(side="right", padx=15, pady=12)

        # Info
        info_frame = ctk.CTkFrame(main_frame, fg_color=("gray85", "gray20"))
        info_frame.pack(fill="x", pady=20)
        ctk.CTkLabel(info_frame, text="Nota: Requiere Microsoft Word instalado para una conversion precisa.", text_color=("gray50", "gray60")).pack(pady=15)

    def _build_compression_tab(self, parent) -> None:
        main_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Compresion PDF
        self._create_section_label(main_frame, "Optimizar PDF")

        pdf_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        pdf_frame.pack(fill="x", pady=(0, 20))
        pdf_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(pdf_frame, text="PDF entrada:", width=120, anchor="w").grid(row=0, column=0, padx=15, pady=12)
        ctk.CTkEntry(pdf_frame, textvariable=self.var_pdf_comp_in, placeholder_text="PDF a optimizar...").grid(row=0, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(pdf_frame, text="Examinar", width=100, command=self.on_browse_pdf_comp_in).grid(row=0, column=2, padx=15, pady=12)

        ctk.CTkLabel(pdf_frame, text="PDF salida:", width=120, anchor="w").grid(row=1, column=0, padx=15, pady=12)
        ctk.CTkEntry(pdf_frame, textvariable=self.var_pdf_comp_out, placeholder_text="PDF optimizado...").grid(row=1, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(pdf_frame, text="Guardar como", width=100, command=self.on_browse_pdf_comp_out).grid(row=1, column=2, padx=15, pady=12)

        ctk.CTkButton(pdf_frame, text="Optimizar PDF", width=150, fg_color="#9C27B0", hover_color="#7B1FA2", command=self.on_compress_pdf).grid(row=2, column=2, padx=15, pady=15)

        # Compresion DOCX
        self._create_section_label(main_frame, "Comprimir Imagenes en DOCX")

        docx_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        docx_frame.pack(fill="x", pady=(0, 15))
        docx_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(docx_frame, text="DOCX entrada:", width=120, anchor="w").grid(row=0, column=0, padx=15, pady=12)
        ctk.CTkEntry(docx_frame, textvariable=self.var_docx_comp_in, placeholder_text="DOCX a comprimir...").grid(row=0, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(docx_frame, text="Examinar", width=100, command=self.on_browse_docx_comp_in).grid(row=0, column=2, padx=15, pady=12)

        ctk.CTkLabel(docx_frame, text="DOCX salida:", width=120, anchor="w").grid(row=1, column=0, padx=15, pady=12)
        ctk.CTkEntry(docx_frame, textvariable=self.var_docx_comp_out, placeholder_text="DOCX comprimido...").grid(row=1, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(docx_frame, text="Guardar como", width=100, command=self.on_browse_docx_comp_out).grid(row=1, column=2, padx=15, pady=12)

        # Opciones de compresion
        options_frame = ctk.CTkFrame(docx_frame, fg_color="transparent")
        options_frame.grid(row=2, column=0, columnspan=3, padx=15, pady=10, sticky="ew")

        ctk.CTkLabel(options_frame, text="Calidad JPEG (1-95):").pack(side="left", padx=(0, 5))
        ctk.CTkEntry(options_frame, textvariable=self.var_quality, width=60).pack(side="left", padx=5)

        ctk.CTkLabel(options_frame, text="Ancho max:").pack(side="left", padx=(20, 5))
        ctk.CTkEntry(options_frame, textvariable=self.var_max_w, width=60, placeholder_text="0=auto").pack(side="left", padx=5)

        ctk.CTkLabel(options_frame, text="Alto max:").pack(side="left", padx=(20, 5))
        ctk.CTkEntry(options_frame, textvariable=self.var_max_h, width=60, placeholder_text="0=auto").pack(side="left", padx=5)

        ctk.CTkButton(docx_frame, text="Comprimir DOCX", width=150, fg_color="#009688", hover_color="#00796B", command=self.on_compress_docx).grid(row=3, column=2, padx=15, pady=15)

    def _build_batch_tab(self, parent) -> None:
        main_frame = ctk.CTkFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self._create_section_label(main_frame, "Conversion por Lotes")

        # Lista de archivos
        list_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        list_frame.pack(fill="both", expand=True, pady=(0, 15))

        # Textbox como lista (CTkTextbox para mostrar archivos)
        self.file_listbox = ctk.CTkTextbox(list_frame, height=200)
        self.file_listbox.pack(fill="both", expand=True, padx=15, pady=15)
        self.file_listbox.configure(state="disabled")
        self.batch_files: list[Path] = []

        # Botones de archivo
        btn_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 15))
        ctk.CTkButton(btn_frame, text="Agregar PDFs", width=120, command=self.on_add_pdfs).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Agregar DOCXs", width=120, command=self.on_add_docxs).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Limpiar Lista", width=120, fg_color="#F44336", hover_color="#D32F2F", command=self.on_clear_list).pack(side="left", padx=5)

        # Carpeta de salida
        output_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        output_frame.pack(fill="x", pady=(0, 15))
        output_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(output_frame, text="Carpeta salida:", width=120, anchor="w").grid(row=0, column=0, padx=15, pady=12)
        ctk.CTkEntry(output_frame, textvariable=self.var_outdir_batch, placeholder_text="Carpeta donde guardar los archivos...").grid(row=0, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(output_frame, text="Elegir", width=100, command=self.on_choose_outdir_batch).grid(row=0, column=2, padx=15, pady=12)

        # Opciones
        options_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        options_frame.pack(fill="x", pady=(0, 15))

        checks_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        checks_frame.pack(fill="x", padx=15, pady=12)

        ctk.CTkCheckBox(checks_frame, text="PDF → DOCX", variable=self.var_batch_pdf2docx).pack(side="left", padx=10)
        ctk.CTkCheckBox(checks_frame, text="Modo Imagen", variable=self.var_batch_raster).pack(side="left", padx=10)
        ctk.CTkCheckBox(checks_frame, text="DOCX → PDF", variable=self.var_batch_docx2pdf).pack(side="left", padx=10)
        ctk.CTkCheckBox(checks_frame, text="Sobrescribir", variable=self.var_batch_overwrite).pack(side="left", padx=10)

        dpi_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        dpi_frame.pack(fill="x", padx=15, pady=(0, 12))
        ctk.CTkLabel(dpi_frame, text="DPI (modo imagen):").pack(side="left")
        ctk.CTkEntry(dpi_frame, textvariable=self.var_batch_dpi, width=60).pack(side="left", padx=10)

        ctk.CTkButton(options_frame, text="Iniciar Conversion", width=180, height=40, fg_color="#4CAF50", hover_color="#388E3C", font=ctk.CTkFont(size=14, weight="bold"), command=self.on_run_batch).pack(side="right", padx=15, pady=12)

    def _build_status_bar(self) -> None:
        status_frame = ctk.CTkFrame(self, height=50, fg_color=("gray85", "gray20"))
        status_frame.pack(fill="x", padx=20, pady=(0, 20))
        status_frame.pack_propagate(False)

        self.lbl_status = ctk.CTkLabel(status_frame, text="Listo", font=ctk.CTkFont(size=12))
        self.lbl_status.pack(side="left", padx=20, pady=10)

        self.progress = ctk.CTkProgressBar(status_frame, width=200)
        self.progress.pack(side="right", padx=20, pady=10)
        self.progress.set(0)

    def _create_section_label(self, parent, text: str) -> None:
        ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(10, 5))

    def _create_card(self, parent, title: str, description: str) -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=("gray85", "gray20"), corner_radius=10)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(15, 5))
        ctk.CTkLabel(card, text=description, font=ctk.CTkFont(size=11), text_color=("gray50", "gray60")).pack(pady=(0, 5))
        return card

    def _toggle_theme(self) -> None:
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            ctk.set_appearance_mode("light")
            self.theme_switch.deselect()
        else:
            ctk.set_appearance_mode("dark")
            self.theme_switch.select()

    # --- File browsers ---
    def on_browse_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar PDF",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos", "*.*")],
        )
        if path:
            self.var_input.set(path)
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

    # --- Conversions ---
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
            messagebox.showerror("Rango invalido", "Las paginas de inicio/fin deben ser numeros enteros.")
            return

        input_pdf = Path(in_path)
        output_docx = Path(out_path) if out_path else input_pdf.with_suffix(".docx")
        overwrite = bool(self.var_overwrite.get())

        th = threading.Thread(
            target=self._convert_pdf2docx_task,
            args=(input_pdf, output_docx, start_i, end_i, overwrite),
            daemon=True,
        )
        th.start()

    def _convert_pdf2docx_task(self, input_pdf: Path, output_docx: Path, start_i: Optional[int], end_i: Optional[int], overwrite: bool) -> None:
        try:
            self._set_status("Convirtiendo...", indeterminate=True)
            pdf_to_docx(input_pdf, output_docx, start_i, end_i, overwrite)
        except Exception as e:
            self._set_status("Error en la conversion")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return

        self._set_status("Conversion completada")
        self.after(0, lambda: messagebox.showinfo("Listo", f"Archivo creado:\n{output_docx}"))

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
            self._set_status("Convirtiendo (imagen)...", indeterminate=True)
            pdf_to_docx_raster(input_pdf, output_docx, dpi=dpi)
        except Exception as e:
            self._set_status("Error en conversion por imagen")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return
        self._set_status("Conversion completada")
        self.after(0, lambda: messagebox.showinfo("Listo", f"Archivo creado:\n{output_docx}"))

    def on_convert_pdf2docx_ocr(self) -> None:
        in_path = self.var_input.get().strip()
        out_path = self.var_output.get().strip()
        if not in_path:
            messagebox.showwarning("Falta archivo", "Selecciona un archivo PDF de entrada.")
            return
        input_pdf = Path(in_path)
        output_docx = Path(out_path) if out_path else input_pdf.with_suffix(".docx")
        dpi = int(self.var_ocr_dpi.get()) if str(self.var_ocr_dpi.get()).strip() else 300
        lang = self.var_ocr_lang.get().strip() or "spa"
        th = threading.Thread(target=self._convert_pdf2docx_ocr_task, args=(input_pdf, output_docx, dpi, lang), daemon=True)
        th.start()

    def _convert_pdf2docx_ocr_task(self, input_pdf: Path, output_docx: Path, dpi: int, lang: str) -> None:
        try:
            self._set_status("Convirtiendo (OCR)...", indeterminate=True)
            ocr_pdf_to_docx(input_pdf, output_docx, dpi=dpi, lang=lang)
        except Exception as e:
            self._set_status("Error en OCR")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return
        self._set_status("Conversion completada")
        self.after(0, lambda: messagebox.showinfo("Listo", f"Archivo creado:\n{output_docx}"))

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
            self._set_status("Convirtiendo DOCX a PDF...", indeterminate=True)
            docx_to_pdf(input_docx, output_pdf, overwrite)
        except Exception as e:
            self._set_status("Error en DOCX a PDF")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return
        self._set_status("Conversion completada")
        self.after(0, lambda: messagebox.showinfo("Listo", f"Archivo creado:\n{output_pdf}"))

    # --- Compression ---
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
            self._set_status("Optimizando PDF...", indeterminate=True)
            compress_pdf(inp, out)
        except Exception as e:
            self._set_status("Error al optimizar PDF")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return
        self._set_status("PDF optimizado")
        self.after(0, lambda: messagebox.showinfo("Listo", f"PDF optimizado:\n{out}"))

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
            self._set_status("Comprimiendo imagenes DOCX...", indeterminate=True)
            compress_docx_images(inp, out, quality=q, max_width=max_w, max_height=max_h)
        except Exception as e:
            self._set_status("Error al comprimir DOCX")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return
        self._set_status("DOCX comprimido")
        self.after(0, lambda: messagebox.showinfo("Listo", f"DOCX comprimido:\n{out}"))

    # --- Batch ---
    def on_add_pdfs(self) -> None:
        paths = filedialog.askopenfilenames(title="Seleccionar PDFs", filetypes=[("PDF", "*.pdf")])
        for p in paths:
            self.batch_files.append(Path(p))
        self._update_file_list()

    def on_add_docxs(self) -> None:
        paths = filedialog.askopenfilenames(title="Seleccionar DOCXs", filetypes=[("DOCX", "*.docx")])
        for p in paths:
            self.batch_files.append(Path(p))
        self._update_file_list()

    def on_clear_list(self) -> None:
        self.batch_files.clear()
        self._update_file_list()

    def _update_file_list(self) -> None:
        self.file_listbox.configure(state="normal")
        self.file_listbox.delete("1.0", "end")
        for i, f in enumerate(self.batch_files, 1):
            self.file_listbox.insert("end", f"{i}. {f.name}\n")
        self.file_listbox.configure(state="disabled")

    def on_choose_outdir_batch(self) -> None:
        path = filedialog.askdirectory(title="Elegir carpeta de salida")
        if path:
            self.var_outdir_batch.set(path)

    def on_run_batch(self) -> None:
        if not self.batch_files:
            messagebox.showwarning("Sin archivos", "Agrega archivos a la lista primero.")
            return
        outdir = Path(self.var_outdir_batch.get()) if self.var_outdir_batch.get().strip() else Path.cwd()
        do_pdf2docx = bool(self.var_batch_pdf2docx.get())
        do_raster = bool(self.var_batch_raster.get())
        do_docx2pdf = bool(self.var_batch_docx2pdf.get())
        dpi = int(self.var_batch_dpi.get()) if str(self.var_batch_dpi.get()).strip() else 200
        overwrite = bool(self.var_batch_overwrite.get())
        th = threading.Thread(target=self._run_batch_task, args=(list(self.batch_files), outdir, do_pdf2docx, do_raster, do_docx2pdf, dpi, overwrite), daemon=True)
        th.start()

    def _run_batch_task(self, items: list[Path], outdir: Path, do_pdf2docx: bool, do_raster: bool, do_docx2pdf: bool, dpi: int, overwrite: bool) -> None:
        try:
            self._set_status("Procesando lote...", indeterminate=True)
            pdfs = [p for p in items if p.suffix.lower() == ".pdf"]
            docxs = [p for p in items if p.suffix.lower() == ".docx"]
            total = 0
            if do_pdf2docx:
                total += len(pdfs)
            if do_docx2pdf:
                total += len(docxs)

            done = 0
            if do_pdf2docx:
                for p in pdfs:
                    tgt = outdir / (p.stem + ".docx")
                    if do_raster:
                        pdf_to_docx_raster(p, tgt, dpi=dpi, overwrite=overwrite)
                    else:
                        pdf_to_docx(p, tgt, None, None, overwrite)
                    done += 1
                    self._update_progress(done, total)
            if do_docx2pdf:
                for d in docxs:
                    tgt = outdir / (d.stem + ".pdf")
                    docx_to_pdf(d, tgt, overwrite)
                    done += 1
                    self._update_progress(done, total)
        except Exception as e:
            self._set_status("Error en lote")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return
        self._set_status("Lote completado")
        self.after(0, lambda: messagebox.showinfo("Listo", f"Lote completado en:\n{outdir}"))

    def _update_progress(self, current: int, total: int) -> None:
        if total > 0:
            self.after(0, lambda: self.progress.set(current / total))

    def _set_status(self, text: str, indeterminate: bool = False) -> None:
        def update():
            self.lbl_status.configure(text=text)
            if indeterminate:
                self.progress.configure(mode="indeterminate")
                self.progress.start()
            else:
                self.progress.stop()
                self.progress.configure(mode="determinate")
                self.progress.set(0)
        self.after(0, update)


if __name__ == "__main__":
    app = Pdf2WordApp()
    app.mainloop()
