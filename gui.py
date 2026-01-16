import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import Optional, Callable
from datetime import datetime

from tools import (
    pdf_to_docx, pdf_to_docx_with_progress,
    docx_to_pdf,
    compress_pdf_with_progress,
    compress_docx_images_with_progress,
    pdf_to_docx_raster, pdf_to_docx_raster_with_progress,
    ocr_pdf_to_docx_with_progress,
    convert_image,
    get_image_info, extract_images_from_pdf, extract_images_from_docx,
    SUPPORTED_IMAGE_FORMATS
)

# Configurar apariencia
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ProgressModal(ctk.CTkToplevel):
    """Modal de progreso con log detallado estilo terminal."""

    def __init__(self, parent, title: str = "Procesando..."):
        super().__init__(parent)
        self.title(title)
        self.geometry("550x420")
        self.resizable(False, False)

        # Centrar en pantalla
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 550) // 2
        y = (self.winfo_screenheight() - 420) // 2
        self.geometry(f"550x420+{x}+{y}")

        # Modal behavior
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close_attempt)

        self._completed = False
        self._cancelled = False
        self._build_ui(title)

    def _build_ui(self, title: str) -> None:
        # Header
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=20, pady=(20, 10))

        self.title_label = ctk.CTkLabel(
            header, text=title,
            font=ctk.CTkFont(size=18, weight="bold")
        )
        self.title_label.pack(side="left")

        self.status_label = ctk.CTkLabel(
            header, text="Iniciando...",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60")
        )
        self.status_label.pack(side="right")

        # Progress bar
        progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        progress_frame.pack(fill="x", padx=20, pady=10)

        self.progress_bar = ctk.CTkProgressBar(progress_frame, height=20, corner_radius=8)
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            progress_frame, text="0%",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.progress_label.pack(pady=(5, 0))

        # Log area (estilo terminal)
        log_frame = ctk.CTkFrame(self, fg_color=("gray85", "gray20"))
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.log_text = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=("gray95", "gray10"),
            text_color=("gray20", "gray80"),
            corner_radius=8
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)
        self.log_text.configure(state="disabled")

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))

        self.cancel_btn = ctk.CTkButton(
            btn_frame, text="Cancelar", width=120,
            fg_color="#F44336", hover_color="#D32F2F",
            command=self._on_cancel
        )
        self.cancel_btn.pack(side="left")

        self.close_btn = ctk.CTkButton(
            btn_frame, text="Cerrar", width=120,
            state="disabled",
            command=self._on_close
        )
        self.close_btn.pack(side="right")

    def log(self, message: str, level: str = "info") -> None:
        """Agrega un mensaje al log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = ""
        color = None

        if level == "success":
            prefix = "[OK] "
            color = "#4CAF50"
        elif level == "error":
            prefix = "[ERROR] "
            color = "#F44336"
        elif level == "warning":
            prefix = "[AVISO] "
            color = "#FF9800"
        elif level == "progress":
            prefix = ">>> "
            color = "#2196F3"

        def update():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", f"[{timestamp}] {prefix}{message}\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")

        self.after(0, update)

    def set_progress(self, current: int, total: int, message: str = "") -> None:
        """Actualiza la barra de progreso."""
        if total > 0:
            progress = current / total
            percent = int(progress * 100)

            def update():
                self.progress_bar.set(progress)
                self.progress_label.configure(text=f"{percent}%")
                if message:
                    self.status_label.configure(text=message)

            self.after(0, update)

    def set_status(self, status: str) -> None:
        """Actualiza el texto de estado."""
        self.after(0, lambda: self.status_label.configure(text=status))

    def set_title(self, title: str) -> None:
        """Actualiza el titulo."""
        self.after(0, lambda: self.title_label.configure(text=title))

    def complete(self, success: bool = True, message: str = "") -> None:
        """Marca el proceso como completado."""
        self._completed = True

        def update():
            self.progress_bar.set(1.0)
            self.progress_label.configure(text="100%")

            if success:
                self.status_label.configure(text="Completado", text_color="#4CAF50")
                self.log(message or "Proceso completado exitosamente", "success")
            else:
                self.status_label.configure(text="Error", text_color="#F44336")
                self.log(message or "El proceso termino con errores", "error")

            self.cancel_btn.configure(state="disabled")
            self.close_btn.configure(state="normal")

        self.after(0, update)

    def is_cancelled(self) -> bool:
        """Verifica si el usuario cancelo."""
        return self._cancelled

    def _on_cancel(self) -> None:
        self._cancelled = True
        self.log("Cancelando operacion...", "warning")
        self.cancel_btn.configure(state="disabled")
        self.status_label.configure(text="Cancelando...")

    def _on_close(self) -> None:
        self.grab_release()
        self.destroy()

    def _on_close_attempt(self) -> None:
        if self._completed:
            self._on_close()
        else:
            self._on_cancel()


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

        # Variables - Imagenes
        self.var_img_input = ctk.StringVar()
        self.var_img_output = ctk.StringVar()
        self.var_img_format = ctk.StringVar(value="PNG")
        self.var_img_quality = ctk.IntVar(value=95)
        self.var_img_width = ctk.StringVar()
        self.var_img_height = ctk.StringVar()
        self.var_img_maintain_aspect = ctk.BooleanVar(value=True)
        self.var_img_overwrite = ctk.BooleanVar(value=False)
        self.var_img_outdir = ctk.StringVar()
        self.img_batch_files: list[Path] = []

        # Variables - Extraccion
        self.var_extract_input = ctk.StringVar()
        self.var_extract_outdir = ctk.StringVar()
        self.var_extract_format = ctk.StringVar(value="PNG")

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
        tab3 = self.tabview.add("Imagenes")
        tab4 = self.tabview.add("Compresion")
        tab5 = self.tabview.add("Lotes")

        self._build_pdf2docx_tab(tab1)
        self._build_docx2pdf_tab(tab2)
        self._build_images_tab(tab3)
        self._build_compression_tab(tab4)
        self._build_batch_tab(tab5)

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

    def _build_images_tab(self, parent) -> None:
        main_frame = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # === Seccion 1: Conversion individual ===
        self._create_section_label(main_frame, "Convertir Imagen")

        single_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        single_frame.pack(fill="x", pady=(0, 20))
        single_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(single_frame, text="Imagen:", width=120, anchor="w").grid(row=0, column=0, padx=15, pady=12)
        ctk.CTkEntry(single_frame, textvariable=self.var_img_input, placeholder_text="Selecciona una imagen...").grid(row=0, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(single_frame, text="Examinar", width=100, command=self.on_browse_img_input).grid(row=0, column=2, padx=15, pady=12)

        ctk.CTkLabel(single_frame, text="Guardar como:", width=120, anchor="w").grid(row=1, column=0, padx=15, pady=12)
        ctk.CTkEntry(single_frame, textvariable=self.var_img_output, placeholder_text="Ruta de salida...").grid(row=1, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(single_frame, text="Guardar como", width=100, command=self.on_browse_img_output).grid(row=1, column=2, padx=15, pady=12)

        # Opciones de formato
        options_frame = ctk.CTkFrame(single_frame, fg_color="transparent")
        options_frame.grid(row=2, column=0, columnspan=3, padx=15, pady=10, sticky="ew")

        ctk.CTkLabel(options_frame, text="Formato:").pack(side="left", padx=(0, 5))
        format_menu = ctk.CTkOptionMenu(
            options_frame,
            variable=self.var_img_format,
            values=["JPG", "PNG", "WEBP", "ICO", "BMP", "GIF", "TIFF"],
            width=100,
            fg_color="#3B8ED0",
            button_color="#36719F",
            button_hover_color="#144870"
        )
        format_menu.pack(side="left", padx=10)

        ctk.CTkLabel(options_frame, text="Calidad:").pack(side="left", padx=(20, 5))
        ctk.CTkEntry(options_frame, textvariable=self.var_img_quality, width=60).pack(side="left", padx=5)

        ctk.CTkCheckBox(options_frame, text="Sobrescribir", variable=self.var_img_overwrite).pack(side="left", padx=20)

        # Redimension
        resize_frame = ctk.CTkFrame(single_frame, fg_color="transparent")
        resize_frame.grid(row=3, column=0, columnspan=3, padx=15, pady=10, sticky="ew")

        ctk.CTkLabel(resize_frame, text="Redimensionar:").pack(side="left", padx=(0, 10))
        ctk.CTkLabel(resize_frame, text="Ancho:").pack(side="left")
        ctk.CTkEntry(resize_frame, textvariable=self.var_img_width, width=70, placeholder_text="px").pack(side="left", padx=5)
        ctk.CTkLabel(resize_frame, text="Alto:").pack(side="left", padx=(10, 0))
        ctk.CTkEntry(resize_frame, textvariable=self.var_img_height, width=70, placeholder_text="px").pack(side="left", padx=5)
        ctk.CTkCheckBox(resize_frame, text="Mantener proporcion", variable=self.var_img_maintain_aspect).pack(side="left", padx=15)

        # Info de imagen
        self.img_info_label = ctk.CTkLabel(single_frame, text="", text_color=("gray50", "gray60"))
        self.img_info_label.grid(row=4, column=0, columnspan=2, padx=15, pady=5, sticky="w")

        ctk.CTkButton(
            single_frame, text="Convertir Imagen", width=160, height=45,
            fg_color="#9C27B0", hover_color="#7B1FA2",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.on_convert_image
        ).grid(row=4, column=2, padx=15, pady=15)

        # === Seccion 2: Conversion por lotes ===
        self._create_section_label(main_frame, "Conversion por Lotes")

        batch_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        batch_frame.pack(fill="x", pady=(0, 20))

        # Lista de imagenes
        self.img_listbox = ctk.CTkTextbox(batch_frame, height=120)
        self.img_listbox.pack(fill="x", padx=15, pady=15)
        self.img_listbox.configure(state="disabled")

        btn_frame = ctk.CTkFrame(batch_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 10))
        ctk.CTkButton(btn_frame, text="Agregar Imagenes", width=140, command=self.on_add_images).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Limpiar", width=100, fg_color="#F44336", hover_color="#D32F2F", command=self.on_clear_images).pack(side="left", padx=5)

        # Carpeta salida
        out_frame = ctk.CTkFrame(batch_frame, fg_color="transparent")
        out_frame.pack(fill="x", padx=15, pady=5)
        out_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(out_frame, text="Carpeta salida:", width=100, anchor="w").grid(row=0, column=0, padx=0, pady=5)
        ctk.CTkEntry(out_frame, textvariable=self.var_img_outdir, placeholder_text="Carpeta de destino...").grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(out_frame, text="Elegir", width=80, command=self.on_choose_img_outdir).grid(row=0, column=2, padx=0, pady=5)

        ctk.CTkButton(
            batch_frame, text="Convertir Lote", width=180, height=45,
            fg_color="#4CAF50", hover_color="#388E3C",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.on_convert_images_batch
        ).pack(pady=15)

        # === Seccion 3: Extraer imagenes ===
        self._create_section_label(main_frame, "Extraer Imagenes de Documentos")

        extract_frame = ctk.CTkFrame(main_frame, fg_color=("gray90", "gray17"))
        extract_frame.pack(fill="x", pady=(0, 15))
        extract_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(extract_frame, text="PDF o DOCX:", width=120, anchor="w").grid(row=0, column=0, padx=15, pady=12)
        ctk.CTkEntry(extract_frame, textvariable=self.var_extract_input, placeholder_text="Archivo PDF o DOCX...").grid(row=0, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(extract_frame, text="Examinar", width=100, command=self.on_browse_extract_input).grid(row=0, column=2, padx=15, pady=12)

        ctk.CTkLabel(extract_frame, text="Carpeta salida:", width=120, anchor="w").grid(row=1, column=0, padx=15, pady=12)
        ctk.CTkEntry(extract_frame, textvariable=self.var_extract_outdir, placeholder_text="Carpeta donde extraer...").grid(row=1, column=1, padx=5, pady=12, sticky="ew")
        ctk.CTkButton(extract_frame, text="Elegir", width=100, command=self.on_choose_extract_outdir).grid(row=1, column=2, padx=15, pady=12)

        extract_opts = ctk.CTkFrame(extract_frame, fg_color="transparent")
        extract_opts.grid(row=2, column=0, columnspan=3, padx=15, pady=10, sticky="ew")

        ctk.CTkLabel(extract_opts, text="Formato:").pack(side="left", padx=(0, 5))
        ctk.CTkOptionMenu(
            extract_opts,
            variable=self.var_extract_format,
            values=["PNG", "JPG", "WEBP"],
            width=100
        ).pack(side="left", padx=10)

        ctk.CTkButton(
            extract_frame, text="Extraer Imagenes", width=160, height=45,
            fg_color="#FF9800", hover_color="#F57C00",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.on_extract_images
        ).grid(row=2, column=2, padx=15, pady=15)

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

        ctk.CTkButton(options_frame, text="Iniciar Conversion", width=220, height=50, fg_color="#4CAF50", hover_color="#388E3C", font=ctk.CTkFont(size=16, weight="bold"), corner_radius=12, command=self.on_run_batch).pack(side="right", padx=15, pady=15)

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

        # Crear modal de progreso
        modal = ProgressModal(self, "Convirtiendo PDF a DOCX")
        modal.log(f"Archivo: {input_pdf.name}")
        modal.log(f"Salida: {output_docx.name}")

        def task():
            try:
                def progress_cb(current, total, msg):
                    modal.set_progress(current, total, msg)
                    modal.log(msg, "progress")

                pdf_to_docx_with_progress(
                    input_pdf, output_docx, start_i, end_i, overwrite,
                    progress_callback=progress_cb,
                    cancel_check=modal.is_cancelled
                )
                modal.complete(True, f"Archivo creado: {output_docx.name}")
            except InterruptedError:
                modal.complete(False, "Operacion cancelada por el usuario")
            except Exception as e:
                modal.log(str(e), "error")
                modal.complete(False, str(e))

        th = threading.Thread(target=task, daemon=True)
        th.start()

    def on_convert_pdf2docx_raster(self) -> None:
        in_path = self.var_input.get().strip()
        out_path = self.var_output.get().strip()
        if not in_path:
            messagebox.showwarning("Falta archivo", "Selecciona un archivo PDF de entrada.")
            return

        input_pdf = Path(in_path)
        output_docx = Path(out_path) if out_path else input_pdf.with_suffix(".docx")
        dpi = int(self.var_raster_dpi.get()) if str(self.var_raster_dpi.get()).strip() else 200
        overwrite = bool(self.var_overwrite.get())

        modal = ProgressModal(self, "Convirtiendo PDF a Imagen")
        modal.log(f"Archivo: {input_pdf.name}")
        modal.log(f"DPI: {dpi}")

        def task():
            try:
                def progress_cb(current, total, msg):
                    modal.set_progress(current, total, msg)
                    modal.log(msg, "progress")

                pdf_to_docx_raster_with_progress(
                    input_pdf, output_docx, dpi=dpi, overwrite=overwrite,
                    progress_callback=progress_cb,
                    cancel_check=modal.is_cancelled
                )
                modal.complete(True, f"Archivo creado: {output_docx.name}")
            except InterruptedError:
                modal.complete(False, "Operacion cancelada")
            except Exception as e:
                modal.log(str(e), "error")
                modal.complete(False, str(e))

        th = threading.Thread(target=task, daemon=True)
        th.start()

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

        modal = ProgressModal(self, "Conversion OCR")
        modal.log(f"Archivo: {input_pdf.name}")
        modal.log(f"Idioma: {lang} | DPI: {dpi}")

        def task():
            try:
                def progress_cb(current, total, msg):
                    modal.set_progress(current, total, msg)
                    modal.log(msg, "progress")

                ocr_pdf_to_docx_with_progress(
                    input_pdf, output_docx, dpi=dpi, lang=lang,
                    progress_callback=progress_cb,
                    cancel_check=modal.is_cancelled
                )
                modal.complete(True, f"Archivo creado: {output_docx.name}")
            except InterruptedError:
                modal.complete(False, "Operacion cancelada")
            except Exception as e:
                modal.log(str(e), "error")
                modal.complete(False, str(e))

        th = threading.Thread(target=task, daemon=True)
        th.start()

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

        input_path = Path(inp)
        output_path = Path(out)

        modal = ProgressModal(self, "Optimizando PDF")
        modal.log(f"Archivo: {input_path.name}")
        original_size = input_path.stat().st_size
        modal.log(f"Tamaño original: {self._format_size(original_size)}")

        def task():
            try:
                def progress_cb(current, total, msg):
                    modal.set_progress(current, total, msg)
                    modal.log(msg, "progress")

                result = compress_pdf_with_progress(
                    input_path, output_path,
                    progress_callback=progress_cb,
                    cancel_check=modal.is_cancelled
                )
                modal.log(f"Tamaño final: {self._format_size(result['new_size'])}")
                modal.log(f"Reduccion: {result['reduction_percent']:.1f}%", "success")
                modal.complete(True, f"PDF optimizado ({result['reduction_percent']:.1f}% reducido)")
            except InterruptedError:
                modal.complete(False, "Operacion cancelada")
            except Exception as e:
                modal.log(str(e), "error")
                modal.complete(False, str(e))

        th = threading.Thread(target=task, daemon=True)
        th.start()

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

        input_path = Path(inp)
        output_path = Path(out)

        modal = ProgressModal(self, "Comprimiendo DOCX")
        modal.log(f"Archivo: {input_path.name}")
        modal.log(f"Calidad JPEG: {q}%")
        original_size = input_path.stat().st_size
        modal.log(f"Tamaño original: {self._format_size(original_size)}")

        def task():
            try:
                def progress_cb(current, total, msg):
                    modal.set_progress(current, total, msg)
                    modal.log(msg, "progress")

                result = compress_docx_images_with_progress(
                    input_path, output_path, quality=q,
                    max_width=max_w, max_height=max_h,
                    progress_callback=progress_cb,
                    cancel_check=modal.is_cancelled
                )
                modal.log(f"Imagenes procesadas: {result['images_processed']}")
                modal.log(f"Tamaño final: {self._format_size(result['new_size'])}")
                modal.log(f"Reduccion: {result['reduction_percent']:.1f}%", "success")
                modal.complete(True, f"DOCX comprimido ({result['reduction_percent']:.1f}% reducido)")
            except InterruptedError:
                modal.complete(False, "Operacion cancelada")
            except Exception as e:
                modal.log(str(e), "error")
                modal.complete(False, str(e))

        th = threading.Thread(target=task, daemon=True)
        th.start()

    def _format_size(self, size_bytes: int) -> str:
        """Formatea bytes a formato legible."""
        for unit in ["B", "KB", "MB", "GB"]:
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} TB"

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
        items = list(self.batch_files)

        pdfs = [p for p in items if p.suffix.lower() == ".pdf"]
        docxs = [p for p in items if p.suffix.lower() == ".docx"]

        modal = ProgressModal(self, "Conversion por Lotes")
        modal.log(f"Archivos totales: {len(items)}")
        modal.log(f"PDFs: {len(pdfs)} | DOCXs: {len(docxs)}")
        modal.log(f"Carpeta destino: {outdir}")

        def task():
            try:
                total = 0
                if do_pdf2docx:
                    total += len(pdfs)
                if do_docx2pdf:
                    total += len(docxs)

                if total == 0:
                    modal.log("No hay archivos para procesar", "warning")
                    modal.complete(True, "Sin archivos para procesar")
                    return

                done = 0
                errors = []

                if do_pdf2docx:
                    mode = "imagen" if do_raster else "editable"
                    modal.log(f"Convirtiendo PDFs a DOCX (modo {mode})...")

                    for p in pdfs:
                        if modal.is_cancelled():
                            raise InterruptedError("Cancelado")

                        tgt = outdir / (p.stem + ".docx")
                        modal.log(f"Procesando: {p.name}", "progress")

                        try:
                            if do_raster:
                                pdf_to_docx_raster(p, tgt, dpi=dpi, overwrite=overwrite)
                            else:
                                pdf_to_docx(p, tgt, None, None, overwrite)
                            modal.log(f"Completado: {p.name}", "success")
                        except Exception as e:
                            modal.log(f"Error en {p.name}: {e}", "error")
                            errors.append(p.name)

                        done += 1
                        modal.set_progress(done, total, f"{done}/{total} archivos")

                if do_docx2pdf:
                    modal.log("Convirtiendo DOCXs a PDF...")

                    for d in docxs:
                        if modal.is_cancelled():
                            raise InterruptedError("Cancelado")

                        tgt = outdir / (d.stem + ".pdf")
                        modal.log(f"Procesando: {d.name}", "progress")

                        try:
                            docx_to_pdf(d, tgt, overwrite)
                            modal.log(f"Completado: {d.name}", "success")
                        except Exception as e:
                            modal.log(f"Error en {d.name}: {e}", "error")
                            errors.append(d.name)

                        done += 1
                        modal.set_progress(done, total, f"{done}/{total} archivos")

                if errors:
                    modal.complete(True, f"Completado con {len(errors)} errores")
                else:
                    modal.complete(True, f"Lote completado: {done} archivos")

            except InterruptedError:
                modal.complete(False, "Operacion cancelada")
            except Exception as e:
                modal.log(str(e), "error")
                modal.complete(False, str(e))

        th = threading.Thread(target=task, daemon=True)
        th.start()

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

    # --- Image handlers ---
    def on_browse_img_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar Imagen",
            filetypes=[
                ("Imagenes", "*.png *.jpg *.jpeg *.webp *.bmp *.gif *.tiff *.ico"),
                ("PNG", "*.png"),
                ("JPEG", "*.jpg *.jpeg"),
                ("WebP", "*.webp"),
                ("ICO", "*.ico"),
                ("Todos", "*.*"),
            ],
        )
        if path:
            self.var_img_input.set(path)
            # Proponer nombre de salida
            fmt = self.var_img_format.get().lower()
            out_path = Path(path).with_suffix(f".{fmt}")
            self.var_img_output.set(str(out_path))
            # Mostrar info
            try:
                info = get_image_info(Path(path))
                self.img_info_label.configure(
                    text=f"{info['width']}x{info['height']} px | {info['format']} | {info['file_size_human']}"
                )
            except Exception:
                self.img_info_label.configure(text="")

    def on_browse_img_output(self) -> None:
        fmt = self.var_img_format.get().lower()
        path = filedialog.asksaveasfilename(
            title="Guardar imagen como",
            defaultextension=f".{fmt}",
            filetypes=[(fmt.upper(), f"*.{fmt}")],
        )
        if path:
            self.var_img_output.set(path)

    def on_convert_image(self) -> None:
        inp = self.var_img_input.get().strip()
        out = self.var_img_output.get().strip()
        if not inp:
            messagebox.showwarning("Falta archivo", "Selecciona una imagen de entrada.")
            return
        if not out:
            messagebox.showwarning("Falta archivo", "Especifica la ruta de salida.")
            return

        fmt = self.var_img_format.get().lower()
        quality = int(self.var_img_quality.get()) if str(self.var_img_quality.get()).strip() else 95
        quality = max(1, min(100, quality))
        overwrite = bool(self.var_img_overwrite.get())

        # Redimension
        resize = None
        w_str = self.var_img_width.get().strip()
        h_str = self.var_img_height.get().strip()
        if w_str or h_str:
            try:
                w = int(w_str) if w_str else 0
                h = int(h_str) if h_str else 0
                if w > 0 or h > 0:
                    resize = (w if w > 0 else 10000, h if h > 0 else 10000)
            except ValueError:
                messagebox.showerror("Error", "Ancho y alto deben ser numeros enteros.")
                return

        maintain_aspect = bool(self.var_img_maintain_aspect.get())

        th = threading.Thread(
            target=self._convert_image_task,
            args=(Path(inp), Path(out), fmt, quality, resize, maintain_aspect, overwrite),
            daemon=True
        )
        th.start()

    def _convert_image_task(self, inp: Path, out: Path, fmt: str, quality: int, resize, maintain_aspect: bool, overwrite: bool) -> None:
        try:
            self._set_status("Convirtiendo imagen...", indeterminate=True)
            convert_image(inp, out, fmt, quality, resize, maintain_aspect, overwrite)
        except Exception as e:
            self._set_status("Error al convertir imagen")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return
        self._set_status("Imagen convertida")
        self.after(0, lambda: messagebox.showinfo("Listo", f"Imagen guardada:\n{out}"))

    def on_add_images(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Seleccionar Imagenes",
            filetypes=[
                ("Imagenes", "*.png *.jpg *.jpeg *.webp *.bmp *.gif *.tiff *.ico"),
                ("Todos", "*.*"),
            ],
        )
        for p in paths:
            self.img_batch_files.append(Path(p))
        self._update_img_list()

    def on_clear_images(self) -> None:
        self.img_batch_files.clear()
        self._update_img_list()

    def _update_img_list(self) -> None:
        self.img_listbox.configure(state="normal")
        self.img_listbox.delete("1.0", "end")
        for i, f in enumerate(self.img_batch_files, 1):
            self.img_listbox.insert("end", f"{i}. {f.name}\n")
        self.img_listbox.configure(state="disabled")

    def on_choose_img_outdir(self) -> None:
        path = filedialog.askdirectory(title="Elegir carpeta de salida")
        if path:
            self.var_img_outdir.set(path)

    def on_convert_images_batch(self) -> None:
        if not self.img_batch_files:
            messagebox.showwarning("Sin archivos", "Agrega imagenes a la lista primero.")
            return

        outdir = Path(self.var_img_outdir.get()) if self.var_img_outdir.get().strip() else Path.cwd()
        fmt = self.var_img_format.get().lower()
        quality = int(self.var_img_quality.get()) if str(self.var_img_quality.get()).strip() else 95
        overwrite = bool(self.var_img_overwrite.get())
        files = list(self.img_batch_files)

        # Redimension
        resize = None
        w_str = self.var_img_width.get().strip()
        h_str = self.var_img_height.get().strip()
        if w_str or h_str:
            try:
                w = int(w_str) if w_str else 0
                h = int(h_str) if h_str else 0
                if w > 0 or h > 0:
                    resize = (w if w > 0 else 10000, h if h > 0 else 10000)
            except ValueError:
                pass

        maintain_aspect = bool(self.var_img_maintain_aspect.get())

        modal = ProgressModal(self, "Convirtiendo Imagenes")
        modal.log(f"Imagenes: {len(files)}")
        modal.log(f"Formato destino: {fmt.upper()}")
        modal.log(f"Calidad: {quality}%")

        def task():
            try:
                total = len(files)
                done = 0
                errors = []

                for f in files:
                    if modal.is_cancelled():
                        raise InterruptedError("Cancelado")

                    ext = fmt if fmt != "jpeg" else "jpg"
                    target = outdir / (f.stem + "." + ext)
                    modal.log(f"Convirtiendo: {f.name}", "progress")

                    try:
                        convert_image(f, target, fmt, quality, resize, maintain_aspect, overwrite)
                        modal.log(f"Completado: {f.name} -> {target.name}", "success")
                    except Exception as e:
                        modal.log(f"Error en {f.name}: {e}", "error")
                        errors.append(f.name)

                    done += 1
                    modal.set_progress(done, total, f"{done}/{total} imagenes")

                if errors:
                    modal.complete(True, f"Completado con {len(errors)} errores")
                else:
                    modal.complete(True, f"Convertidas {done} imagenes")

            except InterruptedError:
                modal.complete(False, "Operacion cancelada")
            except Exception as e:
                modal.log(str(e), "error")
                modal.complete(False, str(e))

        th = threading.Thread(target=task, daemon=True)
        th.start()

    # --- Extract images handlers ---
    def on_browse_extract_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Seleccionar PDF o DOCX",
            filetypes=[
                ("Documentos", "*.pdf *.docx"),
                ("PDF", "*.pdf"),
                ("DOCX", "*.docx"),
            ],
        )
        if path:
            self.var_extract_input.set(path)
            # Proponer carpeta de salida
            parent = Path(path).parent / f"{Path(path).stem}_imagenes"
            self.var_extract_outdir.set(str(parent))

    def on_choose_extract_outdir(self) -> None:
        path = filedialog.askdirectory(title="Elegir carpeta de salida")
        if path:
            self.var_extract_outdir.set(path)

    def on_extract_images(self) -> None:
        inp = self.var_extract_input.get().strip()
        outdir = self.var_extract_outdir.get().strip()
        if not inp:
            messagebox.showwarning("Falta archivo", "Selecciona un archivo PDF o DOCX.")
            return
        if not outdir:
            messagebox.showwarning("Falta carpeta", "Especifica la carpeta de salida.")
            return

        fmt = self.var_extract_format.get().lower()
        th = threading.Thread(
            target=self._extract_images_task,
            args=(Path(inp), Path(outdir), fmt),
            daemon=True
        )
        th.start()

    def _extract_images_task(self, inp: Path, outdir: Path, fmt: str) -> None:
        try:
            self._set_status("Extrayendo imagenes...", indeterminate=True)
            ext = inp.suffix.lower()
            if ext == ".pdf":
                extracted = extract_images_from_pdf(inp, outdir, fmt)
            elif ext == ".docx":
                extracted = extract_images_from_docx(inp, outdir)
            else:
                raise ValueError(f"Formato no soportado: {ext}")
        except Exception as e:
            self._set_status("Error al extraer")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            return

        self._set_status("Extraccion completada")
        self.after(0, lambda: messagebox.showinfo("Listo", f"Extraidas {len(extracted)} imagenes en:\n{outdir}"))


if __name__ == "__main__":
    app = Pdf2WordApp()
    app.mainloop()
