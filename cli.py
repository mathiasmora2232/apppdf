import argparse
from pathlib import Path

from tools import (
    analyze_pdf,
    pdf_to_docx,
    pdf_to_docx_smart,
    docx_to_pdf,
    compress_pdf,
    compress_docx_images,
    pdf_to_docx_raster,
    batch_pdf_to_docx,
    batch_docx_to_pdf,
    scan_files,
    ocr_pdf_to_docx,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="convertor",
        description="Herramientas: auditoría, PDF→DOCX, DOCX→PDF, compresión PDF y DOCX.",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    pa = sub.add_parser("analyze", help="Auditar PDF y sugerir mejor estrategia")
    pa.add_argument("input", help="Ruta al PDF")

    p1 = sub.add_parser("pdf2docx", help="Convertir PDF a DOCX editable")
    p1.add_argument("input", help="Ruta al PDF")
    p1.add_argument("-o", "--output", help="Ruta del DOCX de salida")
    p1.add_argument("--start", type=int, help="Página inicial (1-basado)")
    p1.add_argument("--end", type=int, help="Página final (1-basado)")
    p1.add_argument("--overwrite", action="store_true", help="Sobrescribe si el DOCX existe")

    ps = sub.add_parser("smart-pdf2docx", help="Conversión inteligente: editable con fallback 1:1")
    ps.add_argument("input", help="Ruta al PDF")
    ps.add_argument("-o", "--output", help="Ruta del DOCX de salida")
    ps.add_argument("--start", type=int, help="Página inicial (1-basado)")
    ps.add_argument("--end", type=int, help="Página final (1-basado)")
    ps.add_argument("--ocr-lang", default="spa", help="Idioma OCR para PDFs escaneados")
    ps.add_argument("--raster-dpi", type=int, default=220, help="DPI del fallback 1:1")
    ps.add_argument("--overwrite", action="store_true", help="Sobrescribe si el DOCX existe")

    p1r = sub.add_parser("pdf2docx-raster", help="PDF → DOCX por imagen (máxima fidelidad, no editable)")
    p1r.add_argument("input", help="Ruta al PDF")
    p1r.add_argument("-o", "--output", help="Ruta del DOCX de salida")
    p1r.add_argument("--dpi", type=int, default=200, help="Resolución de render")
    p1r.add_argument("--overwrite", action="store_true", help="Sobrescribe si el DOCX existe")

    pocr = sub.add_parser("ocr-pdf2docx", help="OCR: PDF (imagen) → DOCX (texto)")
    pocr.add_argument("input", help="Ruta al PDF")
    pocr.add_argument("-o", "--output", help="Ruta del DOCX de salida")
    pocr.add_argument("--dpi", type=int, default=300, help="DPI para render de páginas")
    pocr.add_argument("--lang", default="spa", help="Idioma Tesseract, ej.: spa, eng, spa+eng")

    p2 = sub.add_parser("docx2pdf", help="Convertir DOCX a PDF")
    p2.add_argument("input", help="Ruta al DOCX")
    p2.add_argument("-o", "--output", help="Ruta del PDF de salida")
    p2.add_argument("--overwrite", action="store_true", help="Sobrescribe si el PDF existe")

    p3 = sub.add_parser("compress-pdf", help="Comprimir/optimizar PDF")
    p3.add_argument("input", help="Ruta al PDF")
    p3.add_argument("-o", "--output", required=True, help="Ruta PDF de salida")
    p3.add_argument("--mode", choices=["lossless", "balanced", "aggressive"], default="lossless")
    p3.add_argument("--quality", type=int, default=65, help="Calidad JPEG para modos con pérdida")
    p3.add_argument("--dpi", type=int, default=150, help="DPI de rasterización para modos con pérdida")

    p4 = sub.add_parser("compress-docx", help="Comprimir imágenes dentro de DOCX")
    p4.add_argument("input", help="Ruta al DOCX")
    p4.add_argument("-o", "--output", required=True, help="Ruta DOCX de salida")
    p4.add_argument("--quality", type=int, default=75, help="Calidad JPEG (1-95)")
    p4.add_argument("--max-width", type=int, help="Ancho máximo de imagen")
    p4.add_argument("--max-height", type=int, help="Alto máximo de imagen")

    p5 = sub.add_parser("batch", help="Procesar por lotes en una carpeta")
    p5.add_argument("input", help="Carpeta a procesar")
    p5.add_argument("--outdir", required=True, help="Carpeta de salida")
    p5.add_argument("--pdf2docx", action="store_true", help="Convertir PDF a DOCX editable")
    p5.add_argument("--pdf2docx-raster", action="store_true", help="Convertir PDF a DOCX por imagen")
    p5.add_argument("--dpi", type=int, default=200, help="DPI para modo raster")
    p5.add_argument("--docx2pdf", action="store_true", help="Convertir DOCX a PDF")
    p5.add_argument("--overwrite", action="store_true", help="Sobrescribir archivos de salida")

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.cmd == "analyze":
        profile = analyze_pdf(Path(args.input))
        print("=== Auditoría PDF ===")
        for key, value in profile.items():
            print(f"{key}: {value}")

    elif args.cmd == "pdf2docx":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".docx")
        pdf_to_docx(inp, out, args.start, args.end, args.overwrite)
        print(f"Conversión completada: {out}")

    elif args.cmd == "smart-pdf2docx":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".docx")
        result = pdf_to_docx_smart(
            inp,
            out,
            start=args.start,
            end=args.end,
            overwrite=args.overwrite,
            ocr_lang=args.ocr_lang,
            raster_dpi=args.raster_dpi,
        )
        print(f"Conversión completada ({result['mode_used']}): {out}")

    elif args.cmd == "pdf2docx-raster":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".docx")
        pdf_to_docx_raster(inp, out, dpi=args.dpi, overwrite=args.overwrite)
        print(f"Conversión (raster) completada: {out}")

    elif args.cmd == "ocr-pdf2docx":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".docx")
        ocr_pdf_to_docx(inp, out, dpi=args.dpi, lang=args.lang)
        print(f"OCR completado (texto): {out}")

    elif args.cmd == "docx2pdf":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".pdf")
        docx_to_pdf(inp, out, args.overwrite)
        print(f"Conversión completada: {out}")

    elif args.cmd == "compress-pdf":
        result = compress_pdf(
            Path(args.input),
            Path(args.output),
            mode=args.mode,
            image_quality=max(20, min(95, args.quality)),
            dpi=max(72, min(300, args.dpi)),
        )
        print(
            f"PDF optimizado: {args.output} | reducción: {result['reduction_percent']:.1f}% "
            f"| modo: {result['mode']}"
        )

    elif args.cmd == "compress-docx":
        q = max(1, min(95, args.quality))
        compress_docx_images(
            Path(args.input),
            Path(args.output),
            quality=q,
            max_width=args.max_width,
            max_height=args.max_height,
        )
        print(f"DOCX comprimido: {args.output}")

    elif args.cmd == "batch":
        folder = Path(args.input)
        outdir = Path(args.outdir)
        pdfs, docxs = scan_files(folder)
        if args.pdf2docx or args.pdf2docx_raster:
            mode = "raster" if args.pdf2docx_raster else "editable"
            ok, errs = batch_pdf_to_docx(pdfs, outdir, mode=mode, overwrite=args.overwrite, dpi=args.dpi)
            print(f"PDF→DOCX ({mode}) completado en: {outdir} (ok={ok}, errores={len(errs)})")
            for f, msg in errs:
                print(f" - ERROR {f}: {msg}")
        if args.docx2pdf:
            ok, errs = batch_docx_to_pdf(docxs, outdir, overwrite=args.overwrite)
            print(f"DOCX→PDF completado en: {outdir} (ok={ok}, errores={len(errs)})")
            for f, msg in errs:
                print(f" - ERROR {f}: {msg}")


if __name__ == "__main__":
    main()
