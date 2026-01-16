import argparse
from pathlib import Path
from typing import Optional

from tools import (
    pdf_to_docx,
    docx_to_pdf,
    compress_pdf,
    compress_docx_images,
    pdf_to_docx_raster,
    batch_pdf_to_docx,
    batch_docx_to_pdf,
    scan_files,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="convertor",
        description="Herramientas: PDF→DOCX, DOCX→PDF, compresión PDF y DOCX.",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    # pdf2docx
    p1 = sub.add_parser("pdf2docx", help="Convertir PDF a DOCX")
    p1.add_argument("input", help="Ruta al PDF")
    p1.add_argument("-o", "--output", help="Ruta del DOCX de salida")
    p1.add_argument("--start", type=int, help="Página inicial (1-basado)")
    p1.add_argument("--end", type=int, help="Página final (1-basado)")
    p1.add_argument("--overwrite", action="store_true", help="Sobrescribe si el DOCX existe")

    # pdf2docx-raster (máxima fidelidad visual)
    p1r = sub.add_parser("pdf2docx-raster", help="PDF → DOCX por imagen (máxima fidelidad, no editable)")
    p1r.add_argument("input", help="Ruta al PDF")
    p1r.add_argument("-o", "--output", help="Ruta del DOCX de salida")
    p1r.add_argument("--dpi", type=int, default=200, help="Resolución de render (por defecto 200 DPI)")
    p1r.add_argument("--overwrite", action="store_true", help="Sobrescribe si el DOCX existe")

    # docx2pdf
    p2 = sub.add_parser("docx2pdf", help="Convertir DOCX a PDF")
    p2.add_argument("input", help="Ruta al DOCX")
    p2.add_argument("-o", "--output", help="Ruta del PDF de salida")
    p2.add_argument("--overwrite", action="store_true", help="Sobrescribe si el PDF existe")

    # compress-pdf
    p3 = sub.add_parser("compress-pdf", help="Optimizar PDF (limpieza/deflate)")
    p3.add_argument("input", help="Ruta al PDF")
    p3.add_argument("-o", "--output", help="Ruta del PDF de salida (optimizado)", required=True)

    # compress-docx
    p4 = sub.add_parser("compress-docx", help="Comprimir imágenes dentro de DOCX")
    p4.add_argument("input", help="Ruta al DOCX")
    p4.add_argument("-o", "--output", help="Ruta del DOCX de salida (comprimido)", required=True)
    p4.add_argument("--quality", type=int, default=75, help="Calidad JPEG (1-95, por defecto 75)")
    p4.add_argument("--max-width", type=int, help="Ancho máximo de imagen")
    p4.add_argument("--max-height", type=int, help="Alto máximo de imagen")

    # batch (carpeta)
    p5 = sub.add_parser("batch", help="Procesar por lotes en una carpeta")
    p5.add_argument("input", help="Carpeta a procesar")
    p5.add_argument("--outdir", help="Carpeta de salida", required=True)
    p5.add_argument("--pdf2docx", action="store_true", help="Convertir todos los PDF a DOCX (editable)")
    p5.add_argument("--pdf2docx-raster", action="store_true", help="Convertir todos los PDF a DOCX por imagen (máxima fidelidad)")
    p5.add_argument("--dpi", type=int, default=200, help="DPI para modo raster")
    p5.add_argument("--docx2pdf", action="store_true", help="Convertir todos los DOCX a PDF")
    p5.add_argument("--overwrite", action="store_true", help="Sobrescribir archivos de salida si existen")

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.cmd == "pdf2docx":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".docx")
        pdf_to_docx(inp, out, args.start, args.end, args.overwrite)
        print(f"Conversión completada: {out}")

    elif args.cmd == "pdf2docx-raster":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".docx")
        dpi = getattr(args, 'dpi', 200)
        pdf_to_docx_raster(inp, out, dpi=dpi, overwrite=args.overwrite)
        print(f"Conversión (raster) completada: {out}")

    elif args.cmd == "docx2pdf":
        inp = Path(args.input)
        out = Path(args.output) if args.output else inp.with_suffix(".pdf")
        docx_to_pdf(inp, out, args.overwrite)
        print(f"Conversión completada: {out}")

    elif args.cmd == "compress-pdf":
        inp = Path(args.input)
        out = Path(args.output)
        compress_pdf(inp, out)
        print(f"PDF optimizado: {out}")

    elif args.cmd == "compress-docx":
        inp = Path(args.input)
        out = Path(args.output)
        q = max(1, min(95, args.quality))
        compress_docx_images(inp, out, quality=q, max_width=args.max_width, max_height=args.max_height)
        print(f"DOCX comprimido: {out}")

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
