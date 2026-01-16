import argparse
from pathlib import Path
from typing import Optional

from pdf2docx import Converter


def convert_file(input_pdf: Path, output_docx: Path, start_page: Optional[int], end_page: Optional[int], overwrite: bool) -> None:
    if not input_pdf.exists():
        raise FileNotFoundError(f"No existe el PDF: {input_pdf}")

    output_docx = output_docx.with_suffix(".docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)

    if output_docx.exists() and not overwrite:
        raise FileExistsError(f"El archivo de salida ya existe: {output_docx}. Usa --overwrite para reemplazarlo.")

    start_arg = 0 if start_page is None else max(0, start_page - 1)
    end_arg = None if end_page is None else max(start_arg, end_page - 1)

    cv = Converter(str(input_pdf))
    try:
        cv.convert(str(output_docx), start=start_arg, end=end_arg)
    finally:
        cv.close()


def batch_convert(input_dir: Path, out_dir: Path, start_page: Optional[int], end_page: Optional[int], overwrite: bool) -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"No existe la carpeta: {input_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        print("No se encontraron archivos .pdf en la carpeta.")
        return

    for pdf in pdf_files:
        target = out_dir / (pdf.stem + ".docx")
        try:
            print(f"Convirtiendo: {pdf} -> {target}")
            convert_file(pdf, target, start_page, end_page, overwrite)
            print(f"OK: {target}")
        except Exception as e:
            print(f"ERROR con {pdf}: {e}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf2word",
        description="Convierte PDF a DOCX (Word). Soporta archivo único o carpeta.",
    )
    parser.add_argument("input", help="Ruta al PDF o a una carpeta con PDFs")
    parser.add_argument("-o", "--output", help="Ruta del archivo DOCX de salida (para un PDF)")
    parser.add_argument("--outdir", help="Carpeta de salida (para conversión de carpeta)")
    parser.add_argument("--start", type=int, help="Página inicial (1-basado)")
    parser.add_argument("--end", type=int, help="Página final (1-basado)")
    parser.add_argument("--overwrite", action="store_true", help="Sobrescribe si el DOCX existe")
    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input)

    if input_path.is_file():
        if input_path.suffix.lower() != ".pdf":
            raise ValueError("El archivo de entrada debe ser un PDF (.pdf)")

        if args.output:
            output_path = Path(args.output)
        else:
            output_path = input_path.with_suffix(".docx")

        print(f"Convirtiendo archivo: {input_path} -> {output_path}")
        convert_file(input_path, output_path, args.start, args.end, args.overwrite)
        print(f"Conversión completada: {output_path}")

    elif input_path.is_dir():
        out_dir = Path(args.outdir) if args.outdir else input_path
        print(f"Convirtiendo carpeta: {input_path} -> {out_dir}")
        batch_convert(input_path, out_dir, args.start, args.end, args.overwrite)
        print("Conversión por lotes finalizada.")
    else:
        raise FileNotFoundError("La ruta de entrada no existe.")


if __name__ == "__main__":
    main()
