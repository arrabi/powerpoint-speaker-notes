
import sys
import time
import os
import argparse
from utils.pptx_tools import process_presentation


def pptx_to_pdf(pptx_path, pdf_path=None):
    import shutil
    soffice_path = shutil.which("soffice")
    if soffice_path is None:
        print("LibreOffice (soffice) is not installed or not in PATH. PDF export skipped.")
        return
    out_dir = os.path.dirname(os.path.abspath(pptx_path))
    cmd = f'"{soffice_path}" --headless --convert-to pdf --outdir "{out_dir}" "{pptx_path}"'
    result = os.system(cmd)
    if result == 0:
        pdf_file = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
        pdf_path = os.path.join(out_dir, pdf_file)
        print(f"PDF created: {pdf_path}")
    else:
        print("PDF conversion failed. Make sure LibreOffice is installed and accessible.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate speaker notes slides with images and optional external notes (Markdown).")
    parser.add_argument("input", help="Input PPTX file")
    parser.add_argument("output", nargs="?", help="Output PPTX file (optional)")
    parser.add_argument("--notes", help="Path to Markdown notes file")
    args = parser.parse_args()

    input_pptx = args.input

    # ensure output directory exists
    if not os.path.exists("data_out"):
        os.makedirs("data_out")
    if args.output:
        output_pptx = args.output
    else:
        input_filename = os.path.basename(input_pptx)
        output_pptx = f"data_out/{input_filename.rsplit('.', 1)[0]}_output_{time.strftime('%Y%m%d_%H%M%S')}.pptx"

    process_presentation(input_pptx, output_pptx, notes_path=args.notes)
    pptx_to_pdf(output_pptx)
