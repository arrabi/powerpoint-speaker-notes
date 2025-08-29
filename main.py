
import sys
import time
import os
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
    if len(sys.argv) < 2:
        print("Usage: python main.py input.pptx output.pptx")
        sys.exit(1)
    input_pptx = sys.argv[1]

    #if argv[2] is missing, then let the outputfile name = <input filename>_output_datetimestamp.pptx
    #make sure the output file is in the data_out/ folder
    if not os.path.exists("data_out"):
        os.makedirs("data_out")
    if len(sys.argv) < 3:
        #remember to remove the folder name from input_pptx name
        #input_pptx will have a value with a folder e.g. data/test_input.pptx
        input_filename = os.path.basename(input_pptx)
        #the datetime stamp should be text readable
        output_pptx = f"data_out/{input_filename.rsplit('.', 1)[0]}_output_{time.strftime('%Y%m%d_%H%M%S')}.pptx"
    else:
        output_pptx = sys.argv[2]
    process_presentation(input_pptx, output_pptx)
    pptx_to_pdf(output_pptx)
