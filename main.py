import sys
import time
from utils.pptx_tools import process_presentation

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python main.py input.pptx output.pptx")
        sys.exit(1)
    input_pptx = sys.argv[1]

    #if argv[2] is missing, then let the outputfile name = <input filename>_output_datetimestamp.pptx
    if len(sys.argv) < 3:
        output_pptx = f"{input_pptx.rsplit('.', 1)[0]}_output_{int(time.time())}.pptx"
    else:
        output_pptx = sys.argv[2]
    process_presentation(input_pptx, output_pptx)
