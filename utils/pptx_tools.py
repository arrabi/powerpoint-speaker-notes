from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image
import tempfile
import os
import subprocess

# Helper: add page number to slide
def add_page_number(slide, page_num):
    left = Inches(8.5) - Inches(1.2)
    top = Inches(7.5) - Inches(0.5)
    width = Inches(1)
    height = Inches(0.4)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    # Ensure only one paragraph and set just the numeric slide number
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"{page_num}"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor(80, 80, 80)
    p.alignment = PP_ALIGN.RIGHT

# Helper: export slide as image (requires PowerPoint or LibreOffice for best results)
import hashlib
import shutil

def get_slide_images(input_pptx):
    # Hash the pptx file to create a unique temp folder
    with open(input_pptx, 'rb') as f:
        pptx_hash = hashlib.md5(f.read()).hexdigest()
    base = os.path.basename(input_pptx)
    name = os.path.splitext(base)[0]
    base_dir = os.path.join('temp', f'{name}-{pptx_hash}')
    img_dir = os.path.join(base_dir, 'slides')
    os.makedirs(img_dir, exist_ok=True)

    # Reuse if already exported
    first_img = os.path.join(img_dir, 'slide_output-1.png')
    if os.path.exists(first_img):
        return img_dir

    # Locate LibreOffice
    soffice_path = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice_path is None:
        raise RuntimeError("LibreOffice (soffice) is not installed or not in PATH. Cannot export slides.")

    # 1) Convert PPTX to PDF into base_dir
    try:
        subprocess.run([
            soffice_path, "--headless", "--convert-to", "pdf", "--outdir", base_dir, input_pptx
        ], check=True)
    except subprocess.CalledProcessError:
        raise RuntimeError("LibreOffice failed to convert PPTX to PDF.")

    pdf_path = os.path.join(base_dir, f"{name}.pdf")
    if not os.path.exists(pdf_path):
        # Some LO builds may name output differently; attempt to find the first .pdf in base_dir
        candidates = [f for f in os.listdir(base_dir) if f.lower().endswith('.pdf')]
        if candidates:
            pdf_path = os.path.join(base_dir, candidates[0])
        else:
            raise RuntimeError("PDF not found after LibreOffice conversion.")

    # 2) Convert PDF to PNG images using pdftoppm
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm is None:
        raise RuntimeError("pdftoppm (poppler-utils) is not installed. Please install it to export images.")

    prefix = os.path.join(img_dir, 'slide_output')
    try:
        subprocess.run([pdftoppm, "-png", pdf_path, prefix], check=True)
    except subprocess.CalledProcessError:
        raise RuntimeError("pdftoppm failed to convert PDF to images.")

    return img_dir

def export_slide_as_image(prs, slide_idx, tmpdir, input_pptx=None):
    # Use real slide image if available
    if input_pptx:
        img_dir = get_slide_images(input_pptx)
        img_path = os.path.join(img_dir, f'slide_output-{slide_idx+1}.png')
        if os.path.exists(img_path):
            return img_path
    # fallback: placeholder
    img_path = os.path.join(tmpdir, f"slide_{slide_idx+1}.png")
    img = Image.new('RGB', (1280, 720), color=(240, 240, 240))
    from PIL import ImageDraw, ImageFont
    draw = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 80)
    except:
        font = ImageFont.load_default()
    draw.text((img.width//2-100, img.height//2-40), f"Slide {slide_idx+1}", fill=(100,100,100), font=font)
    img.save(img_path)
    return img_path

# Main processing function
def process_presentation(input_pptx, output_pptx):
    prs = Presentation(input_pptx)
    tmpdir = tempfile.mkdtemp()
    orig_slide_count = len(prs.slides)
    slide_indices = list(range(orig_slide_count))
    page_num = 1
    for offset, slide_idx in enumerate(slide_indices):
        slide = prs.slides[slide_idx + offset]
        add_page_number(slide, page_num)
        # Export slide as image (real image if possible)
        img_path = export_slide_as_image(prs, slide_idx, tmpdir, input_pptx=input_pptx)
        # Duplicate slide after current
        blank_slide_layout = prs.slide_layouts[6]  # blank
        new_slide = prs.slides.add_slide(blank_slide_layout)
        # Move new slide to correct position
        prs.slides._sldIdLst.insert(slide_idx + offset + 1, prs.slides._sldIdLst[-1])
        # Add image to new slide (upper-right, preserve aspect ratio at ~30% slide width)
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        margin = int(slide_width * 0.05)
        try:
            with Image.open(img_path) as im:
                img_w, img_h = im.size
        except Exception:
            img_w, img_h = (1280, 720)
        # target width is 30% of slide width (50% bigger than previous 20%)
        target_w = int(slide_width * 0.30)
        target_h = int(target_w * img_h / img_w)
        # ensure height fits within slide minus margins
        max_h = int(slide_height - 2 * margin)
        if target_h > max_h and max_h > 0:
            target_h = max_h
            target_w = int(target_h * img_w / img_h)
        left = int(slide_width - target_w - margin)
        top = int(slide_height * 0.05)
        new_slide.shapes.add_picture(img_path, left, top, width=target_w, height=target_h)
        page_num += 1
    prs.save(output_pptx)
    print(f"Saved: {output_pptx}")
