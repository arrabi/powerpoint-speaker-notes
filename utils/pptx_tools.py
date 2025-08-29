from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image
import tempfile
import os

# Helper: add page number to slide
def add_page_number(slide, page_num):
    left = Inches(8.5) - Inches(1.2)
    top = Inches(7.5) - Inches(0.5)
    width = Inches(1)
    height = Inches(0.4)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = f"Page {page_num}"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = RGBColor(80, 80, 80)
    tf.paragraphs[0].alignment = PP_ALIGN.RIGHT

# Helper: export slide as image (requires PowerPoint or LibreOffice for best results)
def export_slide_as_image(prs, slide_idx, tmpdir):
    # python-pptx cannot render slides to images. This is a placeholder for manual or external rendering.
    # For MVP, create a placeholder image with slide number for clarity.
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
        # Export slide as image (placeholder)
        img_path = export_slide_as_image(prs, slide_idx + offset, tmpdir)
        # Duplicate slide after current
        blank_slide_layout = prs.slide_layouts[6]  # blank
        new_slide = prs.slides.add_slide(blank_slide_layout)
        # Move new slide to correct position
        prs.slides._sldIdLst.insert(slide_idx + offset + 1, prs.slides._sldIdLst[-1])
        # Add image to new slide (upper right, 0.2x slide size)
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        img_side = int(min(slide_width, slide_height) * 0.2)
        left = slide_width - img_side - int(slide_width * 0.05)
        top = int(slide_height * 0.05)
        new_slide.shapes.add_picture(
            img_path,
            left,
            top,
            width=img_side,
            height=img_side
        )
        os.remove(img_path)
        page_num += 1
    prs.save(output_pptx)
    print(f"Saved: {output_pptx}")
