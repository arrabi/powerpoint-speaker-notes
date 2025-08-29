# Powerpoint-Speaker-Notes (PSN) - Product Requirements Document (PRD)

## 1. Overview
PowerPoint's digital speaker view is useful for presenters, but it cannot be printed or used as physical flip cards. The Powerpoint-Speaker-Notes (PSN) project aims to generate a new slide after each original slide, containing the speaker notes and a screenshot of the original slide, to enable presenters to use printed flip cards.

## 2. Goals
- Allow presenters to print or use physical flip cards with speaker notes and slide images.
- Automate the process of generating speaker-note-slides for any PowerPoint presentation.

## 3. Features
- **Slide Duplication:** For each slide in the original presentation, insert a new slide immediately after it.
- **Speaker Notes Extraction:** Copy the speaker notes from the original slide to the new slide.
- **External Notes Upload:** Allow users to upload a PDF or Markdown file containing speaker notes for each slide. The app will parse these files and insert the provided notes into the corresponding speaker-note-slides.
- **Slide Screenshot:** Insert a screenshot or image of the original slide into the new slide.
- **Batch Processing:** Support processing of entire presentations in one operation.
- **Output Format:** Save the modified presentation as a new file, preserving the original.
 - **Original Slide Page Numbers:** Add a page number to each original slide, ensuring the numbering only counts the original slides and not the inserted speaker notes slides. This prevents confusion in page count and maintains accurate slide references for presenters and printed materials.

## 4. User Stories
- As a presenter, I want to print my speaker notes with slide images so I can use them as flip cards.
- As a user, I want the tool to work with any .pptx file without manual editing.
- As a user, I want the original presentation to remain unchanged.

## 5. Technical Requirements
- **Input:** Standard PowerPoint (.pptx) files.
- **Output:** Modified .pptx file with speaker-note-slides inserted.
- **Dependencies:**
  - Python 3.x
  - `python-pptx` for PowerPoint file manipulation
  - Image processing library (e.g., Pillow)
  - (Optional) Library for rendering slides to images (e.g., `unoconv`, `libreoffice`, or a cloud API)

## 6. Recommended Technology Stack

- **Language:** Python 3.x
- **PowerPoint Manipulation:** python-pptx
- **Image Processing:** Pillow (PIL)
- **Slide Screenshot Generation:** LibreOffice (command line/unoconv), or pptx2pdf + pdf2image, or manual screenshots
- **PDF Parsing:** PyPDF2 or pdfplumber
- **Markdown Parsing:** markdown or mistune

This stack is chosen for rapid prototyping, scriptability, and strong library support for the required features.

## 7. Project Organization

Suggested folder structure for maintainability and clarity:

```
powerpoint-speaker-notes/
├── main.py                  # Main script to run the utility
├── requirements.txt         # Python dependencies
├── README.md
├── PSN_spec.md
├── .github/
│   └── copilot-instructions.md
├── utils/
│   ├── __init__.py
│   ├── pptx_tools.py        # pptx manipulation helpers
│   ├── image_tools.py       # slide image generation helpers
│   ├── notes_parser.py      # PDF/Markdown notes parsing helpers
├── samples/
│   ├── example.pptx         # Sample input files
│   ├── notes.md
│   └── notes.pdf
└── output/
  └── (generated files go here)
```

This structure keeps the project modular, easy to test, and simple to extend.

## 8. Non-Goals
- Editing or formatting speaker notes content
- Supporting non-PowerPoint formats (e.g., Google Slides, PDF)

## 9. Open Questions
- What is the best method to generate high-quality slide screenshots?
- Should the tool support custom slide templates for the speaker-note-slides?

## 10. Success Metrics
- The tool can process a presentation and generate a new .pptx with speaker-note-slides for every original slide.
- The output is printable and usable as flip cards.

---
_Last updated: August 29, 2025_
