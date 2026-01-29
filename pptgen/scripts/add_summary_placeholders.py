#!/usr/bin/env python3
"""Add summary placeholders to the Insight template."""

from pptx import Presentation
from pptx.util import Inches, Pt
import os
import tempfile
import shutil
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(ROOT, "templates", "2026_Insight_PPT_Template.potx")
OUTPUT = os.path.join(ROOT, "templates", "2026_Insight_PPT_Template_Summary.potx")

# Convert .potx to .pptx temporarily
temp_dir = tempfile.mkdtemp()
temp_pptx = os.path.join(temp_dir, "temp.pptx")

with zipfile.ZipFile(TEMPLATE_PATH, 'r') as zip_in:
    with zipfile.ZipFile(temp_pptx, 'w') as zip_out:
        for item in zip_in.infolist():
            data = zip_in.read(item.filename)
            if item.filename == '[Content_Types].xml':
                data = data.replace(
                    b'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
                    b'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
                )
            zip_out.writestr(item, data)

# Load the insight template
prs = Presentation(temp_pptx)

# Get the first slide
if len(prs.slides) > 0:
    slide = prs.slides[0]
else:
    # Create a slide if none exists
    layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(layout)

# Add textbox with summary placeholders
left = Inches(0.5)
top = Inches(1.5)
width = Inches(9)
height = Inches(5.5)
text_box = slide.shapes.add_textbox(left, top, width, height)
tf = text_box.text_frame
tf.word_wrap = True

fields = [
    ("Total Revenue", "{Total Revenue}"),
    ("Total Projects", "{Total Projects}"),
    ("Total Clients", "{Total Clients}"),
    ("Won Projects", "{Won Projects}"),
    ("", ""),
    ("Average Revenue per Project", "{Average Revenue per Project}"),
    ("Total GP $", "{Total GP $}"),
    ("Average GP %", "{Average GP %}"),
    ("Top Client", "{Top Client}"),
    ("Top Client Revenue", "{Top Client Revenue}"),
]

for i, (label, placeholder) in enumerate(fields):
    if i == 0:
        p = tf.paragraphs[0]
    else:
        p = tf.add_paragraph()
    if label:
        p.text = f"{label}: {placeholder}"
    else:
        p.text = ""
    p.font.size = Pt(18)

prs.save(OUTPUT)
print(f"Created: {OUTPUT}")

# Cleanup
shutil.rmtree(temp_dir, ignore_errors=True)

