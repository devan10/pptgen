#!/usr/bin/env python3
"""Simple PPT generator.

Reads a CSV or Excel file and for each row adds a slide to the template.
Placeholders in template slides should use the form `{column_name}`.
"""

import argparse
import os
import shutil
import tempfile
import zipfile
import pandas as pd
from pptx import Presentation


def replace_placeholders(slide, row):
    for shape in slide.shapes:
        if not hasattr(shape, "text"):
            continue
        text = shape.text
        for col in row.index:
            placeholder = "{" + str(col) + "}"
            if placeholder in text:
                val = "" if pd.isna(row[col]) else str(row[col])
                text = text.replace(placeholder, val)
        if text != shape.text:
            try:
                shape.text = text
            except Exception:
                if hasattr(shape, "text_frame"):
                    shape.text_frame.clear()
                    shape.text_frame.text = text


def build_ppt(data_path, template_path, output_path, sheet=None):
    if data_path.lower().endswith(".csv"):
        df = pd.read_csv(data_path)
    else:
        df = pd.read_excel(data_path, sheet_name=sheet)

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    # Handle .potx (template) files by converting to .pptx for loading
    actual_template = template_path
    temp_dir = None
    if template_path.lower().endswith(".potx"):
        temp_dir = tempfile.mkdtemp()
        actual_template = os.path.join(temp_dir, "temp_template.pptx")
        # Convert .potx to .pptx by adjusting the content type in [Content_Types].xml
        with zipfile.ZipFile(template_path, 'r') as zip_in:
            with zipfile.ZipFile(actual_template, 'w') as zip_out:
                for item in zip_in.infolist():
                    data = zip_in.read(item.filename)
                    # Replace the presentation template content type with presentation type
                    if item.filename == '[Content_Types].xml':
                        data = data.replace(
                            b'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml',
                            b'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'
                        )
                    zip_out.writestr(item, data)
    
    try:
        prs = Presentation(actual_template)
        layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]

        for _, row in df.iterrows():
            slide = prs.slides.add_slide(layout)
            replace_placeholders(slide, row)

        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
        prs.save(output_path)
        print(f"Wrote {output_path}")
    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(description="Generate PPT from template and data file")
    parser.add_argument("--data", required=True, help="Path to Excel (.xlsx) or CSV file")
    parser.add_argument("--template", default="template.pptx", help="Path to template PPTX")
    parser.add_argument("--output", default="output.pptx", help="Output PPTX path")
    parser.add_argument("--sheet", default=None, help="Excel sheet name or index (optional)")
    args = parser.parse_args()

    build_ppt(args.data, args.template, args.output, sheet=args.sheet)


if __name__ == "__main__":
    main()
