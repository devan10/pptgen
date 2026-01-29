#!/usr/bin/env python3
"""Test Revenue Detail template with revenue data."""

import subprocess
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data", "Practice Revenue Tracking 2025.xlsx")
TEMPLATE = os.path.join(ROOT, "templates", "Revenue_Detail_Template.potx")
OUTPUT_DIR = os.path.join(ROOT, "out")
OUTPUT = os.path.join(OUTPUT_DIR, "revenue_detail_output.pptx")
GENERATOR = os.path.join(ROOT, "src", "generate_ppt.py")

os.makedirs(OUTPUT_DIR, exist_ok=True)

cmd = ["python", GENERATOR, "--data", DATA, "--template", TEMPLATE, "--output", OUTPUT, "--sheet", "Detail"]
print("Running:", " ".join(cmd))
proc = subprocess.run(cmd, capture_output=True, text=True)
print(proc.stdout)
if proc.stderr:
    print("STDERR:", proc.stderr)

if os.path.exists(OUTPUT):
    file_size = os.path.getsize(OUTPUT)
    print(f"\n✓ Test succeeded: {OUTPUT} ({file_size:,} bytes)")
else:
    print("\n✗ Test failed")
    sys.exit(1)
