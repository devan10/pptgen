#!/usr/bin/env python3
"""Quick test for the 2026 Insight PPT template."""

import subprocess
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data", "sample_data.csv")
TEMPLATE = os.path.join(ROOT, "templates", "2026_Insight_PPT_Template.potx")
OUTPUT_DIR = os.path.join(ROOT, "out")
OUTPUT = os.path.join(OUTPUT_DIR, "insight_test_output.pptx")
GENERATOR = os.path.join(ROOT, "src", "generate_ppt.py")

if not os.path.exists(TEMPLATE):
    print(f"Template not found: {TEMPLATE}")
    sys.exit(1)

os.makedirs(OUTPUT_DIR, exist_ok=True)

cmd = ["python", GENERATOR, "--data", DATA, "--template", TEMPLATE, "--output", OUTPUT]
print("Running:", " ".join(cmd))
proc = subprocess.run(cmd, capture_output=True, text=True)
print(proc.stdout)
if proc.stderr:
    print("STDERR:", proc.stderr)

if os.path.exists(OUTPUT):
    print(f"\n✓ Test succeeded: {OUTPUT}")
else:
    print("\n✗ Test failed, output not found")
    sys.exit(1)
