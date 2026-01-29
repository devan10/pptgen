#!/usr/bin/env python3
"""Generate a summary PPT from Excel data."""

import subprocess
import os
import sys
import pandas as pd
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data", "Practice Revenue Tracking 2025.xlsx")
TEMPLATE = os.path.join(ROOT, "templates", "Revenue_Summary_Template.potx")
OUTPUT_DIR = os.path.join(ROOT, "out")
OUTPUT = os.path.join(OUTPUT_DIR, "revenue_summary.pptx")
GENERATOR = os.path.join(ROOT, "src", "generate_ppt.py")

# Create summary data
df = pd.read_excel(DATA, sheet_name='Detail')
df['Rev'] = pd.to_numeric(df['Rev'], errors='coerce')
df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')
df['GP $'] = pd.to_numeric(df['GP $'], errors='coerce')
df['GP %'] = pd.to_numeric(df['GP %'], errors='coerce')

summary_df = pd.DataFrame([{
    'Total Revenue': f"${df['Rev'].sum():,.2f}",
    'Total Clients': str(df['Client'].nunique()),
    'Total Projects': str(len(df)),
    'Won Projects': str(len(df[df['Status'] == 'Won'])),
    'Average Revenue per Project': f"${df['Rev'].mean():,.2f}",
    'Total GP $': f"${df['GP $'].sum():,.2f}",
    'Average GP %': f"{df['GP %'].mean():.1%}",
    'Top Client': df.groupby('Client')['Rev'].sum().idxmax(),
    'Top Client Revenue': f"${df.groupby('Client')['Rev'].sum().max():,.2f}",
}])

# Write summary to temporary CSV
summary_csv = os.path.join(tempfile.gettempdir(), 'summary.csv')
summary_df.to_csv(summary_csv, index=False)

os.makedirs(OUTPUT_DIR, exist_ok=True)

cmd = ["python", GENERATOR, "--data", summary_csv, "--template", TEMPLATE, "--output", OUTPUT]
print("Running:", " ".join(cmd))
proc = subprocess.run(cmd, capture_output=True, text=True)
print(proc.stdout)
if proc.stderr:
    print("STDERR:", proc.stderr)

if os.path.exists(OUTPUT):
    file_size = os.path.getsize(OUTPUT)
    print(f"\n✓ Summary PPT created: {OUTPUT} ({file_size:,} bytes)")
else:
    print("\n✗ Failed to create summary PPT")
    sys.exit(1)

# Cleanup
try:
    os.remove(summary_csv)
except:
    pass

