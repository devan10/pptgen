#!/usr/bin/env python3
"""Generate a PPT with projects out for signature."""

import pandas as pd
import os
import subprocess

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Read Excel data
excel_file = os.path.join(ROOT, "data", "Practice Revenue Tracking 2025.xlsx")
df = pd.read_excel(excel_file, sheet_name="Detail")

# Filter for "Out for Signature" status (case-insensitive)
out_for_sig = df[df['Status'].str.contains('out for signature', case=False, na=False)].copy()

print(f"✓ Found {len(out_for_sig)} projects out for signature")

if len(out_for_sig) == 0:
    print("No projects found with 'Out for Signature' status")
    exit(1)

# Reset index for numbering
out_for_sig = out_for_sig.reset_index(drop=True)

# Convert Revenue to currency format
out_for_sig['Revenue_Formatted'] = out_for_sig['Rev'].apply(lambda x: f"${x:,.0f}" if pd.notna(x) else "")

# Create a data structure for table rows
table_data = {}
for row_idx, row in out_for_sig.iterrows():
    table_data[f"Client_{row_idx+1}"] = str(row['Client'])
    table_data[f"Project_Name_{row_idx+1}"] = str(row['Project Name'])
    table_data[f"Days_to_Sign_{row_idx+1}"] = str(int(row['Days to Sign'])) if pd.notna(row['Days to Sign']) else ""
    table_data[f"Revenue_{row_idx+1}"] = row['Revenue_Formatted']

# Write as CSV with one row of data
csv_file = os.path.join(ROOT, "data", "out_for_sig_data.csv")
columns = list(table_data.keys())
with open(csv_file, 'w') as f:
    f.write(','.join(columns) + '\n')
    f.write(','.join([f'"{table_data[col]}"' for col in columns]) + '\n')

print(f"✓ Created data CSV: {csv_file}")

# Generate PPT using the generator
template = os.path.join(ROOT, "templates", "2026_Insight_PPT_Template_OutForSig.potx")
output = os.path.join(ROOT, "out", "out_for_signature.pptx")

cmd = [
    "python",
    os.path.join(ROOT, "src", "generate_ppt.py"),
    "--data", csv_file,
    "--template", template,
    "--output", output
]

result = subprocess.run(cmd, capture_output=True, text=True)
if result.returncode != 0:
    print("STDERR:", result.stderr)
    exit(1)

file_size = os.path.getsize(output)
print(f"✓ Out for Signature PPT created: {output} ({file_size:,} bytes)")
