#!/usr/bin/env python3
"""Generate summary statistics from the revenue data."""

import pandas as pd
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data", "Practice Revenue Tracking 2025.xlsx")

df = pd.read_excel(DATA, sheet_name='Detail')

# Clean numeric columns
df['Rev'] = pd.to_numeric(df['Rev'], errors='coerce')
df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce')
df['GP $'] = pd.to_numeric(df['GP $'], errors='coerce')
df['GP %'] = pd.to_numeric(df['GP %'], errors='coerce')

# Calculate summary stats
summary = {
    'Total Revenue': f"${df['Rev'].sum():,.2f}",
    'Total Hours': f"{df['Hours'].sum():,.0f}",
    'Total Projects': len(df),
    'Total Clients': df['Client'].nunique(),
    'Average Revenue per Project': f"${df['Rev'].mean():,.2f}",
    'Average Hours per Project': f"{df['Hours'].mean():,.1f}",
    'Total GP $': f"${df['GP $'].sum():,.2f}",
    'Average GP %': f"{df['GP %'].mean():.1%}",
    'Top Client': df.groupby('Client')['Rev'].sum().idxmax(),
    'Top Client Revenue': f"${df.groupby('Client')['Rev'].sum().max():,.2f}",
    'Won Projects': len(df[df['Status'] == 'Won']),
    'Status Breakdown': df['Status'].value_counts().to_dict(),
}

for key, val in summary.items():
    print(f"{key}: {val}")
