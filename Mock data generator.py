import pandas as pd
import numpy as np
import os
from datetime import datetime

# =====================================================================
# SYNTHETIC DATA CONFIGURATION
# =====================================================================
TOTAL_ROWS = 10000
OUTPUT_DIR = r"U:\python\poster_assets"  # Change this to your preferred local directory
FILE_NAME = f"Mock_QC_Dataset_Poster_{datetime.now().strftime('%Y%m%d')}.xlsx"

if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# Define the PT Rounds you want to simulate
pt_rounds = ['PT round A', 'PT round B', 'PT round C', 'PT round D', 'PT round E']

# Define representative analytes across your 5 chemical classes
analytes_config = {
    # PCBs (Standard ICES-6 indicators + PCB 118)
    "PCB 28": {"mean": 5.0, "sd": 0.5},
    "PCB 52": {"mean": 6.5, "sd": 0.6},
    "PCB 101": {"mean": 10.2, "sd": 1.0},
    "PCB 118": {"mean": 8.5, "sd": 0.8},
    "PCB 138": {"mean": 14.0, "sd": 1.4},
    "PCB 153": {"mean": 12.0, "sd": 1.2},
    "PCB 180": {"mean": 9.5, "sd": 0.9},

    # Dioxins & Furans
    "2,3,7,8-TCDD": {"mean": 1.5, "sd": 0.15},
    "1,2,3,7,8-PeCDD": {"mean": 2.8, "sd": 0.25},
    "1,2,3,4,7,8-HxCDF": {"mean": 3.2, "sd": 0.3},
    "2,3,4,7,8-PeCDF": {"mean": 4.1, "sd": 0.4},
    "OCDD": {"mean": 45.0, "sd": 4.5},
    "OCDF": {"mean": 18.5, "sd": 1.8},

    # PBDEs
    "BDE-47": {"mean": 6.8, "sd": 0.6},
    "BDE-99": {"mean": 7.2, "sd": 0.7},
    "BDE-100": {"mean": 3.5, "sd": 0.35},
    "BDE-153": {"mean": 2.2, "sd": 0.2},
    "BDE-154": {"mean": 1.8, "sd": 0.18},
    "BDE-209": {"mean": 25.0, "sd": 3.0},

    # PAHs (Formatted to match your extraction script exactly)
    "benzo[a]anthracene": {"mean": 3.5, "sd": 0.35},
    "Chrysene": {"mean": 4.5, "sd": 0.4},
    "benzo[b]fluoranthene": {"mean": 5.2, "sd": 0.5},
    "benzo[a]pyrene": {"mean": 2.0, "sd": 0.2},
    "indeno[1,2,3-cd]pyrene": {"mean": 1.8, "sd": 0.18},
    "benzo[g,h,i]perylene": {"mean": 2.5, "sd": 0.25},
    "PAH4 (sum)": {"mean": 15.2, "sd": 1.5},

    # PFAS (Including linear/branched splits)
    "L-PFHxS": {"mean": 0.45, "sd": 0.05},
    "Br-PFHxS": {"mean": 0.15, "sd": 0.02},
    "PFOA": {"mean": 0.50, "sd": 0.04},
    "PFNA": {"mean": 0.35, "sd": 0.03},
    "L-PFOS": {"mean": 0.85, "sd": 0.08},
    "Sum of 4 PFAS": {"mean": 2.5, "sd": 0.2}
}

print(f"🚀 INITIATING SYNTHETIC DATA GENERATION ({TOTAL_ROWS} Batches)...")

# =====================================================================
# DATA GENERATION ENGINE
# =====================================================================
# 1. Generate Batch IDs and assign random PT Rounds
batch_ids = [f"MOCK-{str(i).zfill(5)}" for i in range(1, TOTAL_ROWS + 1)]
random_pts = np.random.choice(pt_rounds, TOTAL_ROWS)

data = {
    "Batch ID": batch_ids,
    "PT Round": random_pts
}

# 2. Generate normally distributed analytical data with injected outliers
for analyte, params in analytes_config.items():
    mean = params["mean"]
    sd = params["sd"]
    
    # Base distribution (healthy laboratory data)
    values = np.random.normal(loc=mean, scale=sd, size=TOTAL_ROWS)
    
    # Inject ~5% Action level outliers (Spikes > 3 SD)
    action_indices = np.random.choice(TOTAL_ROWS, int(TOTAL_ROWS * 0.05), replace=False)
    values[action_indices] = values[action_indices] + np.random.choice([4 * sd, -4 * sd], len(action_indices))
    
    # Inject ~10% Warning level outliers (Spikes > 2 SD but < 3 SD)
    warning_indices = np.random.choice(TOTAL_ROWS, int(TOTAL_ROWS * 0.10), replace=False)
    values[warning_indices] = values[warning_indices] + np.random.choice([2.5 * sd, -2.5 * sd], len(warning_indices))
    
    # Ensure no negative concentrations
    values = np.clip(values, a_min=0.01, a_max=None)
    
    data[analyte] = np.round(values, 3)

# =====================================================================
# EXPORT & FORMATTING
# =====================================================================
df_mock = pd.DataFrame(data)

output_path = os.path.join(OUTPUT_DIR, FILE_NAME)
print(f"💾 Saving large dataset to {output_path}...")

# Use pandas to write, openpyxl for formatting
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_mock.to_excel(writer, index=False, sheet_name='QC_Data')
    
    # Auto-adjust column widths for a professional look
    worksheet = writer.sheets['QC_Data']
    for idx, col in enumerate(df_mock.columns, 1):
        max_len = max(df_mock[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.column_dimensions[worksheet.cell(row=1, column=idx).column_letter].width = min(max_len, 20)

print("🎉 SYNTHETIC MASTER DATASET COMPLETE. Ready for Tkinter Dashboard injection.")
