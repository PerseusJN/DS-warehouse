import glob
import os
import zipfile
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

print("=" * 60)
print(" UK NON-RESIDENTIAL EPC DYNAMIC CITY DASHBOARD")
print("=" * 60)

# Default path
FOLDER_PATH = r"C:\Users\JiwasaNawareeLowCarb\Downloads\non-domestic-csv"

# Safeguard: ask if not found
if not os.path.exists(FOLDER_PATH):
  print(f"\n⚠️ Default path not found: {FOLDER_PATH}")
  FOLDER_PATH = input(
      "👉 Please paste your folder path OR your 'non-domestic-csv.zip' file"
      " path: "
  ).strip().strip('"\'')

# Automatically extract if user provides a .zip file path
if FOLDER_PATH.endswith(".zip") and os.path.isfile(FOLDER_PATH):
  print(f"\n📦 Detected a .zip file. Extracting automatically...")
  extract_dir = FOLDER_PATH[:-4]  # Removes '.zip' from the path
  with zipfile.ZipFile(FOLDER_PATH, "r") as zip_ref:
    zip_ref.extractall(extract_dir)
  FOLDER_PATH = extract_dir
  print(f"✅ Successfully extracted to folder: {FOLDER_PATH}")

if not os.path.exists(FOLDER_PATH):
  print("❌ Error: The path provided does not exist. Please check it.")
  exit()

# Ask for the city or town right in the terminal
target_city = (
    input(
        "\n👉 Enter the City or Town you want to analyse (e.g., Exeter,"
        " Plymouth, Bristol): "
    )
    .strip()
    .upper()
)

if not target_city:
  print("❌ City name cannot be empty.")
  exit()

print(f"\n📂 Scanning all yearly certificate files for '{target_city}'...")
cert_files = glob.glob(os.path.join(FOLDER_PATH, "certificates-*.csv"))

# If not found directly, check subfolders (in case it extracted into a nested folder)
if not cert_files:
  cert_files = glob.glob(
      os.path.join(FOLDER_PATH, "**", "certificates-*.csv"), recursive=True
  )

if not cert_files:
  print(f"❌ Error: No certificate CSV files found in: {FOLDER_PATH}")
  exit()

# 1. Loop through all years and filter records for the requested city
all_city_dfs = []

for file in cert_files:
  try:
    chunk_iter = pd.read_csv(
        file, low_memory=False, chunksize=50000, on_bad_lines="skip"
    )
    for chunk in chunk_iter:
      town_col_exists = "posttown" in chunk.columns
      if town_col_exists:
        matched = chunk[
            chunk["posttown"]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.contains(target_city)
        ]
      else:
        chunk["search_scope"] = (
            chunk.astype(str).agg(" ".join, axis=1).str.upper()
        )
        matched = chunk[chunk["search_scope"].str.contains(target_city)]

      if not matched.empty:
        all_city_dfs.append(matched)
  except Exception as e:
    print(f"⚠️ Could not read {file}: {e}")

if not all_city_dfs:
  print(f"\n❌ No properties found for '{target_city}' across the dataset.")
  exit()

df = pd.concat(all_city_dfs, ignore_index=True)
print(f"\n✨ Found {len(df)} total property records for {target_city}!")

# 2. Clean Rating Bands & Calculate Expiry Status
rating_col = "asset_rating_band"
date_col = "lodgement_date"

if rating_col in df.columns:
  df[rating_col] = (
      df[rating_col].fillna("UNKNOWN").astype(str).str.upper().str.strip()
  )
else:
  df[rating_col] = "UNKNOWN"

rating_order = {
    "G": 1,
    "F": 2,
    "E": 3,
    "D": 4,
    "C": 5,
    "B": 6,
    "A": 7,
    "A+": 8,
    "UNKNOWN": 9,
}
df["rating_score"] = df[rating_col].map(rating_order).fillna(9)

# Calculate 10-year expiry rule (Current year is 2026)
if date_col in df.columns:
  df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
  df["expiry_date"] = df[date_col] + pd.DateOffset(years=10)
  current_date = pd.Timestamp("2026-09-07")
  df["certificate_status"] = df["expiry_date"].apply(
      lambda x: "EXPIRED" if pd.notnull(x) and x < current_date else "ACTIVE"
  )
else:
  df["certificate_status"] = "UNKNOWN"

# Find the correct Floor Area column name dynamically and make it numeric
floor_area_col = next(
    (c for c in ["total_useful_floor_area", "floor_area"] if c in df.columns),
    None,
)

if floor_area_col:
  df[floor_area_col] = pd.to_numeric(df[floor_area_col], errors="coerce")

# Pull only the most recent certificate for each unique place (address1 + postcode)
if "address1" in df.columns and "postcode" in df.columns and date_col in df.columns:
  df = df.sort_values(by=date_col, ascending=False)
  df = df.drop_duplicates(subset=["address1", "postcode"], keep="first")
  print(
      f"✨ Filtered to latest certificate per property. Unique properties:"
      f" {len(df)}"
  )

# 3. Sort: Worst ratings (G) first, moving up to best (A+)
df_sorted = df.sort_values(by="rating_score", ascending=True)

# Select key dashboard columns
selected_cols = [
    c
    for c in [
        "address1",
        "postcode",
        rating_col,
        "certificate_status",
        floor_area_col,
        date_col,
        "expiry_date",
        "property_type",
        "certificate_number",
    ]
    if c and c in df_sorted.columns
]

# 4. Export Master Dashboard to OneDrive with custom city name, colored headers & auto-fit columns
safe_city_name = target_city.replace(" ", "_")
output_filename = os.path.join(
    r"C:\Users\JiwasaNawareeLowCarb\OneDrive - Low Carbon Estates Limited",
    f"Dashboard_{safe_city_name}_All_Years.xlsx",
)

with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
  df_sorted[selected_cols].to_excel(writer, index=False, sheet_name="Sheet1")
  worksheet = writer.sheets["Sheet1"]

  # Style Headers: Navy blue background, white bold text, centered
  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

  for cell in worksheet[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

  # Auto-adjust column widths so content and headers fit nicely without '########'
  for col in worksheet.columns:
    max_len = 0
    col_letter = col[0].column_letter
    for cell in col:
      if cell.value is not None:
        cell_str = str(cell.value)
        if len(cell_str) > max_len:
          max_len = len(cell_str)
    worksheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

print("=" * 110)
print(
    f" TOP PROPERTIES FOR {target_city} (RANKED: G ➔ A+, MOST RECENT ONLY)"
)
print("=" * 110)
print(df_sorted[selected_cols].head(25).to_string(index=False))
print("=" * 110)
print(f"\n📊 Dashboard saved directly to your OneDrive:\n{output_filename}")
