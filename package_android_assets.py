"""
Script to prepare and package assets for Android APK building.
Creates an offline-ready bundle with embedded data.
"""
import os, shutil

ROOT = os.path.dirname(os.path.abspath(__file__))
ANDROID_DIR = os.path.join(ROOT, "pulmo_android")
ASSETS_DIR = os.path.join(ANDROID_DIR, "app", "src", "main", "assets")
os.makedirs(ASSETS_DIR, exist_ok=True)

print("Preparing Android offline asset bundle...")

# 1. Copy web static files
static_dir = os.path.join(ROOT, "pulmo_app", "static")
for f in os.listdir(static_dir):
    s = os.path.join(static_dir, f)
    d = os.path.join(ASSETS_DIR, f)
    if os.path.isfile(s):
        shutil.copy2(s, d)
        print(f"  Copied {f}")

# 2. Copy JSON datasets into assets/api/ so WebView can fetch them offline directly
api_assets_dir = os.path.join(ASSETS_DIR, "api")
os.makedirs(api_assets_dir, exist_ok=True)

data_dir = os.path.join(ROOT, "pulmo_app", "data")
name_map = {
    "radiology_cases.json": "radiology",
    "mcq_bank.json": "mcqs",
    "chapters_summary.json": "chapters",
    "crofton_encyclopedia.json": "crofton",
    "murray_compendium.json": "murray",
    "felson_roentgen_signs.json": "felson",
    "felson_full_book.json": "felson-full",
    "harrison_clinical_algorithms.json": "harrison",
    "harrison_22e_full.json": "harrison-full",
    "fishman_2026_full.json": "fishman",
    "clinical_examination_guide.json": "clinical-exam",
    "order_sheets.json": "orders",
    "pharmacology.json": "pharmacology",
    "viva_stations.json": "viva"
}

for json_file, api_route in name_map.items():
    src_path = os.path.join(data_dir, json_file)
    if os.path.exists(src_path):
        dst_path = os.path.join(api_assets_dir, api_route)
        shutil.copy2(src_path, dst_path)
        # Also copy with .json extension
        shutil.copy2(src_path, dst_path + ".json")
        print(f"  Packaged {json_file} -> assets/api/{api_route}")

# 3. Copy image assets
for folder in ["thoracic_atlas_assets", "asthma_assets", "pneumonia_figures", "bronchiectasis_figures"]:
    src_folder = os.path.join(ROOT, folder)
    dst_folder = os.path.join(ASSETS_DIR, "assets", folder)
    if os.path.exists(src_folder):
        os.makedirs(dst_folder, exist_ok=True)
        for img in os.listdir(src_folder):
            s = os.path.join(src_folder, img)
            d = os.path.join(dst_folder, img)
            if os.path.isfile(s):
                shutil.copy2(s, d)
        print(f"  Packaged images from {folder}")

print("\nAndroid offline asset bundle generated successfully in pulmo_android/app/src/main/assets/!")
