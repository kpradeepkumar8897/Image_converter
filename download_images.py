import openpyxl
import os
import pandas as pd
import requests
from PIL import Image
from io import BytesIO
from zipfile import ZipFile

# Load the Excel file
excel_path = r"C:\Users\pkoppula\Desktop\Python images code\images.xlsx"  # make sure this file is in the same folder as the script
df = pd.read_excel(excel_path)

# Create folder to store images
output_dir = "downloaded_images"
os.makedirs(output_dir, exist_ok=True)

# Get image columns (excluding ASIN)
image_columns = df.columns[1:]

downloaded_files = []

for _, row in df.iterrows():
    asin = row["ASIN"]
    for col in image_columns:
        url = row[col]
        if pd.notna(url) and isinstance(url, str) and url.startswith("http"):
            try:
                response = requests.get(url, timeout=10)
                response.raise_for_status()
                img = Image.open(BytesIO(response.content)).convert("RGB")
                filename = f"{asin}.{col}.jpg"
                filepath = os.path.join(output_dir, filename)
                img.save(filepath, format="JPEG")
                downloaded_files.append(filepath)
                print(f"Downloaded: {filename}")
            except Exception as e:
                print(f"Failed to download {url}: {e}")

# Zip the images
zip_filename = "ASIN_Images.zip"
with ZipFile(zip_filename, 'w') as zipf:
    for file in downloaded_files:
        zipf.write(file, arcname=os.path.basename(file))

print(f"\nAll images zipped into {zip_filename}")
