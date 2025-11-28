import pandas as pd
import requests
from PIL import Image
from io import BytesIO
import zipfile
import base64
import json
import uuid

def handler(request):
    try:
        # Read Excel file
        file = request.files.get("excel_file")
        format_choice = request.form.get("format_choice")

        if not file:
            return {
                "statusCode": 400,
                "body": json.dumps({"error": "Excel file missing"})
            }

        df = pd.read_excel(file)
        if "name" not in df.columns or "link" not in df.columns:
            return {
                "statusCode": 400,
                "body": json.dumps({"error": "Excel must have name & link columns"})
            }

        # Create an in-memory ZIP
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:

            for _, row in df.iterrows():
                name = str(row["name"]).strip()
                url = str(row["link"]).strip()

                try:
                    r = requests.get(url, timeout=10)
                    r.raise_for_status()
                    img = Image.open(BytesIO(r.content))

                    if format_choice == "jpg":
                        img = img.convert("RGB")
                        file_bytes = BytesIO()
                        img.save(file_bytes, "JPEG", quality=95)
                        zipf.writestr(f"{name}.jpg", file_bytes.getvalue())

                    elif format_choice == "png":
                        file_bytes = BytesIO()
                        img.save(file_bytes, "PNG")
                        zipf.writestr(f"{name}.png", file_bytes.getvalue())

                except Exception as e:
                    print("Error:", e)
                    continue

        zip_buffer.seek(0)

        # Encode ZIP for Vercel
        encoded_zip = base64.b64encode(zip_buffer.read()).decode()

        return {
            "statusCode": 200,
            "headers": {
                "Content-Type": "application/zip",
                "Content-Disposition": "attachment; filename=converted.zip",
                "Content-Transfer-Encoding": "base64"
            },
            "body": encoded_zip,
            "isBase64Encoded": True
        }

    except Exception as e:
        return {
            "statusCode": 500,
            "body": json.dumps({"error": str(e)})
        }
