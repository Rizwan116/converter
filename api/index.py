from flask import Flask, request, send_file, render_template
import pandas as pd
import requests
from PIL import Image
from io import BytesIO
import tempfile
import zipfile
import os
from vercel_flask import VercelFlask

app = VercelFlask(__name__, template_folder="../templates")


def convert_image(name, url, format_choice, folder):
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()

        img = Image.open(BytesIO(response.content))

        if format_choice == "jpg":
            img = img.convert("RGB")

        output_path = os.path.join(folder, f"{name}.{format_choice}")
        img.save(output_path, format_choice.upper())

        return output_path

    except Exception:
        return None


@app.route("/", methods=["GET"])
def form():
    return render_template("index.html")


@app.route("/", methods=["POST"])
def convert():
    excel_file = request.files.get("excel_file")
    format_choice = request.form.get("format_choice")

    if not excel_file:
        return "Please upload an Excel file."

    df = pd.read_excel(excel_file)

    if "name" not in df.columns or "link" not in df.columns:
        return "Excel must include 'name' and 'link' columns."

    with tempfile.TemporaryDirectory() as tempdir:
        zip_path = os.path.join(tempdir, "converted.zip")

        with zipfile.ZipFile(zip_path, "w") as zf:
            for _, row in df.iterrows():
                name = str(row["name"]).strip()
                url = str(row["link"]).strip()
                out = convert_image(name, url, format_choice, tempdir)
                if out:
                    zf.write(out, os.path.basename(out))

        return send_file(zip_path, as_attachment=True, download_name="converted.zip")


# Export for vercel
app = app
