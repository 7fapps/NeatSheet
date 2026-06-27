"""
app.py
------
NeatSheet Web App.
Lets a user upload a CSV/Excel file in their browser,
cleans it using cleaner.py, and lets them download the result.

Run with:  python app.py
Then open: http://127.0.0.1:5000 in your browser.
"""

import os
import uuid
import pandas as pd
from flask import Flask, request, render_template, send_file, flash, redirect, url_for

from cleaner import clean_dataframe, save_clean_excel, extract_table_from_docx, extract_table_from_image

app = Flask(__name__)
app.secret_key = "neatsheet-dev-secret"  # only used for flash messages, fine for local use

# Folders for temporary file storage
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ALLOWED_EXTENSIONS = {"csv", "xlsx", "xls", "docx", "jpg", "jpeg", "png"}

# Make sure these folders exist (create them if missing)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def allowed_file(filename: str) -> bool:
    """Checks the uploaded file has an extension we support."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET"])
def index():
    """Shows the upload page."""
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """Handles the uploaded file: cleans it, then sends back the result."""

    # --- Step 1: Check a file was actually sent ---
    if "file" not in request.files:
        flash("No file was selected.")
        return redirect("/")

    file = request.files["file"]

    if file.filename == "":
        flash("No file was selected.")
        return redirect("/")

    if not allowed_file(file.filename):
        flash("Unsupported file type. Please upload a CSV, Excel, Word, or photo file (.csv, .xlsx, .xls, .docx, .jpg, .png).")
        return redirect("/")

    # --- Step 2: Save the uploaded file temporarily ---
    # We use a random ID in the filename so multiple users uploading
    # at the same time don't overwrite each other's files.
    file_id = uuid.uuid4().hex
    extension = file.filename.rsplit(".", 1)[1].lower()
    upload_path = os.path.join(UPLOAD_FOLDER, f"{file_id}.{extension}")
    file.save(upload_path)

    # unparsed_lines only gets filled in when processing a photo (OCR).
    # For all other file types it stays empty.
    unparsed_lines = []

    # --- Step 3: Load it into pandas (method depends on file type) ---
    try:
        if extension == "csv":
            df = pd.read_csv(upload_path)
        elif extension in ("xlsx", "xls"):
            df = pd.read_excel(upload_path)
        elif extension == "docx":
            df = extract_table_from_docx(upload_path)
        elif extension in ("jpg", "jpeg", "png"):
            df, unparsed_lines = extract_table_from_image(upload_path)
        else:
            # This shouldn't happen because allowed_file() already checked,
            # but it's good practice to handle unexpected cases anyway.
            flash("Unsupported file type.")
            return redirect("/")
    except ValueError as e:
        # Raised by extract_table_from_docx / extract_table_from_image
        # when no usable data is found
        flash(str(e))
        return redirect("/")
    except Exception as e:
        flash(f"Could not read the file: {e}")
        return redirect("/")

    if df.empty:
        flash("The uploaded file appears to be empty.")
        return redirect("/")

    # --- Step 4: Clean it ---
    cleaned_df = clean_dataframe(df)

    # --- Step 5: Save the cleaned result as an Excel file ---
    output_path = os.path.join(OUTPUT_FOLDER, f"{file_id}_NeatSheet_Output.xlsx")
    save_clean_excel(cleaned_df, output_path, unparsed_lines=unparsed_lines)

    # --- Step 6: Clean up the uploaded file (we don't need it anymore) ---
    os.remove(upload_path)

    # --- Step 7: Show a preview page instead of downloading immediately ---
    # We convert the cleaned DataFrame to a simple HTML table so the user
    # can see exactly what NeatSheet did before they download anything.
    preview_html = cleaned_df.to_html(
        classes="preview-table",
        index=False,
        border=0,
        na_rep=""  # show blank instead of "NaN" for empty cells
    )

    return render_template(
        "preview.html",
        table_html=preview_html,
        row_count=len(cleaned_df),
        unparsed_lines=unparsed_lines,
        file_id=file_id
    )


@app.route("/download/<file_id>")
def download(file_id):
    """
    Serves the previously cleaned Excel file for download.
    file_id is the random ID generated during /upload, used here to
    find the matching file in the outputs folder.
    """
    output_path = os.path.join(OUTPUT_FOLDER, f"{file_id}_NeatSheet_Output.xlsx")

    if not os.path.exists(output_path):
        flash("That file is no longer available. Please upload again.")
        return redirect("/")

    return send_file(
        output_path,
        as_attachment=True,
        download_name="NeatSheet_Output.xlsx"
    )


if __name__ == "__main__":
    # debug=True auto-reloads the server when you save code changes.
    # Turn this off (debug=False) before sharing this with anyone else.
    app.run(debug=True)
