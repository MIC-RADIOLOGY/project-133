import os
import re
from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg'}
EXCEL_TEMPLATE = 'template.xlsx'  # Your quotation template file

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def parse_ocr_text(text):
    """
    Parses OCR text into structured items.
    Expected columns: DESCRIPTION, TARRIF/TARIFF, MOD, QTY, FEES, AMOUNT
    """

    # Clean multiple spaces
    text = re.sub(r"[ ]{2,}", " ", text)

    lines = [line.strip() for line in text.split("\n") if line.strip()]
    items = []

    # Normalize headers since OCR may produce variants like "TARRIF" / "TARIFF" and " AMOUNT"
    header_regex = re.compile(
        r"DESCRIPTION\s+TAR?RIFF?\s+MOD\s+QTY\s+FEES\s+AMOUNT",
        re.IGNORECASE
    )

    header_found = False

    for line in lines:
        if not header_found and header_regex.search(line.replace(" ", "")):
            header_found = True
            continue

        if header_found:
            # Split the data row
            parts = line.split(" ")

            if len(parts) >= 6:
                description = parts[0]
                tarrif = parts[1]
                mod = parts[2]
                qty = parts[3]
                fees = parts[4]
                amount = parts[5]

                items.append({
                    "description": description,
                    "tarrif": tarrif,
                    "mod": mod,
                    "qty": qty,
                    "fees": fees,
                    "amount": amount
                })

    return items


def export_to_excel(parsed_items):
    """
    Writes parsed OCR data into the Excel template starting at:
    - A22 for description
    - A23 for details
    - Next scan increments by 2 rows
    Always writes TOTAL to G22
    """

    wb = load_workbook(EXCEL_TEMPLATE)
    ws = wb.active

    # Hard-coded row start
    base_description_row = 22
    base_details_row = 23

    current_row = 0

    for item in parsed_items:
        desc_row = base_description_row + (current_row * 2)
        detail_row = base_details_row + (current_row * 2)

        # Write description
        ws[f"A{desc_row}"] = item["description"]

        # Write details row
        ws[f"A{detail_row}"] = item["tarrif"]
        ws[f"B{detail_row}"] = item["mod"]
        ws[f"C{detail_row}"] = item["qty"]
        ws[f"D{detail_row}"] = item["fees"]
        ws[f"E{detail_row}"] = item["amount"]

        current_row += 1

    # TOTAL ALWAYS GOES IN G22 (as per user instruction)
    total = sum(float(item["amount"]) for item in parsed_items if item["amount"].replace('.', '', 1).isdigit())
    ws["G22"] = total

    output_path = os.path.join(UPLOAD_FOLDER, "Quotation_Output.xlsx")
    wb.save(output_path)

    return output_path


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if 'file' not in request.files:
            return "No file part"

        file = request.files['file']

        if file.filename == '':
            return "No selected file"

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Run OCR (replace with your OCR function)
            import pytesseract
            from PIL import Image

            text = pytesseract.image_to_string(Image.open(filepath))

            items = parse_ocr_text(text)

            if not items:
                return "No valid data detected from OCR."

            exported_file = export_to_excel(items)

            return send_file(exported_file, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
