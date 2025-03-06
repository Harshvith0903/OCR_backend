import os
import re
import cv2
import pytesseract
import numpy as np
import pandas as pd
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS  # Allows communication with Next.js frontend
from pdf2image import convert_from_path
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)  # Enable CORS to allow frontend requests

# Define upload and output directories
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER

# Define regex patterns for extracting specific text
pattern = r"\b[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+\b"
tag_pattern = r"\b[A-Z]{2,}\b"

def process_pdf(pdf_path, output_excel_path):
    """Processes the uploaded PDF to extract red-highlighted text and save it in an Excel file."""
    images = convert_from_path(pdf_path, dpi=300)
    extracted_data = []

    for page_number, img in enumerate(images, start=1):
        image_path = os.path.join(UPLOAD_FOLDER, f"page_{page_number}.png")
        img.save(image_path, "PNG")
        
        # Read the image and convert to HSV color space
        image = cv2.imread(image_path)
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

        # Define red color range for text detection
        lower_red1, upper_red1 = np.array([0, 100, 100]), np.array([10, 255, 255])
        lower_red2, upper_red2 = np.array([170, 100, 100]), np.array([180, 255, 255])
        
        # Create masks for detecting red text
        mask = cv2.inRange(hsv, lower_red1, upper_red1) + cv2.inRange(hsv, lower_red2, upper_red2)
        red_text_only = cv2.bitwise_and(image, image, mask=mask)
        gray = cv2.cvtColor(red_text_only, cv2.COLOR_BGR2GRAY)
        gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
        
        # Perform OCR to extract text
        extracted_text = pytesseract.image_to_string(gray, config="--oem 3 --psm 6")
        words = extracted_text.split()
        for word in words:
            if re.search(pattern, word):
                extracted_data.append(word)
    
    # Create DataFrame with extracted text
    df = pd.DataFrame(extracted_data, columns=["Extracted Text"])
    
    # Extract unique uppercase tags and count occurrences
    df["Tag"] = df["Extracted Text"].apply(lambda x: re.search(tag_pattern, x).group(0) if re.search(tag_pattern, x) else None)
    tag_counts = df["Tag"].value_counts().reset_index()
    tag_counts.columns = ["Tag", "Count"]

    # Save extracted data and tag counts to Excel
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Extracted Text", index=False)
        tag_counts.to_excel(writer, sheet_name="Tag Counts", index=False)
    
    # Adjust column width in Excel for better readability
    wb = load_workbook(output_excel_path)
    ws = wb["Extracted Text"]
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(output_excel_path)

@app.route("/upload", methods=["POST"])
def upload_file():
    """Handles PDF file upload and triggers processing."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    if not file.filename.endswith(".pdf"):
        return jsonify({"error": "Invalid file format. Only PDFs are allowed"}), 400

    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)
    
    output_excel_path = os.path.join(app.config["OUTPUT_FOLDER"], "extracted_data.xlsx")
    process_pdf(file_path, output_excel_path)
    
    return jsonify({"message": "File processed successfully", "download_url": "/download"}), 200

@app.route("/download", methods=["GET"])
def download_file():
    """Provides a downloadable link for the processed Excel file."""
    output_excel_path = os.path.join(app.config["OUTPUT_FOLDER"], "extracted_data.xlsx")
    return send_file(output_excel_path, as_attachment=True)

@app.route("/tags", methods=["GET"])
def get_tags():
    """Returns extracted tags and their counts in JSON format."""
    output_excel_path = os.path.join(app.config["OUTPUT_FOLDER"], "extracted_data.xlsx")
    
    if not os.path.exists(output_excel_path):
        return jsonify({"tags": []})

    df = pd.read_excel(output_excel_path, sheet_name="Tag Counts")
    tags = df.rename(columns={"Tag": "Tag", "Count": "Count"}).to_dict(orient="records")
    
    return jsonify({"tags": tags})

if __name__ == "__main__":
    app.run(debug=True)
