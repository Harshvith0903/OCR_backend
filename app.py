import re
import cv2
import pytesseract
import numpy as np
import pandas as pd
import json
import base64
from io import BytesIO
from pdf2image import convert_from_bytes

# Define regex patterns for extracting specific text
pattern = r"\b[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+-[A-Za-z0-9]+\b"
tag_pattern = r"\b[A-Z]{2,}\b"

def process_pdf(pdf_bytes):
    """Processes a PDF from bytes, extracts text, and returns an Excel file as bytes."""
    images = convert_from_bytes(pdf_bytes, dpi=300)
    extracted_data = []

    for img in images:
        # Convert PIL Image to OpenCV format
        img_np = np.array(img)
        image = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)

        # Convert to HSV
        hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)

        # Define red color range for text detection
        lower_red1, upper_red1 = np.array([0, 100, 100]), np.array([10, 255, 255])
        lower_red2, upper_red2 = np.array([170, 100, 100]), np.array([180, 255, 255])

        # Create mask for detecting red text
        mask = cv2.inRange(hsv, lower_red1, upper_red1) + cv2.inRange(hsv, lower_red2, upper_red2)
        red_text_only = cv2.bitwise_and(image, image, mask=mask)

        # Convert to grayscale and apply threshold
        gray = cv2.cvtColor(red_text_only, cv2.COLOR_BGR2GRAY)
        gray = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

        # Perform OCR
        extracted_text = pytesseract.image_to_string(gray, config="--oem 3 --psm 6")
        words = extracted_text.split()
        extracted_data.extend([word for word in words if re.search(pattern, word)])

    # Create DataFrame with extracted text
    df = pd.DataFrame(extracted_data, columns=["Extracted Text"])

    # Extract unique uppercase tags and count occurrences
    df["Tag"] = df["Extracted Text"].apply(lambda x: re.search(tag_pattern, x).group(0) if re.search(tag_pattern, x) else None)
    tag_counts = df["Tag"].value_counts().reset_index()
    tag_counts.columns = ["Tag", "Count"]

    # Save extracted data and tag counts to an in-memory Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Extracted Text", index=False)
        tag_counts.to_excel(writer, sheet_name="Tag Counts", index=False)

        # Auto-adjust column width while writing
        workbook = writer.book
        for sheet_name in ["Extracted Text", "Tag Counts"]:
            worksheet = workbook[sheet_name]
            for col in worksheet.columns:
                max_length = max((len(str(cell.value)) for cell in col if cell.value), default=0)
                worksheet.column_dimensions[col[0].column_letter].width = max_length + 2

    output.seek(0)

    # Encode the Excel file in base64
    return base64.b64encode(output.getvalue()).decode("utf-8")

def lambda_handler(event, context):
    """AWS Lambda function to process PDF file from API Gateway."""
    try:
        # Decode incoming request body
        body = json.loads(event["body"])

        # Extract PDF file from event (base64 encoded)
        pdf_base64 = body.get("pdf_base64")
        if not pdf_base64:
            return {
                "statusCode": 400,
                "body": json.dumps({"error": "No PDF file provided"})
            }

        # Decode base64 PDF
        pdf_bytes = base64.b64decode(pdf_base64)

        # Process the PDF
        excel_base64 = process_pdf(pdf_bytes)

        return {
            "statusCode": 200,
            "body": json.dumps({
                "message": "File processed successfully",
                "excel_base64": excel_base64
            }),
            "headers": {"Content-Type": "application/json"}
        }

    except Exception as e:
        return {
            "statusCode": 500,
            "body": json.dumps({"error": str(e)})
        }
