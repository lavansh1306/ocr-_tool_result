import os
import re
import pytesseract
import pandas as pd
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance, ImageFilter

# Set Paths
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
poppler_path = r"C:\Users\Lavansh\Desktop\Personal Stuff\Projects\ocr_tool\poppler-24.08.0\Library\bin"

# 🔍 Common Fixes for Subject Codes
subject_corrections = {
    "ZiIMABIOIT": "21MAB101T",
    "ZiCYBIOL": "21CYB101L",
    "ZICSS1O1": "21CSS101",
    "2ZIGNHIO1": "21IGNH101",
    "2iMESi01L": "21MES101L",
    "ZIGNMIO4L": "21GNM104L",
    "21LEH1O4T": "21LEH104T"
}

# 🔍 Grade Fixes
grade_corrections = {
    "oO": "O", 
    "10)": "10", 
    "At": "A+", 
    "Oo": "O", 
    "12)": "12",
    "©": "O"  # Common OCR mistake
}

def preprocess_image(image):
    """Enhance image for better OCR accuracy."""
    image = image.convert("L")  # Convert to grayscale
    image = image.filter(ImageFilter.SHARPEN)  # Sharpen text
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)  # Increase contrast
    image = image.point(lambda x: 0 if x < 150 else 255)  # Convert to binary (B/W)
    image = image.resize((image.width * 2, image.height * 2), Image.Resampling.LANCZOS)  # Resize for better OCR
    return image

def pdf_to_images(pdf_path):
    """Convert PDF pages to images."""
    return convert_from_path(pdf_path, poppler_path=poppler_path)

def extract_text_from_image(image):
    """Extract text from an image using OCR with better settings."""
    image = preprocess_image(image)
    custom_config = r"--oem 3 --psm 6"  # High accuracy settings
    return pytesseract.image_to_string(image, config=custom_config)

def extract_reg_no(text):
    """Extract Registration Number (RA...) from OCR text."""
    match = re.search(r"(RA\d{9})", text)  # Find RA followed by 9 digits
    return match.group(1) if match else "UNKNOWN_REG_NO"  # If not found, return default

def extract_sgpa_cgpa(text):
    """Extract SGPA and CGPA using regex."""
    sgpa_match = re.search(r"SGPA\s+([\d.]+)", text)
    cgpa_match = re.search(r"CGPA\s+([\d.]+)", text)
    sgpa = float(sgpa_match.group(1)) if sgpa_match else 0.0
    cgpa = float(cgpa_match.group(1)) if cgpa_match else 0.0
    return sgpa, cgpa

def extract_subjects(text):
    """Extract subject details and fix common OCR errors."""
    subjects = []
    lines = text.split("\n")

    for line in lines:
        parts = line.split()
        if len(parts) >= 5:
            code = parts[2]  # Subject Code
            if code in subject_corrections:
                code = subject_corrections[code]  # Fix incorrect codes
            
            grade = parts[-1]
            if grade in grade_corrections:
                grade = grade_corrections[grade]  # Fix OCR-grade mistakes

            subjects.append({
                "Semester": parts[0],
                "Month/Year": parts[1],
                "Code": code,
                "Description": " ".join(parts[3:-2]),  # Subject name
                "Credit": parts[-2],
                "Grade": grade
            })
    
    return subjects

def save_to_excel(data, reg_no, sgpa, cgpa, output_folder="results"):
    """Save extracted data to an Excel file."""
    os.makedirs(output_folder, exist_ok=True)
    
    df = pd.DataFrame(data)
    df["SGPA"] = sgpa  # Add SGPA column
    df["CGPA"] = cgpa  # Add CGPA column
    
    output_file = os.path.join(output_folder, f"{reg_no}_result.xlsx")
    df.to_excel(output_file, index=False, engine="openpyxl")
    
    print(f"✅ Results saved as: {output_file}")

def main():
    pdf_path = input("Enter the path of the PDF file: ").strip()

    if not os.path.exists(pdf_path):
        print("❌ Error: File not found!")
        return

    print(f"Processing: {pdf_path}")
    
    images = pdf_to_images(pdf_path)
    full_text = "\n".join(extract_text_from_image(img) for img in images)

    # 🔍 Debugging Step: Print OCR output
    print("\n🔎 OCR Extracted Text:\n", full_text)

    reg_no = extract_reg_no(full_text)
    sgpa, cgpa = extract_sgpa_cgpa(full_text)
    subjects = extract_subjects(full_text)

    if subjects:
        save_to_excel(subjects, reg_no, sgpa, cgpa)
    else:
        print("❌ Error: Failed to extract some data. Check OCR accuracy.")

if __name__ == "__main__":
    main()
