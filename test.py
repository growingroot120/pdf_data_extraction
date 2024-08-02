import pdfplumber
from PIL import Image
import pytesseract
import os
import io

def extract_names_from_pdf(pdf_path, output_folder):
    # Open the PDF file with pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        names = []

        # Iterate through each page
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text:
                lines = text.split('\n')
                for line_number, line in enumerate(lines):
                    if "Name:" in line:
                        # Extract the bounding box of the line containing "Name:"
                        bbox = None
                        for word in page.extract_words():
                            if "Name:" in word['text']:
                                bbox = word['x0'], word['top'], word['x1'], word['bottom']
                                break
                        
                        if bbox:
                            # Adjust the bounding box to capture the entire line
                            bbox = (bbox[0], bbox[1] - 5, page.width, bbox[3] + 5)

                            # Convert the page to an image
                            page_image = page.to_image(resolution=300)
                            image = Image.open(io.BytesIO(page_image.original))
                            cropped_image = image.crop(bbox)
                            temp_image_path = os.path.join(output_folder, f"temp_name_page{page_number}_line{line_number + 1}.png")
                            cropped_image.save(temp_image_path)

                            # Use OCR to extract text from the cropped image
                            extracted_text = pytesseract.image_to_string(cropped_image)
                            names.append(extracted_text.strip())
    
    return names

# Usage example
pdf_path = "Subject Summaries FULTON BELLOWS LLC 2023 Extract[1].pdf"
output_folder = "D:/Working/freelancer/2024_05_20_textextraction/extract"
os.makedirs(output_folder, exist_ok=True)
extracted_names = extract_names_from_pdf(pdf_path, output_folder)
for name in extracted_names:
    print(name)
