import re
import pandas as pd
import pdfplumber
import pytesseract
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
class HearingTestReport:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path

    @staticmethod
    def _create_csv(data, filename='Subject Summaries FULTON BELLOWS LLC 2023 Extract[1].csv'):
        df = pd.DataFrame(data)
        df.to_csv(filename, index=False)

    def extract_text(self):
        data = []

        with pdfplumber.open(self.pdf_path) as pdf:
            is_first_match = True
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    client_info = {}
                    hearing_test_data = []

                    Company_match = re.search(r"Company:\s*(.+?)\s*Hire Date:", text)
                    if Company_match:
                        client_info["Company"] = Company_match.group(1).strip()

                    HireDate_match = re.search(r"Hire Date:\s*(.+?)\s*ID:", text)
                    if HireDate_match:
                        client_info["Hire Date"] = HireDate_match.group(1).strip()

                    ID_match = re.search(r"ID:\s*(.+?)\s*Language:", text, re.IGNORECASE)
                    if ID_match:
                        client_info["ID"] = ID_match.group(1).strip()

                    Language_match = re.search(r"Language:\s*(.+?)\s*Name:", text, re.IGNORECASE)
                    if Language_match:
                        client_info["Language"] = Language_match.group(1).strip()

                    Name_match = re.search(r"Name:\s*(.*?)\s*Location:", text, re.DOTALL)
                    if Name_match:
                        client_info["Name"] = self.extract_name_from_image(page, Name_match.group(0), page_number)

                    BirthDate_match = re.search(r"Birth Date:\s*(.+)", text)
                    if BirthDate_match:
                        client_info["Birth Date"] = BirthDate_match.group(1).strip()

                    Sex_match = re.search(r"Sex:\s*(.+)", text)
                    if Sex_match:
                        client_info["Sex"] = Sex_match.group(1).strip()

                    Status_match = re.search(r"Status:\s*(.+)", text)
                    if Status_match:
                        client_info["Status"] = Status_match.group(1).strip()

                    for line in lines:
                        # print(f"Line: {line}")
                        if re.match(r'\d+/\s?\d+/\d+\s+\d+:\d+:\d+', line):
                            # Clean the line by removing single spaces within date and time, then collapsing multiple spaces into a single space
                            cleaned_line = re.sub(r'(\d+)/\s?(\d+)/(\d+)\s+(\d+):(\d+):(\d+)', r'\1/\2/\3 \4:\5:\6', line)
                            cleaned_line = re.sub(r'\s+', ' ', cleaned_line).strip()
                            test_data = cleaned_line.split(' ')
                            
                            if is_first_match:
                                # Remove the 3rd and 4th elements if it is the first match
                                if len(test_data) > 3:
                                    del test_data[2]
                                is_first_match = False

                            if len(test_data) >= 17:
                                hearing_test_data.append({
                                    "Date": test_data[0],
                                    "Time": test_data[1],
                                    "500": test_data[2],
                                    "1K": test_data[3],
                                    "2K": test_data[4],
                                    "3K": test_data[5],
                                    "4K": test_data[6],
                                    "6K": test_data[7],
                                    "8K": test_data[8],
                                    "500 (second ear)": test_data[9],
                                    "1K (second ear)": test_data[10],
                                    "2K (second ear)": test_data[11],
                                    "3K (second ear)": test_data[12],
                                    "4K (second ear)": test_data[13],
                                    "5K": test_data[14],
                                    "8K (second ear)": test_data[15],
                                })

                    for test in hearing_test_data:
                        data.append({**client_info, **test})

        self._create_csv(data)

    def extract_name_from_image(self, page, name_line, page_number):
        text = page.extract_text()
        words = page.extract_words()
        for word in words:
            if word['text'] == "Name:":
                # Calculate the bounding box for the entire line containing "Name:"
                bbox = (word['x0'], word['top'], page.width, word['bottom'])
                cropped_page = page.within_bbox(bbox)
                image = cropped_page.to_image()
                image_path = f'Name_line_page_{page_number}.png'
                # image.save(image_path)

                # Use OCR to extract text from the saved image
                extracted_text = pytesseract.image_to_string(Image.open(image_path))
                # print(f"extracted_text: {extracted_text}")
                # Extract the information between "Name:" and "Location:"
                name_match = re.search(r"Name:\s*(.*?)\s*Location:", extracted_text, re.DOTALL)
                if name_match:
                    return name_match.group(1).strip()
                else:
                    return ""

if __name__ == '__main__':
    print('Execution start ... ')
    report = HearingTestReport('Subject Summaries FULTON BELLOWS LLC 2023 Extract[1].pdf')
    report.extract_text()
    print('Execution ends ... ')
