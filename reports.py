import re

import pandas as pd
import pdfplumber


class Report:

    def __init__(self):
        self.pdf_path = 'Reports.pdf'

    @staticmethod
    def _create_csv(data: list):
        df = pd.DataFrame(data)
        df.to_csv('output.csv', index=False)

    def extract_text(self):

        tag = None

        with pdfplumber.open(self.pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_data = page.extract_text()
                if 'Compact Summary Report' in page_data.split('\n')[0]:
                    tag = page_data.split('\n')[1]

        if not tag:
            return

        data = []
        with pdfplumber.open(self.pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                page_data = page.extract_text()
                if tag in page_data.split('\n')[0]:
                    table = page.extract_table()
                    if table:
                        family_brand = [item for item in table[0] if item is not None][0]
                        id_no = [item for item in table[1] if item is not None][0].strip('IDNo:').strip()
                        sid = [item for item in table[1] if item is not None][1].strip('SID:').strip()
                        dob = [item for item in table[1] if item is not None][2].strip('BirthDate:').strip()
                        gender = [item for item in table[1] if item is not None][3].strip('Gender:').strip()
                        serial_match = re.search(r"Serial No:(\d+)", page_data)
                        serial_no = ""
                        if serial_match:
                            serial_no = serial_match.group(1)
                        for ind, line in enumerate(table):
                            filtered_line = [item for item in line if item is not None]
                            if len(filtered_line) == 18 and filtered_line[0] != 'TestDate':
                                data.append({
                                    'Family Brands': family_brand,
                                    'IDNo': id_no,
                                    'SID': sid,
                                    'DOB': dob,
                                    'Gender': gender,
                                    'Serial': serial_no,

                                    'Test Date': filtered_line[0],
                                    '500 R': filtered_line[1],
                                    '1K R': filtered_line[2],
                                    '2K R': filtered_line[3],
                                    '3K R': filtered_line[4],
                                    '4K R': filtered_line[5],
                                    '6K R': filtered_line[6],
                                    '8K R': filtered_line[7],

                                    '500 L': filtered_line[10],
                                    '1K L': filtered_line[11],
                                    '2K L': filtered_line[12],
                                    '3K L': filtered_line[13],
                                    '4K L': filtered_line[14],
                                    '6K L': filtered_line[15],
                                    '8K L': filtered_line[16]
                                })

        self._create_csv(data)


if __name__ == '__main__':
    print('Execution start ... ')
    Report().extract_text()
    print('Execution ends ... ')
