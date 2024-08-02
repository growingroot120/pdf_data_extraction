import re
import pandas as pd
import pdfplumber

class HearingTestReport:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path

    @staticmethod
    def _create_csv(data, filename='hearing_test_history.csv'):
        df = pd.DataFrame(data)
        df.to_csv(filename, index=False)

    def extract_text(self):
        data = []

        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    lines = text.split('\n')
                    client_info = {}
                    right_ear_tests = []
                    left_ear_tests = []

                    # Extract client info
                    name_match = re.search(r"Name:\s*(.+?)\s*Client:", text)
                    if name_match:
                        client_info["Name"] = name_match.group(1).strip()

                    client_match = re.search(r"Client:\s*(.+?)\s*Test Date:", text)
                    if client_match:
                        client_info["Client"] = client_match.group(1).strip()

                    test_date_match = re.search(r"Test Date:\s*(.+?)\s*Employee ID:", text)
                    if test_date_match:
                        client_info["Test Date"] = test_date_match.group(1).strip()

                    employee_id_match = re.search(r"Employee ID:\s*(.+?)\s*Facility:", text)
                    if employee_id_match:
                        client_info["Employee ID"] = employee_id_match.group(1).strip()

                    facility_match = re.search(r"Facility:\s*(.+?)\s*Event ID:", text)
                    if facility_match:
                        client_info["Facility"] = facility_match.group(1).strip()

                    event_id_match = re.search(r"Event ID:\s*(.+?)\s*Department:", text)
                    if event_id_match:
                        client_info["Event ID"] = event_id_match.group(1).strip()

                    department_match = re.search(r"Department:\s*(.+?)\s*Birth Date:", text)
                    if department_match:
                        client_info["Department"] = department_match.group(1).strip()

                    birth_date_match = re.search(r"Birth Date:\s*(.+?)\s*Position:", text)
                    if birth_date_match:
                        client_info["Birth Date"] = birth_date_match.group(1).strip()

                    position_match = re.search(r"Position:\s*(.+?)\s*WorkShift:", text)
                    if position_match:
                        client_info["Position"] = position_match.group(1).strip()

                    workshift_match = re.search(r"WorkShift:\s*(.+?)\s*Home Address:", text)
                    if workshift_match:
                        client_info["WorkShift"] = workshift_match.group(1).strip()

                    home_address_match = re.search(r"Home Address:\s*(.+?)\s*Hire Date:", text)
                    if home_address_match:
                        client_info["Home Address"] = home_address_match.group(1).strip()

                    hire_date_match = re.search(r"Hire Date:\s*(.+?)\s*Examinetics ID:", text)
                    if hire_date_match:
                        client_info["Hire Date"] = hire_date_match.group(1).strip()

                    examinetics_id_match = re.search(r"Examinetics ID:\s*([\d]+)", text)
                    if examinetics_id_match:
                        client_info["Examinetics ID"] = examinetics_id_match.group(1).strip()

                    # Extract hearing test history
                    in_right_ear_section = False
                    in_left_ear_section = False

                    for line in lines:
                        if "Test History for Right Ear" in line:
                            in_right_ear_section = True
                            in_left_ear_section = False
                        elif "Test History for Left Ear" in line:
                            in_right_ear_section = False
                            in_left_ear_section = True

                        if in_right_ear_section or in_left_ear_section:
                            if re.match(r'\d+/\d+/\d+\s+\d+:\d+', line):
                                test_data = re.split(r'\s{1,7}', line)
                                ear = "Right" if in_right_ear_section else "Left"
                                test_entry = {
                                    "Test Date/Time": f"{test_data[0]} {test_data[1]}",
                                    "Type": test_data[2],
                                    "Age": test_data[3],
                                    "ST": test_data[4],
                                    ".5K": test_data[5],
                                    "1K": test_data[6],
                                    "2K": test_data[7],
                                    "3K": test_data[8],
                                    "4K": test_data[9],
                                    "6K": test_data[10],
                                    "8K": test_data[11],
                                    "Ear": ear
                                }
                                if ear == "Right":
                                    right_ear_tests.append(test_entry)
                                else:
                                    left_ear_tests.append(test_entry)

                    # Merge right and left ear data
                    merged_tests = {}

                    # Add right ear tests to merged_tests dictionary
                    for test in right_ear_tests:
                        merged_tests[test["Test Date/Time"]] = {
                            **test,
                            "Right .5K": test[".5K"],
                            "Right 1K": test["1K"],
                            "Right 2K": test["2K"],
                            "Right 3K": test["3K"],
                            "Right 4K": test["4K"],
                            "Right 6K": test["6K"],
                            "Right 8K": test["8K"]
                        }

                    # Add left ear tests to merged_tests dictionary
                    for test in left_ear_tests:
                        if test["Test Date/Time"] in merged_tests:
                            merged_tests[test["Test Date/Time"]].update({
                                "Left .5K": test[".5K"],
                                "Left 1K": test["1K"],
                                "Left 2K": test["2K"],
                                "Left 3K": test["3K"],
                                "Left 4K": test["4K"],
                                "Left 6K": test["6K"],
                                "Left 8K": test["8K"]
                            })
                        else:
                            merged_tests[test["Test Date/Time"]] = {
                                **test,
                                "Left .5K": test[".5K"],
                                "Left 1K": test["1K"],
                                "Left 2K": test["2K"],
                                "Left 3K": test["3K"],
                                "Left 4K": test["4K"],
                                "Left 6K": test["6K"],
                                "Left 8K": test["8K"]
                            }

                    # Convert merged_tests dictionary to a list
                    hearing_test_data = list(merged_tests.values())

                    # Attach client_info to each test entry
                    for test in hearing_test_data:
                        # Remove unnecessary keys
                        test.pop('.5K', None)
                        test.pop('1K', None)
                        test.pop('2K', None)
                        test.pop('3K', None)
                        test.pop('4K', None)
                        test.pop('6K', None)
                        test.pop('8K', None)
                        test.pop('Ear', None)
                        data.append({**client_info, **test})

        self._create_csv(data)

if __name__ == '__main__':
    print('Execution start ... ')
    report = HearingTestReport('AU0013_TestHistory_Event (1) Copy Extract[1].pdf')
    report.extract_text()
    print('Execution ends ... ')
