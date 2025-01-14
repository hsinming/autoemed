import pandas as pd
from helium import *


def login_to_emedical(user_id, password):
    """
    Opens Chrome browser, navigates to the eMedical login page, and logs in with the provided credentials.

    Parameters:
    user_id (str): The user ID for login.
    password (str): The password for login.
    """
    # Start Chrome and navigate to the eMedical login page
    start_chrome('https://www.emedical.immi.gov.au/eMedUI/eMedical')

    # Enter the user ID and password
    write(user_id, into='User id')
    write(password, into='Password')

    # Click the Logon button
    click('Logon')

    # Optional: Add a wait or check to confirm successful login
    wait_until(Text('Logout').exists, timeout_secs=10)


def extract_emedical_numbers(file_path):
    """
    Extracts all eMedical No. from an Excel file and returns them as a list.

    Parameters:
    file_path (str): The path to the Excel file.

    Returns:
    List[str]: A list of all eMedical No. found in the file.
    """
    # Load the Excel file
    excel_content = pd.ExcelFile(file_path)
    all_emedical_numbers = []

    # Iterate through each sheet to extract eMedical No.
    for sheet in excel_content.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet)
        # Extract columns that potentially contain eMedical No.
        for col in df.columns:
            if df[col].dtype == object:  # Check if the column contains strings
                id_type = r'(HAP \d+|TRN \w+|NZER \w+|NZHR \w+|IME \w+|UMI \w+|UCI \w+|CEACBC \w+)'
                emedical_numbers = df[col].dropna().str.extract(id_type)[0]
                emedical_numbers = emedical_numbers.dropna().tolist()
                all_emedical_numbers.extend(emedical_numbers)

    return all_emedical_numbers




if __name__ == "__main__":
    # filepath = "test.xlsx"
    # output = extract_emedical_numbers(filepath)
    # print(output)

    login_to_emedical('e26748', 'Year*2025')
