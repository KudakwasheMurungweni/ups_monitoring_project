import pyautogui
import pytesseract
from PIL import Image
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

# Replace with your DCE login details and coordinates
dce_username = 'your_username'
dce_password = 'your_password'
username_field_x, username_field_y = 150, 250
password_field_x, password_field_y = 150, 300
login_button_x, login_button_y = 150, 350

# Replace with the coordinates of the two "UPS" elements and the path to your Excel file
ups1_x, ups1_y = 100, 200
ups2_x, ups2_y = 200, 300
excel_file = "hvac_checklist.xlsx"

# Email configuration
smtp_server = 'smtp.outlook.com'
smtp_port = 587
username = 'server.automation@delta.co.zw' 
password = 'C@lling!Booth2'
to_email = 'w.mapurisa@delta.co.zw'
subject = 'Daily UPS Report'

def extract_runtime_remaining(image_path):
    try:
        img = Image.open(image_path)
        text = pytesseract.image_to_string(img)
        runtime_remaining = text.split("Runtime Remaining")[1].strip()
        return runtime_remaining
    except IndexError:
        return "Data not found"
    except Exception as e:
        print(f"Error extracting text: {e}")
        return "Extraction error"

def send_email(subject, body, to_email):
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(username, password)
        server.sendmail(username, to_email, msg.as_string())
        server.quit()
        print(f'Email sent to {to_email} at {datetime.now()}')
    except Exception as e:
        print(f'Failed to send email: {e}')

try:
    # Click on the first "UPS" element and extract data
    pyautogui.moveTo(ups1_x, ups1_y)
    pyautogui.click()
    screenshot1 = pyautogui.screenshot()
    screenshot1.save('screenshot1.png')
    runtime_remaining1 = extract_runtime_remaining('screenshot1.png')

    # Click on the second "UPS" element and extract data
    pyautogui.moveTo(ups2_x, ups2_y)
    pyautogui.click()
    screenshot2 = pyautogui.screenshot()
    screenshot2.save('screenshot2.png')
    runtime_remaining2 = extract_runtime_remaining('screenshot2.png')

    # Load the Excel workbook and select the sheet
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook['HVAC Checklist']  # Replace 'Sheet1' with your actual sheet name

    # Identify the target cells (adjust the row and column numbers as needed)
    ups1_runtime_cell = sheet['B13']
    ups2_runtime_cell = sheet['B16']

    # Update the cells with the new values
    ups1_runtime_cell.value = runtime_remaining1
    ups2_runtime_cell.value = runtime_remaining2

    # Save the updated workbook
    workbook.save(excel_file)

    print("Data updated in Excel successfully.")

    # Prepare the email body
    body = f"""
    Daily UPS Report

    UPS 1 Runtime Remaining: {runtime_remaining1}
    UPS 2 Runtime Remaining: {runtime_remaining2}

    The Excel file has been updated with the latest data.
    """

    # Send the email
    send_email(subject, body, to_email)

except Exception as e:
    print("Error:", e)
