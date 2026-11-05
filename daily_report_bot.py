import pyautogui
import pyperclip
import time
from datetime import datetime
from openpyxl import Workbook
import os

# Safety
pyautogui.FAILSAFE = True

# Get current date and time
now = datetime.now()
date_str = now.strftime("%Y-%m-%d")
datetime_str = now.strftime("%Y-%m-%d %H:%M")

print(f"Today's date: {date_str}")
print(f"Current date & time: {datetime_str}")

def open_weather():
    print("Opening Chrome with Nalini G profile...")
    pyautogui.hotkey('win', 'r')
    time.sleep(1)
    pyautogui.write(r'chrome --profile-directory="Default"')
    pyautogui.press('enter')
    time.sleep(2)

    pyautogui.hotkey('ctrl', 'l')
    time.sleep(0.5)
    pyautogui.write('https://www.google.com/search?q=weather+bengaluru')
    pyautogui.press('enter')
    time.sleep(2)

def copy_temperature():
    print("Copying temperature...")
    time.sleep(1)
    # Click on the big temperature number
    pyautogui.click(400, 385, clicks=3)   # ← change to your coordinates
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)
    temperature = pyperclip.paste().strip()
    print("Copied temperature:", temperature)
    return temperature


def save_to_excel(temperature):
    filename = f"daily_report_{date_str}.xlsx"
    
    # Create a new workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Report"
    
    # Headers
    ws['A1'] = "Date & Time"
    ws['B1'] = "Weather"
    ws['C1'] = "Comment"
    
    # Data
    ws['A2'] = datetime_str
    ws['B2'] = temperature
    ws['C2'] = "Good for outdoor activities"
    
    # Save the file
    wb.save(filename)
    print(f"Excel file saved: {filename}")
    return filename

def take_screenshot(excel_file):
    print("Opening Excel and taking screenshot...")
    # Open the Excel file
    os.startfile(excel_file)
    time.sleep(5)  # wait for Excel to open
    
    # Take screenshot
    screenshot_name = f"screenshot_{date_str}.png"
    pyautogui.screenshot(screenshot_name)
    print(f"Screenshot saved: {screenshot_name}")

open_weather()
temp = copy_temperature()
excel_file = save_to_excel(temp)
take_screenshot(excel_file)
print("All done! Daily report bot finished successfully.")