from datetime import datetime
from pathlib import Path
import time

import pyautogui
import pyperclip


WEATHER_URL = (
    "https://www.accuweather.com/en/in/tirunelveli/190785/"
    "weather-forecast/190785"
)
OUTPUT_FOLDER = Path("daily_reports")


def open_application(application_name):
    # Open an application through the Windows Start menu.
    pyautogui.press("win")
    pyautogui.write(application_name, interval=0.08)
    pyautogui.press("enter")
    time.sleep(5)


def copy_weather_data():
    # Open Chrome and create a new tab for the weather website.
    open_application("Google Chrome")
    pyautogui.hotkey("ctrl", "t")
    pyautogui.hotkey("ctrl", "l")
    pyautogui.write(WEATHER_URL, interval=0.01)
    pyautogui.press("enter")
    time.sleep(10)

    # Select and copy the visible weather information from the webpage.
    pyperclip.copy("")
    pyautogui.hotkey("ctrl", "a")
    pyautogui.hotkey("ctrl", "c")
    time.sleep(1)
    weather_data = pyperclip.paste().strip()

    # Stop if Chrome did not copy any webpage text.
    if not weather_data:
        raise RuntimeError("No weather data was copied from Chrome.")

    # Keep the copied webpage text in one Excel cell.
    weather_data = " ".join(weather_data.split())

    # Close the weather tab after copying its data.
    pyautogui.hotkey("ctrl", "w")
    time.sleep(1)
    return weather_data


def create_excel_report(weather_data, report_path):
    # Open Microsoft Excel through the Windows Start menu.
    open_application("Excel")
    pyautogui.hotkey("ctrl", "n")
    time.sleep(3)
    pyautogui.hotkey("ctrl", "home")

    # Create the report values with an automatically generated date and time.
    report_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Paste one value into the selected Excel cell at a time.
    def paste_cell(value):
        pyperclip.copy(value)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.5)

    # Add the column headers to the first worksheet row.
    paste_cell("Date and Time")
    pyautogui.press("right")
    paste_cell("Weather Data")
    pyautogui.press("right")
    paste_cell("Comment")

    # Move to the first cell of the second row.
    pyautogui.press("home")
    pyautogui.press("down")

    # Add the automatically generated date and time to column A.
    paste_cell(report_time)
    pyautogui.press("right")

    # Paste the copied weather webpage text into column B.
    paste_cell(weather_data)
    pyautogui.press("right")

    # Add a short operational comment to column C.
    paste_cell("Weather information collected successfully.")
    time.sleep(2)

    # Save the Excel workbook with today's date in its filename.
    pyautogui.hotkey("ctrl", "shift", "s")
    time.sleep(3)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(str(report_path.resolve()), interval=0.03)
    pyautogui.press("enter")
    time.sleep(4)
    pyautogui.press("enter")
    time.sleep(2)

    # Capture a screenshot of the completed Excel worksheet.
    screenshot_path = report_path.with_suffix(".png")
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_path)
    return screenshot_path


def main():
    # Add a short pause after each PyAutoGUI action.
    pyautogui.PAUSE = 0.5

    # Move the mouse to the top-left corner to stop the script.
    pyautogui.FAILSAFE = True

    # Create the folder that will contain the Excel file and screenshot.
    OUTPUT_FOLDER.mkdir(exist_ok=True)

    # Create filenames using the current date in YYYY-MM-DD format.
    current_date = datetime.now().strftime("%Y-%m-%d")
    report_path = OUTPUT_FOLDER / f"daily_report_{current_date}.xlsx"

    # Copy weather data from Chrome before opening Excel.
    weather_data = copy_weather_data()

    # Build and save the Excel report, then capture its final worksheet.
    screenshot_path = create_excel_report(weather_data, report_path)
    print(f"Excel report saved to: {report_path.resolve()}")
    print(f"Worksheet screenshot saved to: {screenshot_path.resolve()}")


if __name__ == "__main__":
    main()
