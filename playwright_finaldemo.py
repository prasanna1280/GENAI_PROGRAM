from datetime import datetime
from pathlib import Path
import re
import subprocess
import time

import pyautogui
import pyperclip
from playwright.sync_api import sync_playwright


CRICBUZZ_URL = "https://www.cricbuzz.com/"
OUTPUT_FOLDER = Path("cricbuzz_reports")


def main():
    # Configure safe, human-like keyboard and mouse actions.
    pyautogui.PAUSE = 0.4
    pyautogui.FAILSAFE = True
    OUTPUT_FOLDER.mkdir(exist_ok=True)

    # Create filenames with the current date and time.
    current_time = datetime.now()
    date_text = current_time.strftime("Date: %Y-%m-%d")
    timestamp = current_time.strftime("%Y%m%d_%H%M%S")
    report_path = OUTPUT_FOLDER / f"cricbuzz_live_score_{timestamp}.txt"
    cricbuzz_screenshot_path = OUTPUT_FOLDER / f"cricbuzz_live_score_{timestamp}.png"
    notepad_screenshot_path = OUTPUT_FOLDER / f"notepad_{timestamp}.png"

    with sync_playwright() as playwright:
        # Open Cricbuzz in a visible Chromium browser.
        browser = playwright.chromium.launch(headless=False)
        page = browser.new_page(viewport={"width": 1440, "height": 900})
        page.goto(CRICBUZZ_URL, wait_until="domcontentloaded", timeout=60000)

        # Wait for the Cricbuzz page to render its main content.
        page.locator("body").wait_for(state="visible", timeout=30000)
        page.wait_for_timeout(5000)

        # Wait for and click the Live Scores menu.
        live_scores_link = page.get_by_role(
            "link", name=re.compile(r"live\s+scores?", re.I)
        ).first
        live_scores_link.wait_for(state="visible", timeout=30000)
        live_scores_link.click()
        page.wait_for_load_state("domcontentloaded")
        page.locator("body").wait_for(state="visible", timeout=30000)
        page.wait_for_timeout(5000)

        # Save a screenshot of the live scorecard page.
        page.screenshot(path=str(cricbuzz_screenshot_path), full_page=True)

        # Select and copy the visible live scorecard text like a user.
        page.locator("body").click(position={"x": 10, "y": 10})
        page.keyboard.press("Control+A")
        page.keyboard.press("Control+C")
        time.sleep(1)
        score_data = pyperclip.paste().strip()

        # Fall back to Playwright text extraction if the browser clipboard is empty.
        if not score_data:
            score_data = page.locator("body").inner_text().strip()
        if not score_data:
            raise RuntimeError("No Cricbuzz score data was found.")

        # Close the Playwright browser after copying the score.
        browser.close()

    # Open Windows Notepad for the copied report.
    subprocess.Popen(["notepad.exe"])
    time.sleep(2)

    # Paste the date and copied Cricbuzz data into Notepad.
    report_text = f"{date_text}\n\nCricbuzz Live Match Score\n{score_data}"
    pyperclip.copy(report_text)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(2)

    # Save the report as a text file using Notepad's Save As dialog.
    pyautogui.hotkey("ctrl", "shift", "s")
    time.sleep(2)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.write(str(report_path.resolve()), interval=0.02)
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.press("enter")
    time.sleep(2)

    # Capture a screenshot of the saved Notepad report.
    notepad_screenshot = pyautogui.screenshot()
    notepad_screenshot.save(notepad_screenshot_path)

    # Print the output paths in the terminal.
    print(f"Report saved to: {report_path.resolve()}")
    print(f"Cricbuzz screenshot saved to: {cricbuzz_screenshot_path.resolve()}")
    print(f"Notepad screenshot saved to: {notepad_screenshot_path.resolve()}")


if __name__ == "__main__":
    main()
