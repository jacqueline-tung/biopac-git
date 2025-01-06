import os
import pyautogui
import time
from pathlib import Path
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Set up folder paths and file extension
source_folder = Path(r"F:\3. 科技部計畫-VR.ASD\Biopac\Biopac_backup (20211029)")
target_folder = Path(r"F:\3. 科技部計畫-VR.ASD\Biopac\Biopac_transfer")
file_extension = ".acq"
software_path = r"C:\Program Files (x86)\BIOPAC Systems, Inc\Biopac Student Lab 4.1\biopacstudentlab.exe"

# Get list of files with the specified extension
file_list = list(source_folder.rglob(f"*{file_extension}"))

# Create the directory structure for the target folder
for folder in source_folder.rglob("*"):
    if folder.is_dir():
        (target_folder / folder.relative_to(source_folder)).mkdir(parents=True, exist_ok=True)

# Define the automation function
def automate_conversion(input_file, output_file):
    try:
        os.startfile(software_path)
        time.sleep(2)  # Wait for software to load

        # Select software usage mode
        pyautogui.press('left')
        pyautogui.press('enter')
        time.sleep(1)

        # Select software function
        pyautogui.hotkey('alt', 'p')
        time.sleep(1)

        # Open the old file
        pyautogui.hotkey('alt', 'd')
        time.sleep(1)

        # ****Fix start from here*****
        pyautogui.typewrite(str(input_file.parent))
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'n')
        pyautogui.typewrite(input_file.name)
        pyautogui.press('enter')
        time.sleep(3)

        # Save the file with a new name and format
        pyautogui.hotkey('alt', 'f')
        pyautogui.press('a')
        pyautogui.hotkey('alt', 'd')
        pyautogui.typewrite(str(output_file.parent))
        pyautogui.press('enter')
        pyautogui.hotkey('alt', 'n')
        pyautogui.typewrite(output_file.name)
        pyautogui.hotkey('alt', 't')
        pyautogui.press('down', presses=4)  # Select "Windows AcqKnowledge 3 Graph (*acq)"
        pyautogui.press('enter')
        time.sleep(2)

        # Confirm and save the file
        pyautogui.press('enter')

        # Close the software
        pyautogui.hotkey('alt', 'f4')
        time.sleep(2)

        logging.info(f"Converted {input_file} to {output_file}")
    except Exception as e:
        logging.exception(f"Error during conversion: {e}")

# Iterate over all files and convert
for file_path in file_list:
    relative_path = file_path.relative_to(source_folder)  # Get the relative path
    # Modify the file name by changing the suffix to include "_transfer"
    output_file = target_folder / relative_path.with_name(relative_path.stem + "_transfer" + file_extension)
    automate_conversion(file_path, output_file)

print("All files have been converted.")