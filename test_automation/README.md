# Playwright Test Automation for Chat Translator

This repository contains automated tests for evaluating the transliteration accuracy of the Singlish-to-Sinhala Chat Translator application.

## Prerequisites

- Python 3.11 or 3.12
- Google Chrome (or allow Playwright to install Chromium)

## Installation

1. Clone this repository or extract the zip file.
2. Open your Command Prompt or Terminal and navigate to the project directory:
   ```cmd
   cd test_automation
   ```
3. Install the required dependencies:
   ```cmd
   pip install -U pip
   pip install playwright openpyxl
   playwright install
   ```

## Running the Tests

To execute the automation script and record the results into the Excel file, run the following command:

```cmd
python test_automation.py --excel "Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 5000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open
```

The script will:
- Read the test cases from the Excel file (`Assignment 1 - Test cases.xlsx`).
- Automatically open a browser and type the inputs.
- Validate the actual output against the expected output.
- Record the results (`Actual output` and `Status`) back into the Excel file.

## Test Data Population
To regenerate or populate the Excel file programmatically, you can run:
```cmd
python populate_excel.py
```
