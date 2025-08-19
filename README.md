# Excel Data Processor

A Python application that processes Excel files with date-formatted sheets, filters data by team leader and date range, and generates summary reports.

## Features

- Processes Excel files with date-formatted sheets (dd.mm.yyyy)
- Filters data by team leader and date range
- Creates formatted summary reports
- User-friendly GUI interface
- Handles various data types and formats

## Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. Clone or download this repository
2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## How to Test

1. First, generate a test Excel file:
   ```
   python test_excel_processor.py
   ```
   This will create a file named `test_data.xlsx` with sample data.

2. Run the main application:
   ```
   python excel_processor.py
   ```

3. In the application:
   - Click "Browse..." and select the `test_data.xlsx` file
   - Select a team leader from the dropdown
   - Set the date range using the date pickers
   - Click "Process Data" to generate the output

## Test Data

The test data includes:
- 5 days of sample data (from 5 days ago to yesterday)
- 5 agents with 3 different team leaders
- Sample metrics including Login Time, Handled Inbound, AHT, etc.

## Expected Output

The application will create a new Excel file with "_filtered" suffix, containing the filtered data based on your selection.

## Troubleshooting

- If you encounter any errors, make sure all dependencies are installed correctly
- Ensure the Excel file is not open in another program when running the script
- Check that the sheet names follow the dd.mm.yyyy format
