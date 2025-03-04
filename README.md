# Excel Division by Row Count App

A simple tool to divide an Excel file into multiple smaller files based on the number of rows. 

## Usage

1. Run the executable file `excel_division_by_row_count.exe` or run `excel_division_by_row_count_app.py` from the command line.

2. Select the Excel file you want to divide, and define the worksheet name.

3. Enter the row count (e.g. 2 results in excels with 2 rows (+header)).

4. Enter the name of the new worksheet.

5. Select the folder you want the files to be saved to.

6. Enter the name of the new excels (e.g. '_Student' for '1_Student.xlsx', '2_Student.xlsx', etc.).

7. Choose one of the options for dividing the Excel file:

    - "Keep the style of the cells" to retain all cell styles.

    - "No style (date and custom formats still kept)" to keep only date and custom formats.

    - "Much faster no style" for a quicker operation without retaining custom formats. Note: Dates will be converted to include hours, minutes, and seconds.

8. Click the "Submit" button to divide the Excel file. You are done.

## Requirements

- Python 3.6+
- openpyxl
- tkinter
- Pillow
- requests
- math
- pandas

## Installation

- Install Python from the official website: https://www.python.org/downloads/
- Install the required packages by running the following command in the command line:

```bash
pip install openpyxl tkinter Pillow requests pandas
```

## Author

- Mate Mihalovits, PhD - mmateo1120@gmail.com

## Version

- 1.1.0

