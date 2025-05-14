# Excel Data Transfer & Conversion GUI

## Overview

This Python application provides a graphical user interface (GUI) to automate the process of transferring data between multiple Excel files. It allows users to select source Excel files, a base template Excel file, and define specific cell-to-cell mappings for data transfer. Additionally, it supports mathematical conversions of values before they are pasted into the destination cells. The processed files are saved with their original names into a user-specified folder located within the same directory as the base file.

This tool is particularly useful for:
* Quickly reformatting data from multiple spreadsheets into a new template.
* Populating standard reports from various data sources.
* Performing batch data transformations across many files.

## Features

* **Multiple Source File Selection:** Select one or more source Excel files (`.xlsx`, `.xls`).
* **Base File Template:** Select a single Excel file to serve as the template for all output files.
* **Custom Output Folder:** Specify a name for an output folder, which will be created in the same directory as the base file to store the processed files.
* **Flexible Cell Mapping:**
    * Define multiple "From" (source cell) and "To" (destination cell) mappings.
    * Specify source and destination cells by row and column numbers.
* **Optional Value Conversion:**
    * Apply custom mathematical formulas to values before they are transferred.
    * Use 'X' or 'x' in the formula to represent the original cell value.
    * Supports standard Python mathematical operations and functions from the `math` module.
    * "Test Formula" button to preview the conversion with X=1.
* **Dynamic Mapping Rows:** Add or remove mapping configurations as needed.
* **Batch Processing:** Iterates through all source files and applies all defined mappings.
* **Status Log:** Provides real-time feedback on the operations being performed, files selected, and any errors encountered.

## Requirements

* Python 3.x
* `tkinter` (usually included with standard Python installations)
* `openpyxl` library

## Setup

1.  **Ensure Python 3 is installed.** You can download it from [python.org](https://www.python.org/).
2.  **Install the `openpyxl` library.** Open your terminal or command prompt and run:
    ```bash
    pip install openpyxl
    ```

## How to Run

1.  Save the application code as a Python file (e.g., `excel_transfer_gui.py`).
2.  Open your terminal or command prompt.
3.  Navigate to the directory where you saved the file.
4.  Run the script using:
    ```bash
    python excel_transfer_gui.py
    ```
    The application window will appear.

## Using the Application

The GUI is divided into several sections:

### 1. File Selection

* **Select Source Excel File(s):** Click this button to open a file dialog. You can select one or multiple `.xlsx` or `.xls` files that contain the data you want to extract. The names of selected files (or a count) will be displayed.
* **Select Base Excel File:** Click this button to choose a single `.xlsx` or `.xls` file that will be used as the template. A copy of this file will be made for each source file, and the transferred data will be pasted into this copy.
* **Output Folder Name:** Enter the desired name for the folder where the processed Excel files will be saved. This folder will be created inside the same directory as your selected Base Excel File. If a folder with that name already exists, it will be used.

### 2. Cell Mappings

This section allows you to define where data should be copied from and to. Each mapping has its own row.

* **Add Mapping Row Button (in "3. Actions"):** Click this to add a new set of mapping fields if you need to transfer data from multiple locations.
* **For each mapping row:**
    * **From:**
        * **Row:** Enter the row number of the cell in the *source file* from which to copy the value.
        * **Col:** Enter the column number of the cell in the *source file*.
    * **To:**
        * **Row:** Enter the row number of the cell in the *base file copy* where the value should be pasted.
        * **Col:** Enter the column number of the cell in the *base file copy*.
    * **Convert Checkbox:** Check this box if you want to apply a formula to the value from the source cell before pasting it.
    * **Formula (use 'X' or 'x'):** If "Convert" is checked, this text field becomes active.
        * Enter your Python-compatible mathematical formula here.
        * Use `X` (or `x`, it's case-insensitive) to represent the value read from the source cell.
        * Examples: `X * 1.1`, `X / 2 + 5`, `(X - 32) * 5/9`, `math.sqrt(X)`.
    * **Test Formula Button:** After entering a formula, click this button. The application will calculate the formula with `X=1` and display the result or "Invalid equation" in the Status Log and next to the button. This helps verify your formula syntax.

### 3. Actions

* **Add Mapping Row:** Adds another row to the "Cell Mappings" section.
* **Transfer Values:** Once all files are selected, the output folder is named, and mappings are defined, click this button to start the process.
    * The application will iterate through each selected source file.
    * For each source file, it creates a fresh copy of your chosen base file.
    * It then goes through each "From" -> "To" mapping you've defined:
        * Reads the value from the specified cell in the current source file.
        * If a conversion formula is provided and checked, it applies the formula.
        * Writes the resulting value to the specified cell in the copy of the base file.
    * Finally, this modified copy of the base file is saved with the *same name as the source file* into the specified output folder.

### Status Log

This text area at the bottom of the window displays:
* Confirmation of selected files.
* Progress during the transfer process.
* Results of formula tests.
* Any errors or warnings encountered.
* Confirmation of saved output files and their locations.

Always check the Status Log for important information, especially if something doesn't seem to work as expected.

## Important Notes

* **Excel Cell Indexing:** Row and Column numbers in the GUI are 1-based (e.g., Row 1, Col 1 is cell A1).
* **Formula Safety (`eval()`):** The conversion formula feature uses Python's `eval()` function. While powerful, `eval()` can execute arbitrary code if malicious formulas are entered. **Only use formulas from trusted sources or ones you have written yourself.**
* **Output Files:** Processed files will retain the original name of their corresponding source file and will be placed in the output folder you specified (located within the base file's directory).
* **Error Handling:** The application includes basic error handling for file operations and formula evaluation. Check the Status Log for error details.
* **Active Sheet:** The application currently operates on the *active sheet* of both the source and base Excel files.

---

This README should provide a comprehensive guide for users of your application!
