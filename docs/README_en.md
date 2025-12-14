# Excel Data Matching & Filling Tool - User Guide
## Feature Introduction
This is a Python-based GUI desktop application that supports **multi-column joint matching** and Traditional Chinese/English bilingual switching. It is primarily used to fill data from Table 1 into corresponding columns of Table 2 based on specified matching columns.

### Core Features
✅ Support multi-column matching (e.g., order number + product code joint matching)  
✅ One-click switch between Traditional Chinese/English interface  
✅ Excel file reading/filling/saving (supports .xlsx/.xls formats)  
✅ Data preview (first 10 rows) for verifying matching results  
✅ Comprehensive error prompts and error prevention mechanisms  

## Installation & Execution
### Method 1: Run EXE Directly (Recommended)
1. Go to the GitHub Releases page to download the packaged `Excel數據匹配工具.exe`;
2. Double-click to run (no Python installation required);
3. Note: Place the EXE file in a **pure English path** (to avoid Chinese garbled characters).

### Method 2: Run from Source Code
1. Install Python 3.8~3.10 (3.10 recommended);
2. Clone the repository:
   ```bash
   git clone https://github.com/mickyleung/Excel-Match-Filler.git
   cd Excel-Match-Filler
   ```
3. Install dependencies:
   ```bash
   pip install -r src/requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
   ```
4. Run the program:
   ```bash
   python src/excel_matcher.py
   ```

### Method 3: Package EXE Manually
1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Enter the source code directory and package:
   ```bash
   cd src
   pyinstaller -F -w --hidden-import openpyxl --hidden-import xlrd --name Excel數據匹配工具 excel_matcher.py
   ```
3. The packaged EXE is located in the `src/dist/` directory.

## Usage Steps
### 1. File Selection
- **Source File (Table 1)**: Select the Excel file containing the data to be filled;
- **Target File (Table 2)**: Select the Excel file that needs data filling;
- **Worksheet**: Select the corresponding worksheet name (the tool will automatically load all worksheets in the file);
- **Output File**: Set the save path for the filled results (default: fill_result.xlsx).

### 2. Load Columns
Click the "Load Columns" button—the tool will read all column names of the selected worksheet and populate them into subsequent drop-down boxes.

### 3. Match Column Configuration (Required)
Match columns are key columns for associating data between Table 1 and Table 2, supporting multi-column matching:
1. Select a column name from the "Table 1 Match Column" drop-down box → Click "Confirm Selection";
2. Select the corresponding column name from the "Table 2 Match Column" drop-down box → Click "Confirm Selection";
3. Click "Add Match Pair"—the pair will be displayed in the list box below;
4. Repeat steps 1-3 to add multiple match pairs (e.g., order number + product code);
5. To adjust:
   - Hold down the Ctrl key to select pairs in the list box → Click "Remove Selected Match";
   - Click "Clear All Match" to reset all match pairs.

### 4. Fill Column Configuration (Required)
Fill columns are columns that need to be filled from Table 1 to Table 2:
1. Select a column name from the "Table 1 Fill Column" drop-down box → Click "Confirm Selection";
2. Select the target column name from the "Table 2 Fill Column" drop-down box → Click "Confirm Selection";
3. Click "Add Fill Column"—the pair will be displayed in the list box below;
4. Support adding/removing/clearing multiple fill pairs (same operation logic as match columns).

### 5. Preview & Execute
- **Preview Data**: Click "✅ Preview Data"—the tool will display the first 10 rows of matched and filled results for verification;
- **Run Filling**: Click "🚀 Run Filling"—the tool will fill data according to the configuration and save to the output file;
- **Reset All**: Click "🔄 Reset All" to clear all configurations and restore the initial state.

## Frequently Asked Questions
### Q1: Prompt "File is occupied" when running filling?
A1: Close the open Excel files (Table 1/Table 2/Output file) and re-run.

### Q2: Matching result is empty?
A2: Check if the column names of match columns are correct and the data formats are consistent (e.g., number/text format).

### Q3: Interface text is garbled?
A3: Run the EXE file in a pure English path (e.g., `D:\Tools\ExcelTool.exe`).

### Q4: Failed to load columns?
A4: Check if the Excel file is damaged or the worksheet name is correct.

### Q5: Data not filled after multi-column matching?
A5: Ensure all match columns have corresponding data (multi-column matching requires all columns to meet the matching conditions simultaneously).

## System Requirements
- Operating System: Windows 10/11 (64-bit);
- Python Version (Source Code Execution): 3.8~3.10 (3.11+ may have compatibility issues);
- Excel Format: Supports .xlsx/.xls (recommended .xlsx).

## Notes
1. It is recommended to back up the original Excel files before execution to avoid data loss;
2. The data formats of match columns must be consistent (e.g., "Order Number" in Table 1 is text, so it should also be text in Table 2);
3. For multi-column matching, all match columns must be satisfied simultaneously for successful matching;
4. If the EXE is misjudged as malicious software by antivirus programs, manually add it to the trust list;
5. The tool only reads and writes Excel files and does not upload any data—feel free to use it.

## Troubleshooting
| Error Phenomenon               | Possible Cause                          | Solution                                      |
|--------------------------------|-----------------------------------------|-----------------------------------------------|
| EXE fails to launch            | Path contains Chinese/special characters | Move the EXE to a pure English path (e.g., `D:\ExcelTool\`) |
| Failed to load file            | Excel file is occupied/damaged          | Close Excel and use a complete file path      |
| Duplicate data after filling   | Duplicate match column data in Table 1  | Clean up duplicate data in Table 1 and refill |
| Interface buttons are inactive | Missing Python dependencies (source run)| Reinstall dependencies: `pip install -r requirements.txt` |