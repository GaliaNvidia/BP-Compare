# PO Line Comparison Tool

A Python web application that compares Excel files containing Purchase Order (PO) data between two weeks to identify changes, delays, and splits.

## Features

- **Upload Excel Files**: Compare previous week vs. current week PO data
- **Automatic Detection**:
  - Pushed PO lines (commit date changes)
  - Split PO lines (lines divided into multiple entries)
  - Re-pushed lines (pushed multiple times, indicated by >7 days)
- **Push Duration Calculation**: Calculates how many days each PO was pushed
- **Alerts**: Highlights lines pushed more than 7 days with 🚨 ALERT
- **Interactive Filtering**: Filter results by status and alert level
- **Export Results**: Download comparison results as Excel file

## Required Excel File Format

Your Excel files should contain the following columns:

| Column Name | Description | Purpose |
|------------|-------------|---------|
| `Purch.doc.` | Purchase Document Number | PO Number |
| `Item` | Item Number | PO Line |
| `Short text` | Short Text Description | Part Number (PN) |
| `Order` | Order Number | Production Work Order (PWO) |
| `Type` | Type | PO Type |
| `ComDate` | Commit Date | Date used for comparison |

## Installation

1. **Clone or download this repository**

2. **Install Python dependencies**:
```bash
pip install -r requirements.txt
```

## Usage

1. **Run the application**:
```bash
streamlit run app.py
```

2. **Open your web browser**: 
   - The app will automatically open at `http://localhost:8501`
   - If not, manually navigate to the URL shown in the terminal

3. **Upload your Excel files**:
   - Upload the previous week's Excel file in the left column
   - Upload the current week's Excel file in the right column

4. **Review the results**:
   - View summary metrics showing pushed, split, and re-pushed lines
   - See detailed comparison table with color coding:
     - 🔴 **Red background**: Lines with alerts (pushed >7 days)
     - 🟡 **Yellow background**: Split lines
   - Use filters to focus on specific statuses or alerts only

5. **Download results**:
   - Click the "Download Results as Excel" button to export the comparison

## Comparison Logic

### Pushed Lines
A PO line is considered "pushed" when:
- The same PO Number + PO Line exists in both weeks
- The current week's `ComDate` is later than the previous week's `ComDate`
- Days pushed = difference between the two dates

### Split Lines
A PO line is considered "split" when:
- A PO Number has more lines in the current week than the previous week
- The same Part Numbers appear in both weeks (indicating a split rather than a new order)

### Re-Pushed Lines
A PO line is considered "re-pushed" when:
- It was pushed more than 7 days (indicating multiple push events)
- Status will show as "Re-Pushed (>7 days)"

### Alerts
- Any PO line pushed more than 7 days receives a 🚨 ALERT flag
- These are highlighted in red in the results table

## Output Columns

The comparison results include:

- `PO_No`: Purchase Order Number
- `PO_Line`: PO Line Number
- `PN`: Part Number
- `PWO`: Production Work Order
- `PO_Type`: Purchase Order Type
- `Prev_ComDate`: Commit Date from previous week
- `Curr_ComDate`: Commit Date from current week
- `Days_Pushed`: Number of days the PO was pushed
- `Status`: Change type (Pushed, Split, Re-Pushed)
- `Alert`: Alert flag for lines pushed >7 days

## Troubleshooting

### Column Names Not Found
- Ensure your Excel files have the exact column names specified above
- Column names are case-sensitive

### Date Parsing Issues
- Ensure `ComDate` column contains valid date values
- Dates should be in Excel date format

### No Results Showing
- Verify that both files contain overlapping PO numbers
- Check that dates are properly formatted
- Ensure at least one PO line has changed between weeks

## Requirements

- Python 3.8 or higher
- See `requirements.txt` for Python package dependencies

## License

This tool is provided as-is for internal business use.

