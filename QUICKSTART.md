# Quick Start Guide

## 🚀 Getting Started in 3 Steps

### Step 1: Install Dependencies

**Option A - Automatic (Recommended)**
- **Mac/Linux**: Double-click `run.sh` or run `./run.sh` in terminal
- **Windows**: Double-click `run.bat`

**Option B - Manual**
```bash
pip install -r requirements.txt
```

### Step 2: Run the Application

**Option A - Using Scripts**
- **Mac/Linux**: `./run.sh`
- **Windows**: `run.bat`

**Option B - Manual**
```bash
streamlit run app.py
```

The app will automatically open in your browser at `http://localhost:8501`

### Step 3: Upload Your Excel Files

1. Upload your **previous week** Excel file (left side)
2. Upload your **current week** Excel file (right side)
3. View the comparison results automatically!

---

## 🧪 Testing with Sample Data

Want to test the app first? Generate sample Excel files:

```bash
python generate_sample_data.py
```

This creates two files:
- `sample_previous_week.xlsx`
- `sample_current_week.xlsx`

Upload these to the app to see how it works!

---

## 📋 Excel File Requirements

Your Excel files **must** include these columns:

| Column Name | Description |
|------------|-------------|
| `Purch.doc.` | PO Number |
| `Item` | PO Line |
| `Short text` | Part Number (PN) |
| `Order` | Production Work Order (PWO) |
| `Type` | PO Type |
| `ComDate` | Commit Date (for comparison) |

**Note**: Column names must match exactly (case-sensitive)

---

## 🎯 What the App Detects

### 1. **Pushed Lines** 🔴
- PO lines where the commit date moved to a later date
- Shows how many days it was pushed
- **Alert** if pushed >7 days

### 2. **Split Lines** 🟡
- PO lines that were divided into multiple lines
- Same PO number but more line items in current week

### 3. **Re-Pushed Lines** 🔴
- Lines pushed multiple times
- Indicated when pushed >7 days

---

## 📊 Understanding the Results

### Color Coding
- 🔴 **Red Background**: Alert! Pushed >7 days
- 🟡 **Yellow Background**: Split line
- ⚪ **No Color**: Normal push (<7 days)

### Metrics Dashboard
- **Pushed Lines**: Total number of delayed PO lines
- **Split Lines**: Total number of split PO lines
- **Re-Pushed Lines**: Lines pushed multiple times
- **🚨 Alerts**: Lines pushed more than 7 days

### Filtering
- Filter by status (Pushed, Split, Re-Pushed)
- Show alerts only (lines pushed >7 days)

---

## 💾 Exporting Results

Click **"Download Results as Excel"** to export the comparison table with:
- All comparison details
- Push duration calculations
- Alert flags
- Timestamp in filename

---

## ❓ Troubleshooting

### "Column not found" error
- Check that your Excel columns match the required names exactly
- Column names are case-sensitive: `Purch.doc.` not `purch.doc.`

### No results showing
- Ensure both files have overlapping PO numbers
- Verify that `ComDate` column contains valid dates
- Check that at least one date changed between weeks

### Date parsing issues
- Ensure `ComDate` is formatted as a date in Excel
- Dates should be Excel date format, not text

### App won't start
- Make sure Python 3.8+ is installed: `python3 --version`
- Install dependencies: `pip install -r requirements.txt`
- Try running manually: `streamlit run app.py`

---

## 🛠️ System Requirements

- **Python**: 3.8 or higher
- **OS**: Windows, Mac, or Linux
- **Browser**: Modern web browser (Chrome, Firefox, Safari, Edge)
- **RAM**: 2GB minimum (4GB recommended for large files)

---

## 📞 Need Help?

1. Check the main **README.md** for detailed documentation
2. Review the **Troubleshooting** section above
3. Generate sample data to test: `python generate_sample_data.py`

---

**Happy Comparing! 🎉**

