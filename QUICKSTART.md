# ⚡ Quick Start Guide

Get your Message Trace report in under 2 minutes!

## Step 1: Get Your CSV File

Export a detailed message trace from Microsoft 365:

1. Go to [security.microsoft.com](https://security.microsoft.com)
2. Click **Email & collaboration** → **Exchange message trace**
3. Click **Start a trace** → Set your date range, select extended report, and click **Search**
4. Click **Export** → **Download CSV**

## Step 2: Run the Script

### Option A: File Browser (Easiest)
```powershell
.\MessageTraceReport.ps1
```
A file browser will open - select your CSV file.

### Option B: Command Line
```powershell
.\MessageTraceReport.ps1 -CsvPath "C:\Downloads\MessageTrace.csv"
```

## Step 3: View Your Report

The script will ask if you want to open the report:
```
Would you like to open the report now? (Y/N): y
```

Press **Y** and your default browser will open with the interactive report!

---

## 📍 Where's My Report?

By default, the HTML report is saved in the same folder as your CSV file with a `.html` extension.

**Example:**
- Input: `C:\Downloads\MessageTrace.csv`
- Output: `C:\Downloads\MessageTrace.html`

---

## 🎯 What's in the Report?

| Section | What You'll See |
|---------|-----------------|
| **Statistics** | 9 key metrics + 14 interactive charts |
| **Message Details** | Searchable table of all messages |
| **Message Journey** | Track a message through the mail flow |
| **Compliance** | DLP rules, sensitive info types, labels |

---

## 💡 Quick Tips

- **Search**: Use the search box to filter messages across all fields
- **Sort**: Click any column header to sort
- **Details**: Click a table row to see full message details
- **Export**: Use "Export CSV" to download filtered results
- **Expand cells**: In Compliance tables, click cells to expand long values

---

## ❓ Need More Help?

See [INSTRUCTIONS.md](INSTRUCTIONS.md) for detailed documentation.
