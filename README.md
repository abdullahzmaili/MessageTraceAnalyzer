# Message Trace Report Generator

A PowerShell script that transforms Exchange Online Message Trace CSV into interactive HTML reports with advanced filtering, visualization, and compliance analysis capabilities.

## 🌟 Features

### 📊 Statistics & Visualizations
- **9 Key Metrics Dashboard** - Total records, unique senders/recipients, delivery stats, and more
- **14 Interactive Charts** - Powered by Chart.js:
  - Event distribution (pie chart)
  - Mail direction analysis
  - Source breakdown
  - Top senders & recipients
  - Hourly/daily traffic patterns
  - Message size distribution
  - SCL (Spam Confidence Level) analysis
  - Sender domain statistics
  - Threat overview
  - Authentication failures
  - Weekday patterns

### 📋 Message Details Table
- **Full-text search** across all fields
- **Column sorting** - Click headers to sort ascending/descending
- **Advanced filtering** - Filter by sender, recipient, subject, event, or direction
- **Pagination** with configurable page sizes (10, 25, 50, 100, 250 items)
- **Export to CSV** - Download filtered results
- **Modal detail view** - Click any row for complete message information

### 🔍 Message Journey Tracker
- **Visual message flow** - Track messages through the mail pipeline
- **Search by Message ID** - Enter any Message-ID, Network Message ID, or Internal Message ID
- **Timeline view** - See all events for a message in chronological order
- **Expandable details** - Click events to see full technical information

### 🔐 Compliance Investigation (3 Tabs)

#### Data Loss Prevention Tab
- View Data Loss Prevention rule evaluations
- See matched/not-matched status
- Review actions taken and predicates evaluated
- Processing time analysis

#### Sensitive Information Type Tab
- Sensitive Information Type detections
- Confidence scores and severity levels
- Sensitive Information Type events
- Server-side auto-labeling events

#### Sensitivity Labels Tab
- Sensitivity label applications
- Content bits decoding (encryption, watermarks, headers, footers)
- Label type classification

## 📋 Requirements

- **PowerShell 5.1** or later
- **Windows OS** (for file dialog functionality)
- **CSV export** from Exchange Online Message Trace (detailed report)

## 🚀 Quick Start

```powershell
# Run with file browser
.\MessageTraceAnalyzer.ps1

# Run with specific file
.\MessageTraceAnalyzer.ps1 -CsvPath "C:\Reports\MessageTrace.csv"

# Specify output location
.\MessageTraceAnalyzer.ps1 -CsvPath ".\trace.csv" -OutputPath ".\report.html"
```

## 📥 Getting the CSV File

### From Microsoft 365 Defender Portal
1. Go to [security.microsoft.com](https://security.microsoft.com)
2. Click **Email & collaboration** → **Exchange message trace**
3. Click **Start a trace** → Set your date range, select extended report, and click **Search**
4. Click **Export** → **Download CSV**

## 📁 Output

The script generates a single HTML file containing:
- All statistics and visualizations
- Complete message data in searchable tables
- Interactive compliance analysis
- Message journey tracking

## 🎨 Report Sections

| Section | Description |
|---------|-------------|
| **Statistics & Charts** | Dashboard with key metrics and 14 interactive charts |
| **Message Details** | Searchable, sortable table of all messages |
| **Message Journey** | Visual flow tracker for individual messages |
| **Compliance** | DLP rules, SIT detections, and sensitivity labels |

## ⚙️ Parameters

| Parameter | Required | Description |
|-----------|----------|-------------|
| `-CsvPath` | No | Path to the Message Trace CSV file. Opens file browser if not provided. |
| `-OutputPath` | No | Path for the HTML output. Defaults to CSV location with .html extension. |

## 📖 Documentation

- [QUICKSTART.md](QUICKSTART.md) - Get started in 2 minutes
- [INSTRUCTIONS.md](INSTRUCTIONS.md) - Detailed usage guide and reference

## 🔧 Troubleshooting

### Common Issues

**File encoding problems**
- The script automatically tries multiple encodings (Unicode, UTF-8, Default)
- Exchange typically exports as Unicode/UTF-16

**Empty charts or missing data**
- Ensure your CSV is a **detailed** message trace export
- Check that required columns are present

**Browser compatibility**
- Works best in modern browsers (Chrome, Edge, Firefox)
- Chart.js requires JavaScript enabled

## 📝 Version History

| Version | Changes |
|---------|---------|
| 1.0 | Initial release with full feature set |
