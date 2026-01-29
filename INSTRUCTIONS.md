# Message Trace Report - Detailed Instructions

Complete reference guide for the Message Trace Report Generator.

---

## Table of Contents

1. [Installation](#installation)
2. [Getting Message Trace Data](#getting-message-trace-data)
3. [Running the Script](#running-the-script)
4. [Report Navigation](#report-navigation)
5. [Statistics & Charts](#statistics--charts)
6. [Message Details Table](#message-details-table)
7. [Message Journey Tracker](#message-journey-tracker)
8. [Compliance Investigation](#compliance-investigation)
9. [Exporting Data](#exporting-data)
10. [Troubleshooting](#troubleshooting)
11. [Column Reference](#column-reference)

---

## Installation

### Prerequisites

- **Windows PowerShell 5.1** or later (pre-installed on Windows 10/11)
- **Exchange Online** access with appropriate permissions
- Modern web browser (Chrome, Edge, Firefox)

### Setup

1. Download `MessageTraceAnalyzer.ps1` to a local folder
2. (Optional) Unblock the script if downloaded from the internet:
   ```powershell
   Unblock-File -Path .\MessageTraceAnalyzer.ps1
   ```
3. Ensure PowerShell execution policy allows running scripts:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

---

## Getting Message Trace Data

### Method 1: Microsoft 365 Defender Portal (Recommended)

1. Navigate to [https://security.microsoft.com](https://security.microsoft.com)
2. Sign in with your admin credentials
3. Go to **Email & collaboration** → **Exchange message trace**
4. Configure your search:
   - **Date range**: Select start and end dates (up to 90 days)
   - **Sender**: Filter by sender address (optional)
   - **Recipient**: Filter by recipient address (optional)
   - **Message ID**: Search for specific message (optional)
5. Select **select extended repor**
6. Click **Search**
7. Once complete, click **Export** → **Download CSV**

### Method 2: Exchange Admin Center

1. Navigate to [https://admin.exchange.microsoft.com](https://admin.exchange.microsoft.com)
2. Go to **Mail flow** → **Message trace**
3. Click **Start a trace**
4. Configure criteria and run
5. Download the CSV report

### Method 3: PowerShell (Advanced)

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Get message trace
$trace = Get-MessageTrace -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date)

# Get detailed trace
$detailedTrace = Get-MessageTraceDetail -MessageTraceId $trace[0].MessageTraceId -RecipientAddress $trace[0].RecipientAddress

# Export to CSV
$trace | Export-Csv -Path ".\MessageTrace.csv" -NoTypeInformation
```

---

## Running the Script

### Basic Usage

```powershell
# Navigate to script directory
cd "C:\Path\To\Script"

# Run with file browser
.\MessageTraceReport.ps1
```

### With Parameters

```powershell
# Specify input file
.\MessageTraceReport.ps1 -CsvPath "C:\Reports\trace.csv"

# Specify input and output
.\MessageTraceReport.ps1 -CsvPath "C:\Reports\trace.csv" -OutputPath "C:\Reports\report.html"

# Using named parameters
.\MessageTraceReport.ps1 -CsvPath ".\data.csv" -OutputPath ".\output.html"
```

### Parameters Reference

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `-CsvPath` | String | No | (file browser) | Full or relative path to the CSV file |
| `-OutputPath` | String | No | (same as CSV) | Full or relative path for HTML output |

---

## Report Navigation

The report has a sticky navigation bar at the top with four sections:

| Button | Section | Description |
|--------|---------|-------------|
| 📊 Statistics & Charts | `#stats-section` | Overview metrics and visualizations |
| 📋 Message Details | `#data-section` | Searchable message table |
| 🔍 Message Journey | `#journey-section` | Message flow tracker |
| 🔐 Compliance | `#compliance-section` | DLP and labeling analysis |

Click any button to scroll to that section.

---

## Statistics & Charts

### Key Metrics (9 Cards)

| Metric | Description |
|--------|-------------|
| Total Records | Number of message trace entries |
| Unique Senders | Count of distinct sender addresses |
| Unique Recipients | Count of distinct recipient addresses |
| Delivered | Messages with DELIVER event |
| Failed | Messages with FAIL event |
| Avg Size | Average message size in KB |
| Date Range | Span of dates in the data |
| Unique Subjects | Count of distinct subjects |
| Internal | Percentage of internal mail flow |

### Charts (14 Visualizations)

| Chart | Type | Shows |
|-------|------|-------|
| Events Distribution | Pie | Breakdown by event type |
| Direction Analysis | Doughnut | Inbound vs outbound |
| Mail Source | Pie | Source systems (STOREDRIVER, SMTP, etc.) |
| Top 10 Senders | Bar | Most active senders |
| Top 10 Recipients | Bar | Most common recipients |
| Hourly Traffic | Line | Message volume by hour |
| Size Distribution | Doughnut | Small/Medium/Large/Very Large |
| SCL Distribution | Bar | Spam Confidence Level spread |
| Daily Traffic | Bar | Messages per day |
| Weekly Pattern | Bar | Messages by day of week |
| Sender Domains | Bar | Top 10 sending domains |
| Threat Overview | Pie | Threat types detected |
| Auth Failures | Doughnut | DKIM/SPF/DMARC failures |
| Hourly Failures | Line | Failure patterns by hour |

---

## Message Details Table

### Searching

- **Global Search**: Type in "Search all fields..." to search across all columns
- **Filter by Subject**: Filter messages containing specific subject text
- **Filter by Sender**: Filter by sender email address
- **Filter by Direction**: Select "Incoming" or "Outgoing"
- **Filter by Event**: Select specific event types

### Sorting

Click any column header to sort:
- First click: Ascending (▲)
- Second click: Descending (▼)
- Third click: Reset to original order

### Columns Displayed

| Column | Description |
|--------|-------------|
| Date/Time | When the event occurred (UTC) |
| Sender | Sender email address |
| Recipient | Recipient email address |
| Subject | Message subject line |
| Event | Event type (DELIVER, RECEIVE, SEND, FAIL) |
| Direction | Incoming or Outgoing |
| Size | Message size |
| Source | Mail flow source |

### Viewing Details

Click any row to open a modal with complete message information:

- **Subject & Event Info**: Header with message subject and event badge
- **Participants**: Sender, recipient, and return path
- **Message Details**: Direction, size, recipient count, connector
- **Delivery Information**: Status, SCL, client/server IPs and hostnames
- **Technical IDs**: Message ID, Network Message ID, Internal ID, Tenant ID
- **Custom Data**: Parsed custom fields with expandable JSON

---

## Message Journey Tracker

### Searching for Messages

Enter any of these identifiers in the search box:

- **Message-ID**: Standard email Message-ID header
- **Network Message ID**: Exchange internal GUID
- **Subject**: Partial subject text match
- **Sender/Recipient**: Email address

### Journey Card Components

Each message journey shows:

1. **Header**: Subject and participant summary
2. **Summary Bar**: Total events, time span, start time, final status
3. **Flow Visualization**: Visual pipeline with event nodes
4. **Timeline**: Chronological list of all events
5. **Message IDs**: All identifiers for reference

### Event Types in Timeline

| Event | Color | Description |
|-------|-------|-------------|
| DELIVER | Green | Successfully delivered |
| RECEIVE | Blue | Received by server |
| SEND | Orange | Sent to next hop |
| FAIL | Red | Delivery failed |
| DEFER | Red | Delivery deferred |

---

## Compliance Investigation

### Tab 1: Data Loss Prevention

Analyzes Data Loss Prevention rule evaluations from CustomData.

**Summary Statistics:**
- Total Rule Evaluations
- Unique Rules
- Rules Matched / Not Matched
- Unique Policies

**Table Columns:**
| Column | Description |
|--------|-------------|
| Subject | Message subject (expandable) |
| Sender | Sender address (expandable) |
| Recipient | Recipient address (expandable) |
| Rule ID | DLP rule identifier (expandable) |
| Status | Matched ✔ or Not Matched ✖ |
| Actions | Actions taken (expandable) |
| Predicates | Conditions evaluated (expandable) |
| Date/Time | Event timestamp |

**Expandable Details Include:**
- Rule status and IDs
- Policy information
- Processing time
- Full action list with timing
- Predicate evaluation order
- Raw CustomData

### Tab 2: Sensitive Information Type

Tracks sensitive data detections and classification events.

**Summary Statistics:**
- Total Sensitive Information Type Events
- Unique Rules
- Data Classifications
- Auto-Labeling Events
- High Confidence Matches

**Event Types:**
- **DLP Data Classification**: Content matches SIT patterns
- **Server Side Auto Labeling**: Automatic label application

**Table Columns:**
| Column | Description |
|--------|-------------|
| Subject | Message subject (expandable) |
| Sender | Sender address (expandable) |
| Recipient | Recipient address (expandable) |
| Rule ID | SIT rule identifier (expandable) |
| Type | Classification or labeling |
| Predicate | Matched predicate (expandable) |
| Date/Time | Event timestamp |

### Tab 3: Sensitivity Labels

Shows sensitivity label application events.

**Summary Statistics:**
- Total Label Events
- Labeled Messages
- Unique Labels
- Default vs Custom Labels

**Content Bits Decoding:**
| Value | Meaning |
|-------|---------|
| 1 | Encryption |
| 2 | Watermark |
| 4 | Header |
| 8 | Footer |

**Table Columns:**
| Column | Description |
|--------|-------------|
| Subject | Message subject (expandable) |
| Sender | Sender address (expandable) |
| Recipient | Recipient address (expandable) |
| Label ID | Sensitivity label GUID (expandable) |
| Type | Default or Custom label |
| Content Bits | Applied protections |
| Date/Time | Event timestamp |

### Expandable Cells

In all Compliance tables, cells with long values can be expanded:
- **Click** a cell to expand and show full content
- **Click again** to collapse
- Cells show `...` indicator when truncated
- Shows `×` when expanded

---

## Exporting Data

### Export Filtered Results

Each section has an "Export CSV" button:

1. Apply your desired filters
2. Click **📥 Export CSV**
3. A CSV file downloads with currently visible/filtered data

### Export Locations

| Section | Export Button | Contains |
|---------|---------------|----------|
| Message Details | Export CSV | All filtered messages |
| DLP Rules | Export CSV | Filtered DLP rule events |
| SIT/DLP | Export CSV | Filtered classification events |
| Labels | Export CSV | Filtered label events |

---

## Troubleshooting

### Script Won't Run

**Error**: "cannot be loaded because running scripts is disabled"
```powershell
# Fix: Enable script execution
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Error**: "not digitally signed"
```powershell
# Fix: Unblock the downloaded file
Unblock-File -Path .\MessageTraceReport.ps1
```

### File Encoding Issues

**Symptom**: Strange characters or empty data

The script automatically tries multiple encodings. If problems persist:
1. Open CSV in Excel
2. Save As → CSV UTF-8 (Comma delimited)
3. Re-run the script

### Missing Data

**Charts are empty**
- Ensure you exported a **detailed** message trace
- Check the CSV has expected columns

**Compliance section is empty**
- Compliance data comes from `custom_data` column
- Not all messages have DLP/labeling events
- Ensure you have appropriate licensing (DLP, sensitivity labels)

### Browser Issues

**Charts don't display**
- Enable JavaScript
- Allow Chart.js CDN (cdn.jsdelivr.net)
- Try a different browser

**Layout looks wrong**
- Clear browser cache
- Try Chrome or Edge
- Zoom to 100%

---

## Column Reference

### Standard Message Trace Columns

| CSV Column | Report Mapping | Description |
|------------|----------------|-------------|
| `date_time_utc` | DateTime | Event timestamp in UTC |
| `sender_address` | Sender | From address |
| `recipient_address` | Recipient | To address |
| `message_subject` | Subject | Email subject |
| `event_id` | EventId | Event type code |
| `directionality` | Direction | Incoming/Outgoing |
| `total_bytes` | TotalBytes | Message size |
| `source` | Source | Mail flow source |
| `message_id` | MessageId | RFC Message-ID |
| `network_message_id` | NetworkMessageId | Exchange GUID |
| `internal_message_id` | InternalMessageId | Internal ID |
| `tenant_id` | TenantId | M365 tenant GUID |
| `client_ip` | ClientIP | Client IP address |
| `server_ip` | ServerIP | Server IP address |
| `client_hostname` | ClientHostname | Client FQDN |
| `server_hostname` | ServerHostname | Server FQDN |
| `connector_id` | ConnectorId | Connector name |
| `custom_data` | CustomData | Extended data (JSON) |
| `recipient_status` | RecipientStatus | Delivery status |
| `recipient_count` | RecipientCount | Number of recipients |
| `return_path` | ReturnPath | Return-Path header |
| `source_context` | SourceContext | Additional context |
| `message_info` | MessageInfo | Extra message info |

### CustomData Parsing

The `custom_data` column contains semi-colon separated key=value pairs that are automatically parsed:

- **S:SCL** - Spam Confidence Level
- **S:DLPRU** - DLP Rule evaluations (JSON)
- **S:DLPSI** - DLP Sensitive Information (JSON)
- **S:LabelId** - Sensitivity Label GUID
- **S:ContentBits** - Content protection bits
- **S:SPF** - SPF result
- **S:DKIM** - DKIM result
- **S:DMARC** - DMARC result

---

