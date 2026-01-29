<#
.SYNOPSIS
    Generates an interactive HTML report from Exchange Online Message Trace CSV exports.

.DESCRIPTION
    This script reads a Message Trace CSV file exported from Exchange Online and creates
    a modern, interactive HTML report with filtering, sorting, and visualization capabilities.

.DISCLAIMER
    This script has been thoroughly tested across various environments and scenarios, and all tests have passed successfully. However, by using this script, you acknowledge and agree that:
    1. You are responsible for how you use the script and any outcomes resulting from its execution.
    2. The entire risk arising out of the use or performance of the script remains with you.
    3. The author and contributors are not liable for any damages, including data loss, business interruption, or other losses, even if warned of the risks.
        
.PARAMETER CsvPath
    Path to the Message Trace CSV file. If not specified, a file browser dialog will open.

.PARAMETER OutputPath
    Path for the generated HTML report. Defaults to same directory as CSV with .html extension.

.EXAMPLE
    .\MessageTraceAnalyzer.ps1 -CsvPath "C:\Reports\MessageTrace.csv"

.EXAMPLE
    .\MessageTraceAnalyzer.ps1
    # Opens file browser to select CSV file

.NOTES
    Author         : Abdullah Zmaili
    Version        : 1.0
    Date Created   : 2025-December-1
    Date Updated   : 2026-January-29
    Requirements:
    - PowerShell 5.1 or later, Administrator privileges for some checks

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath
)

#region Show-FileDialog
function Show-FileDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dialog.Title = "Select Message Trace CSV File"
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
    return $null
}
#endregion

#region Get-SafeProperty
function Get-SafeProperty { param($Object, $PropertyName); if ($PropertyName -and $Object.PSObject.Properties[$PropertyName]) { return $Object.$PropertyName }; return "" }
#endregion

#region Import-MessageTraceCSV
# Function to import and clean CSV data with auto-encoding detection
function Import-MessageTraceCSV {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    
    Write-Host "Reading Message Trace data from: $Path" -ForegroundColor Cyan
    
    # Try different encodings - Exchange exports often use Unicode/UTF-16
    $encodings = @('Unicode', 'UTF8', 'Default')
    $rawData = $null
    
    foreach ($enc in $encodings) {
        try {
            $testData = Import-Csv -Path $Path -Encoding $enc
            if ($testData.Count -gt 0) {
                $firstProp = ($testData | Select-Object -First 1).PSObject.Properties.Name | Select-Object -First 1
                # Check if property name looks reasonable (no null bytes)
                if ($firstProp -and $firstProp.Length -lt 50 -and $firstProp -notmatch '\x00') {
                    $rawData = $testData
                    Write-Host "Using encoding: $enc" -ForegroundColor Cyan
                    break
                }
            }
        } catch { }
    }
    
    if (-not $rawData) {
        $rawData = Import-Csv -Path $Path
    }
    
    # Check if column names have embedded quotes and create clean objects
    $firstItem = $rawData | Select-Object -First 1
    $propNames = $firstItem.PSObject.Properties.Name
    $hasQuotedColumns = ($propNames | Where-Object { $_ -match '^"' -or $_ -match '"$' }).Count -gt 0
    if ($hasQuotedColumns) {
        Write-Host "Detected quoted column names, cleaning..." -ForegroundColor Yellow
        $messageData = @()
        foreach ($row in $rawData) {
            $cleanObj = New-Object PSObject
            foreach ($prop in $row.PSObject.Properties) {
                $cleanName = $prop.Name -replace '^"|"$', ''  # Remove leading/trailing quotes
                $cleanValue = if ($prop.Value) { $prop.Value.ToString() -replace '^"|"$', '' } else { "" }
                $cleanObj | Add-Member -NotePropertyName $cleanName -NotePropertyValue $cleanValue -Force
            }
            $messageData += $cleanObj
        }
    } else { $messageData = $rawData }
    Write-Host "Successfully loaded $($messageData.Count) records" -ForegroundColor Green
    return $messageData
}
#endregion

#region Get-ColumnMappings
function Get-ColumnMappings {
    param([Parameter(Mandatory = $true)][array]$MessageData)
    $availableColumns = $MessageData | Select-Object -First 1 | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    Write-Host "Detected columns: $($availableColumns -join ', ')" -ForegroundColor Gray
    $cleanColumns = @{}; foreach ($col in $availableColumns) { $cleanColumns[$col] = $col }
    
    # Define column mappings (possible column names for each field)
    $colMappings = @{
        DateTime = @('date_time_utc', 'DateTime', 'Received', 'Date', 'origin_timestamp_utc', 'Timestamp')
        Sender = @('sender_address', 'SenderAddress', 'Sender', 'From', 'sender_from_address', 'P1FromAddress', 'P2FromAddresses')
        Recipient = @('recipient_address', 'RecipientAddress', 'Recipient', 'To', 'Recipients', 'recipient_status')
        Subject = @('message_subject', 'Subject', 'MessageSubject', 'subject')
        EventId = @('event_id', 'EventId', 'Event', 'Status', 'DeliveryStatus', 'EventType')
        Source = @('source', 'Source', 'EventSource')
        Direction = @('directionality', 'Directionality', 'Direction', 'MessageDirection')
        MessageId = @('message_id', 'MessageId', 'InternetMessageId', 'message_trace_id', 'MessageTraceId')
        TotalBytes = @('total_bytes', 'TotalBytes', 'Size', 'MessageSize')
        RecipientCount = @('recipient_count', 'RecipientCount')
        ClientIP = @('client_ip', 'ClientIP', 'original_client_ip', 'OriginalClientIP', 'SenderIP', 'FromIP')
        ServerHostname = @('server_hostname', 'ServerHostname', 'Server')
        RecipientStatus = @('recipient_status', 'RecipientStatus', 'Status', 'DeliveryStatus')
        CustomData = @('custom_data', 'CustomData', 'customdata')
        SourceContext = @('source_context', 'SourceContext')
        MessageInfo = @('message_info', 'MessageInfo')
        TenantId = @('tenant_id', 'TenantId')
        NetworkMessageId = @('network_message_id', 'NetworkMessageId')
    }
    $actualColumns = @{}
    foreach ($key in $colMappings.Keys) {
        $found = $false
        foreach ($possibleName in $colMappings[$key]) {
            if ($found) { break }
            if ($cleanColumns.ContainsKey($possibleName)) { $actualColumns[$key] = $possibleName; $found = $true; break }
        }
        if (-not $found) { $actualColumns[$key] = $null }
    }
    Write-Host "Column mapping:" -ForegroundColor Cyan
    $actualColumns.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray }
    return $actualColumns
}
#endregion

#region Get-MessageTraceStatistics
function Get-MessageTraceStatistics {
    param([Parameter(Mandatory = $true)][array]$MessageData, [Parameter(Mandatory = $true)][hashtable]$ColumnMappings)
    $stats = @{ TotalMessages = $MessageData.Count; UniqueSenders = 0; UniqueRecipients = 0; UniqueSubjects = 0; StartDate = "N/A"; EndDate = "N/A"; EventCounts = @(); DirectionCounts = @(); SourceCounts = @(); TopSenders = @(); TopRecipients = @(); HourlyCounts = @(); SizeDistribution = @(); FailedMessages = 0; DeliveredMessages = 0; AvgSizeKB = 0; PeakHour = "N/A" }
    $senderCol = $ColumnMappings['Sender']; $recipientCol = $ColumnMappings['Recipient']; $subjectCol = $ColumnMappings['Subject']
    $dateCol = $ColumnMappings['DateTime']; $eventCol = $ColumnMappings['EventId']; $directionCol = $ColumnMappings['Direction']; $sourceCol = $ColumnMappings['Source']; $sizeCol = $ColumnMappings['TotalBytes']
    $recipientCountCol = $ColumnMappings['RecipientCount']
    $stats.UniqueSenders = if ($senderCol) { ($MessageData | Where-Object { $_.$senderCol } | Select-Object -ExpandProperty $senderCol -Unique).Count } else { 0 }
    $stats.UniqueRecipients = if ($recipientCol) { ($MessageData | Where-Object { $_.$recipientCol } | ForEach-Object { $_.$recipientCol -split ';' } | Select-Object -Unique).Count } else { 0 }
    $stats.UniqueSubjects = if ($subjectCol) { ($MessageData | Where-Object { $_.$subjectCol } | Select-Object -ExpandProperty $subjectCol -Unique).Count } else { 0 }
    $dates = if ($dateCol) { $MessageData | ForEach-Object { try { [DateTime]::Parse($_.$dateCol) } catch { $null } } | Where-Object { $_ -ne $null } | Sort-Object } else { @() }
    $stats.StartDate = if ($dates.Count -gt 0) { $dates[0].ToString("yyyy-MM-dd HH:mm") } else { "N/A" }
    $stats.EndDate = if ($dates.Count -gt 0) { $dates[-1].ToString("yyyy-MM-dd HH:mm") } else { "N/A" }
    $stats.EventCounts = if ($eventCol) { $MessageData | Where-Object { $_.$eventCol } | Group-Object -Property $eventCol | Select-Object @{N='Event';E={$_.Name}}, @{N='Count';E={$_.Count}} | Sort-Object Count -Descending } else { @() }
    $stats.DirectionCounts = if ($directionCol) { $MessageData | Where-Object { $_.$directionCol } | Group-Object -Property $directionCol | Select-Object @{N='Direction';E={$_.Name}}, @{N='Count';E={$_.Count}} | Sort-Object Count -Descending } else { @() }
    $stats.SourceCounts = if ($sourceCol) { $MessageData | Where-Object { $_.$sourceCol } | Group-Object -Property $sourceCol | Select-Object @{N='Source';E={$_.Name}}, @{N='Count';E={$_.Count}} | Sort-Object Count -Descending } else { @() }
    $stats.TopSenders = if ($senderCol) { $MessageData | Where-Object { $_.$senderCol } | Group-Object -Property $senderCol | Select-Object @{N='Sender';E={$_.Name}}, @{N='Count';E={$_.Count}} | Sort-Object Count -Descending | Select-Object -First 10 } else { @() }
    $stats.TopRecipients = if ($recipientCol) { $MessageData | Where-Object { $_.$recipientCol } | ForEach-Object { $_.$recipientCol -split ';' } | Where-Object { $_ } | Group-Object | Select-Object @{N='Recipient';E={$_.Name}}, @{N='Count';E={$_.Count}} | Sort-Object Count -Descending | Select-Object -First 10 } else { @() }
    if ($dateCol) { $hourlyGroups = $MessageData | ForEach-Object { try { [DateTime]::Parse($_.$dateCol).Hour } catch { $null } } | Where-Object { $_ -ne $null } | Group-Object | Sort-Object Name; $stats.HourlyCounts = $hourlyGroups | Select-Object @{N='Hour';E={[int]$_.Name}}, @{N='Count';E={$_.Count}}; if ($hourlyGroups.Count -gt 0) { $peakGroup = $hourlyGroups | Sort-Object Count -Descending | Select-Object -First 1; $stats.PeakHour = "{0:00}:00" -f [int]$peakGroup.Name } }
    if ($sizeCol) { $sizes = $MessageData | Where-Object { $_.$sizeCol } | ForEach-Object { try { [int]$_.$sizeCol } catch { 0 } } | Where-Object { $_ -gt 0 }; if ($sizes.Count -gt 0) { $stats.AvgSizeKB = [math]::Round(($sizes | Measure-Object -Average).Average / 1024, 1); $stats.LargestMessageKB = [math]::Round(($sizes | Measure-Object -Maximum).Maximum / 1024, 1) } else { $stats.LargestMessageKB = 0 }; $stats.SizeDistribution = @(@{Range='< 10 KB';Count=($sizes | Where-Object { $_ -lt 10240 }).Count}, @{Range='10-100 KB';Count=($sizes | Where-Object { $_ -ge 10240 -and $_ -lt 102400 }).Count}, @{Range='100 KB-1 MB';Count=($sizes | Where-Object { $_ -ge 102400 -and $_ -lt 1048576 }).Count}, @{Range='> 1 MB';Count=($sizes | Where-Object { $_ -ge 1048576 }).Count}) } else { $stats.LargestMessageKB = 0 }
    if ($eventCol) { $stats.FailedMessages = ($MessageData | Where-Object { $_.$eventCol -match 'FAIL|DROP|REJECT' }).Count; $stats.DeliveredMessages = ($MessageData | Where-Object { $_.$eventCol -match 'DELIVER' }).Count; $stats.DeferredMessages = ($MessageData | Where-Object { $_.$eventCol -match 'DEFER' }).Count } else { $stats.DeferredMessages = 0 }
    $customCol = $ColumnMappings['CustomData']
    $stats.SpamMessages = 0; $stats.HighSCLMessages = 0; $stats.BulkMessages = 0; $stats.PhishMessages = 0; $stats.MalwareMessages = 0; $stats.QuarantinedMessages = 0; $stats.DKIMFailures = 0; $stats.SPFFailures = 0
    if ($customCol) { $stats.SpamMessages = ($MessageData | Where-Object { $_.$customCol -match 'SFV=SPM|SFV=SPM|sfv=Spam' }).Count; $stats.HighSCLMessages = ($MessageData | Where-Object { $_.$customCol -match 'SCL=[5-9]|scl=[5-9]' }).Count; $stats.BulkMessages = ($MessageData | Where-Object { $_.$customCol -match 'BCL=[7-9]|bcl=[7-9]' }).Count; $stats.PhishMessages = ($MessageData | Where-Object { $_.$customCol -match 'PHSH|phish|SFTY=9' }).Count; $stats.MalwareMessages = ($MessageData | Where-Object { $_.$customCol -match 'SFV=CLN.*AMP|malware|SFTY=9\.22' }).Count; $stats.QuarantinedMessages = ($MessageData | Where-Object { $_.$customCol -match 'quarantine|SFV=SKQ|SFV=SKA' }).Count; $stats.DKIMFailures = ($MessageData | Where-Object { $_.$customCol -match 'DKIM=fail|dkim=none|DKIM=none' }).Count; $stats.SPFFailures = ($MessageData | Where-Object { $_.$customCol -match 'SPF=fail|spf=fail|SPF=softfail|SPF=none' }).Count }
    # Direction-based stats
    $stats.InboundMessages = 0; $stats.OutboundMessages = 0
    if ($directionCol) { $stats.InboundMessages = ($MessageData | Where-Object { $_.$directionCol -match 'Inbound|Incoming|Receive' }).Count; $stats.OutboundMessages = ($MessageData | Where-Object { $_.$directionCol -match 'Outbound|Outgoing|Send|Originating' }).Count }
    # Date range and traffic stats
    $stats.DateRangeDays = 0; $stats.BusiestDay = "N/A"; $stats.MessagesPerHourAvg = 0
    if ($dates.Count -gt 0) { $stats.DateRangeDays = [math]::Max(1, [math]::Ceiling(($dates[-1] - $dates[0]).TotalDays)); $dailyGroups = $MessageData | ForEach-Object { try { [DateTime]::Parse($_.$dateCol).ToString("yyyy-MM-dd") } catch { $null } } | Where-Object { $_ -ne $null } | Group-Object | Sort-Object Count -Descending; if ($dailyGroups.Count -gt 0) { $stats.BusiestDay = $dailyGroups[0].Name }; $totalHours = [math]::Max(1, ($dates[-1] - $dates[0]).TotalHours); $stats.MessagesPerHourAvg = [math]::Round($MessageData.Count / $totalHours, 1) }
    # Sender/Recipient insights
    $stats.TopSenderVolume = if ($stats.TopSenders.Count -gt 0) { $stats.TopSenders[0].Count } else { 0 }
    $stats.TopRecipientVolume = if ($stats.TopRecipients.Count -gt 0) { $stats.TopRecipients[0].Count } else { 0 }
    $stats.UniqueDomains = 0
    if ($senderCol) { $stats.UniqueDomains = ($MessageData | Where-Object { $_.$senderCol -match '@' } | ForEach-Object { ($_.$senderCol -split '@')[-1] } | Select-Object -Unique).Count }
    # Average recipients per message
    $stats.AvgRecipientsPerMsg = 0
    if ($recipientCountCol) { $recipCounts = $MessageData | Where-Object { $_.$recipientCountCol } | ForEach-Object { try { [int]$_.$recipientCountCol } catch { 1 } }; if ($recipCounts.Count -gt 0) { $stats.AvgRecipientsPerMsg = [math]::Round(($recipCounts | Measure-Object -Average).Average, 1) } }
    elseif ($recipientCol) { $recipCounts = $MessageData | Where-Object { $_.$recipientCol } | ForEach-Object { ($_.$recipientCol -split ';').Count }; if ($recipCounts.Count -gt 0) { $stats.AvgRecipientsPerMsg = [math]::Round(($recipCounts | Measure-Object -Average).Average, 1) } }
    # External senders (messages from outside - simple heuristic based on direction)
    $stats.ExternalSenders = $stats.InboundMessages
    # New chart data: SCL Distribution
    $stats.SCLDistribution = @()
    if ($customCol) {
        $sclCounts = @{}; for ($i = 0; $i -le 9; $i++) { $sclCounts[$i] = 0 }
        $MessageData | ForEach-Object { if ($_.$customCol -match 'SCL=([0-9])') { $sclCounts[[int]$Matches[1]]++ } }
        $stats.SCLDistribution = @(0..9 | ForEach-Object { @{SCL=$_;Count=$sclCounts[$_]} })
    }
    # New chart data: Daily Volume Trend
    $stats.DailyCounts = @()
    if ($dateCol) { $stats.DailyCounts = $MessageData | ForEach-Object { try { [DateTime]::Parse($_.$dateCol).ToString("yyyy-MM-dd") } catch { $null } } | Where-Object { $_ -ne $null } | Group-Object | Select-Object @{N='Date';E={$_.Name}}, @{N='Count';E={$_.Count}} | Sort-Object Date }
    # New chart data: Weekday Distribution
    $stats.WeekdayCounts = @()
    if ($dateCol) {
        $weekdayOrder = @('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')
        $wdGroups = $MessageData | ForEach-Object { try { [DateTime]::Parse($_.$dateCol).DayOfWeek.ToString() } catch { $null } } | Where-Object { $_ -ne $null } | Group-Object
        $stats.WeekdayCounts = $weekdayOrder | ForEach-Object { $day = $_; $found = $wdGroups | Where-Object { $_.Name -eq $day }; @{Weekday=$day;Count=if($found){$found.Count}else{0}} }
    }
    # New chart data: Top Sender Domains
    $stats.TopSenderDomains = @()
    if ($senderCol) { $stats.TopSenderDomains = $MessageData | Where-Object { $_.$senderCol -match '@' } | ForEach-Object { ($_.$senderCol -split '@')[-1].ToLower() } | Group-Object | Select-Object @{N='Domain';E={$_.Name}}, @{N='Count';E={$_.Count}} | Sort-Object Count -Descending | Select-Object -First 10 }
    # New chart data: Threat Overview (Spam, Phish, Malware, Bulk, Clean)
    $stats.ThreatOverview = @()
    $cleanCount = $stats.TotalMessages - $stats.SpamMessages - $stats.PhishMessages - $stats.MalwareMessages - $stats.BulkMessages
    if ($cleanCount -lt 0) { $cleanCount = $stats.TotalMessages - [math]::Max($stats.SpamMessages, [math]::Max($stats.PhishMessages, [math]::Max($stats.MalwareMessages, $stats.BulkMessages))) }
    $stats.ThreatOverview = @(@{Category='Clean';Count=[math]::Max(0,$cleanCount)}, @{Category='Spam';Count=$stats.SpamMessages}, @{Category='Phishing';Count=$stats.PhishMessages}, @{Category='Malware';Count=$stats.MalwareMessages}, @{Category='Bulk';Count=$stats.BulkMessages})
    # New chart data: Authentication Failures
    $stats.AuthFailures = @(@{Type='DKIM Failures';Count=$stats.DKIMFailures}, @{Type='SPF Failures';Count=$stats.SPFFailures})
    # New chart data: Hourly Failures
    $stats.HourlyFailures = @()
    if ($dateCol -and $eventCol) {
        $failedMsgs = $MessageData | Where-Object { $_.$eventCol -match 'FAIL|DEFER|DROP|REJECT' }
        $hfGroups = $failedMsgs | ForEach-Object { try { [DateTime]::Parse($_.$dateCol).Hour } catch { $null } } | Where-Object { $_ -ne $null } | Group-Object | Sort-Object Name
        $stats.HourlyFailures = $hfGroups | Select-Object @{N='Hour';E={[int]$_.Name}}, @{N='Count';E={$_.Count}}
    }
    # New chart data: Daily Direction Trend (Inbound vs Outbound)
    $stats.DailyDirectionTrend = @()
    if ($dateCol -and $directionCol) {
        $inboundByDay = @{}; $outboundByDay = @{}
        $MessageData | ForEach-Object { try { $d = [DateTime]::Parse($_.$dateCol).ToString("yyyy-MM-dd"); $dir = $_.$directionCol; if ($dir -match 'Inbound|Incoming') { if (-not $inboundByDay.ContainsKey($d)) { $inboundByDay[$d] = 0 }; $inboundByDay[$d]++ } elseif ($dir -match 'Outbound|Outgoing|Originating') { if (-not $outboundByDay.ContainsKey($d)) { $outboundByDay[$d] = 0 }; $outboundByDay[$d]++ } } catch {} }
        $allDays = ($inboundByDay.Keys + $outboundByDay.Keys) | Select-Object -Unique | Sort-Object
        $stats.DailyDirectionTrend = $allDays | ForEach-Object { $day = $_; @{Date=$day;Inbound=$(if($inboundByDay.ContainsKey($day)){$inboundByDay[$day]}else{0});Outbound=$(if($outboundByDay.ContainsKey($day)){$outboundByDay[$day]}else{0})} }
    }
    # New chart data: Quarantine Reasons
    $stats.QuarantineReasons = @()
    if ($customCol) {
        $qSpam = ($MessageData | Where-Object { $_.$customCol -match 'SFV=SPM' -and $_.$customCol -match 'quarantine|SFV=SKQ' }).Count
        $qPhish = ($MessageData | Where-Object { $_.$customCol -match 'PHSH|phish|SFTY=9' -and $_.$customCol -match 'quarantine' }).Count
        $qMalware = ($MessageData | Where-Object { $_.$customCol -match 'malware|SFTY=9\.22' -and $_.$customCol -match 'quarantine' }).Count
        $qPolicy = ($MessageData | Where-Object { $_.$customCol -match 'SFV=SKA|SFV=SKB' }).Count
        $qOther = [math]::Max(0, $stats.QuarantinedMessages - $qSpam - $qPhish - $qMalware - $qPolicy)
        $stats.QuarantineReasons = @(@{Reason='Spam';Count=$qSpam}, @{Reason='Phishing';Count=$qPhish}, @{Reason='Malware';Count=$qMalware}, @{Reason='Policy';Count=$qPolicy}, @{Reason='Other';Count=$qOther})
    }
    return $stats
}
#endregion

#region ConvertTo-MessageTraceJson
function ConvertTo-MessageTraceJson {
    param([Parameter(Mandatory = $true)][array]$MessageData, [Parameter(Mandatory = $true)][hashtable]$ColumnMappings)
    $jsonItems = @($MessageData | ForEach-Object { @{ date_time = Get-SafeProperty $_ $ColumnMappings['DateTime']; sender = Get-SafeProperty $_ $ColumnMappings['Sender']; recipient = Get-SafeProperty $_ $ColumnMappings['Recipient']; subject = Get-SafeProperty $_ $ColumnMappings['Subject']; event_id = Get-SafeProperty $_ $ColumnMappings['EventId']; source = Get-SafeProperty $_ $ColumnMappings['Source']; directionality = Get-SafeProperty $_ $ColumnMappings['Direction']; message_id = Get-SafeProperty $_ $ColumnMappings['MessageId']; total_bytes = Get-SafeProperty $_ $ColumnMappings['TotalBytes']; recipient_count = Get-SafeProperty $_ $ColumnMappings['RecipientCount']; recipient_status = Get-SafeProperty $_ $ColumnMappings['RecipientStatus']; client_ip = Get-SafeProperty $_ $ColumnMappings['ClientIP']; server_hostname = Get-SafeProperty $_ $ColumnMappings['ServerHostname']; original_client_ip = Get-SafeProperty $_ $ColumnMappings['ClientIP']; custom_data = Get-SafeProperty $_ $ColumnMappings['CustomData']; source_context = Get-SafeProperty $_ $ColumnMappings['SourceContext']; message_info = Get-SafeProperty $_ $ColumnMappings['MessageInfo']; tenant_id = Get-SafeProperty $_ $ColumnMappings['TenantId']; network_message_id = Get-SafeProperty $_ $ColumnMappings['NetworkMessageId'] } })
    if ($jsonItems.Count -eq 0) { $jsonData = "[]" } elseif ($jsonItems.Count -eq 1) { $jsonData = "[" + ($jsonItems | ConvertTo-Json -Depth 3 -Compress) + "]" } else { $jsonData = $jsonItems | ConvertTo-Json -Depth 3 -Compress }
    Write-Host "Generated JSON with $($jsonItems.Count) items" -ForegroundColor Cyan
    return $jsonData
}
#endregion

#region ConvertTo-StatisticsJson
function ConvertTo-ArrayJson {
    param($Data)
    if ($null -eq $Data -or $Data.Count -eq 0) { return "[]" }
    $json = $Data | ConvertTo-Json -Compress
    if ($Data.Count -eq 1) { return "[$json]" }
    return $json
}

function ConvertTo-StatisticsJson {
    param([Parameter(Mandatory = $true)][hashtable]$Statistics)
    $keys = @('EventCounts','DirectionCounts','SourceCounts','TopSenders','TopRecipients','HourlyCounts','SizeDistribution','SCLDistribution','DailyCounts','WeekdayCounts','TopSenderDomains','ThreatOverview','AuthFailures','HourlyFailures','DailyDirectionTrend','QuarantineReasons')
    $result = @{}
    foreach ($key in $keys) { $result["${key}Json"] = ConvertTo-ArrayJson $Statistics[$key] }
    return $result
}
#endregion

#region Main Script Execution
#region Invoke-MessageTraceReport
function Invoke-MessageTraceReport {
    param([Parameter(Mandatory = $false)][string]$CsvFilePath, [Parameter(Mandatory = $false)][string]$OutputFilePath)
    if (-not $CsvFilePath) { Write-Host "No CSV path provided. Opening file browser..." -ForegroundColor Cyan; $CsvFilePath = Show-FileDialog; if (-not $CsvFilePath) { Write-Host "No file selected. Exiting." -ForegroundColor Yellow; return } }
    if (-not (Test-Path $CsvFilePath)) { Write-Host "Error: CSV file not found at '$CsvFilePath'" -ForegroundColor Red; return }
    if (-not $OutputFilePath) { $OutputFilePath = [System.IO.Path]::ChangeExtension($CsvFilePath, ".html") }
    try { $messageData = Import-MessageTraceCSV -Path $CsvFilePath } catch { Write-Host "Error reading CSV file: $_" -ForegroundColor Red; Write-Host $_.Exception.Message -ForegroundColor Red; return }
    $columnMappings = Get-ColumnMappings -MessageData $messageData
    $statistics = Get-MessageTraceStatistics -MessageData $messageData -ColumnMappings $columnMappings
    $jsonData = ConvertTo-MessageTraceJson -MessageData $messageData -ColumnMappings $columnMappings
    $statsJson = ConvertTo-StatisticsJson -Statistics $statistics
    $htmlContent = Get-HtmlContent -Statistics $statistics
    $htmlContent = $htmlContent.Replace('%%JSONDATA%%', $jsonData).Replace('%%EVENTCOUNTS%%', $statsJson.EventCountsJson).Replace('%%DIRECTIONCOUNTS%%', $statsJson.DirectionCountsJson).Replace('%%SOURCECOUNTS%%', $statsJson.SourceCountsJson).Replace('%%TOPSENDERS%%', $statsJson.TopSendersJson).Replace('%%TOPRECIPIENTS%%', $statsJson.TopRecipientsJson).Replace('%%HOURLYCOUNTS%%', $statsJson.HourlyCountsJson).Replace('%%SIZEDISTRIBUTION%%', $statsJson.SizeDistributionJson).Replace('%%SCLDISTRIBUTION%%', $statsJson.SCLDistributionJson).Replace('%%DAILYCOUNTS%%', $statsJson.DailyCountsJson).Replace('%%WEEKDAYCOUNTS%%', $statsJson.WeekdayCountsJson).Replace('%%TOPSENDERDOMAINS%%', $statsJson.TopSenderDomainsJson).Replace('%%THREATOVERVIEW%%', $statsJson.ThreatOverviewJson).Replace('%%AUTHFAILURES%%', $statsJson.AuthFailuresJson).Replace('%%HOURLYFAILURES%%', $statsJson.HourlyFailuresJson).Replace('%%DAILYDIRECTIONTREND%%', $statsJson.DailyDirectionTrendJson).Replace('%%QUARANTINEREASONS%%', $statsJson.QuarantineReasonsJson)
    Write-Host "HTML content length: $($htmlContent.Length)" -ForegroundColor Cyan
    try { $htmlContent | Out-File -FilePath $OutputFilePath -Encoding UTF8 -Force; Write-Host "`nReport generated successfully!" -ForegroundColor Green; Write-Host "Output file: $OutputFilePath" -ForegroundColor Cyan; $openReport = Read-Host "`nWould you like to open the report now? (Y/N)"; if ($openReport -eq 'Y' -or $openReport -eq 'y') { Start-Process $OutputFilePath } } catch { Write-Host "Error saving HTML report: $_" -ForegroundColor Red }
}
#endregion

#region Get-HtmlContent
function Get-HtmlContent {
    param([Parameter(Mandatory = $true)][hashtable]$Statistics)
    $totalMessages = $Statistics.TotalMessages; $uniqueSenders = $Statistics.UniqueSenders; $uniqueRecipients = $Statistics.UniqueRecipients; $uniqueSubjects = $Statistics.UniqueSubjects; $startDate = $Statistics.StartDate; $endDate = $Statistics.EndDate; $failedMessages = $Statistics.FailedMessages; $deliveredMessages = $Statistics.DeliveredMessages; $avgSizeKB = $Statistics.AvgSizeKB; $peakHour = $Statistics.PeakHour; $deliveryRate = if ($totalMessages -gt 0) { [math]::Round(($deliveredMessages / $totalMessages) * 100, 1) } else { 0 }; $spamMessages = $Statistics.SpamMessages; $highSCLMessages = $Statistics.HighSCLMessages; $bulkMessages = $Statistics.BulkMessages; $phishMessages = $Statistics.PhishMessages
    # New stat variables
    $inboundMessages = $Statistics.InboundMessages; $outboundMessages = $Statistics.OutboundMessages; $deferredMessages = $Statistics.DeferredMessages; $avgRecipientsPerMsg = $Statistics.AvgRecipientsPerMsg; $externalSenders = $Statistics.ExternalSenders; $dkimFailures = $Statistics.DKIMFailures; $spfFailures = $Statistics.SPFFailures; $quarantinedMessages = $Statistics.QuarantinedMessages; $malwareMessages = $Statistics.MalwareMessages; $busiestDay = $Statistics.BusiestDay; $dateRangeDays = $Statistics.DateRangeDays; $messagesPerHourAvg = $Statistics.MessagesPerHourAvg; $largestMessageKB = $Statistics.LargestMessageKB; $topSenderVolume = $Statistics.TopSenderVolume; $topRecipientVolume = $Statistics.TopRecipientVolume; $uniqueDomains = $Statistics.UniqueDomains
    $htmlTemplate = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Message Trace Report</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
:root{--primary-color:#0078d4;--secondary-color:#106ebe;--success-color:#107c10;--warning-color:#ff8c00;--danger-color:#d13438;--bg-color:#f3f2f1;--card-bg:#fff;--text-color:#323130;--text-muted:#605e5c;--border-color:#edebe9}*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background-color:var(--bg-color);color:var(--text-color);line-height:1.6}.container{max-width:1600px;margin:0 auto;padding:20px}header{background:linear-gradient(135deg,var(--primary-color),var(--secondary-color));color:#fff;padding:30px;border-radius:12px;margin-bottom:24px;box-shadow:0 4px 12px rgba(0,120,212,0.3)}header h1{font-size:2rem;font-weight:600;margin-bottom:8px}header p{opacity:0.9;font-size:0.95rem}.stats-grid{display:grid;grid-template-columns:repeat(8,1fr);gap:16px;margin-bottom:24px}.stat-card{background:var(--card-bg);padding:20px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.08);border-left:4px solid var(--primary-color);transition:transform 0.2s,box-shadow 0.2s}.stat-card:hover{transform:translateY(-2px);box-shadow:0 4px 16px rgba(0,0,0,0.12)}.stat-card.clickable{cursor:pointer}.stat-card.clickable:hover{background:linear-gradient(135deg,rgba(0,120,212,0.05),rgba(0,120,212,0.1));border-left-width:6px}.stat-card.clickable.active{background:linear-gradient(135deg,rgba(0,120,212,0.1),rgba(0,120,212,0.15));box-shadow:0 4px 20px rgba(0,120,212,0.25)}.stat-card.success{border-left-color:var(--success-color)}.stat-card.warning{border-left-color:var(--warning-color)}.stat-card.danger{border-left-color:var(--danger-color)}.stat-value{font-size:2rem;font-weight:700;color:var(--primary-color)}.stat-label{color:var(--text-muted);font-size:0.875rem;text-transform:uppercase;letter-spacing:0.5px}.charts-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px}.chart-card{background:var(--card-bg);padding:16px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.08)}.chart-card h3{color:var(--text-color);margin-bottom:12px;font-weight:600;font-size:0.95rem}.chart-container{position:relative;height:250px}.data-section{background:var(--card-bg);border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.08);overflow:hidden}.section-header{padding:20px 24px;background:var(--bg-color);border-bottom:1px solid var(--border-color);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:16px}.section-header h2{font-weight:600;font-size:1.25rem}.filters{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.filter-input{padding:8px 12px;border:1px solid var(--border-color);border-radius:6px;font-size:0.875rem;min-width:200px;transition:border-color 0.2s,box-shadow 0.2s}.filter-input:focus{outline:none;border-color:var(--primary-color);box-shadow:0 0 0 3px rgba(0,120,212,0.1)}.filter-select{padding:8px 12px;border:1px solid var(--border-color);border-radius:6px;font-size:0.875rem;background:#fff;cursor:pointer}.btn{padding:8px 16px;border:none;border-radius:6px;font-size:0.875rem;cursor:pointer;transition:background-color 0.2s}.btn-primary{background:var(--primary-color);color:#fff}.btn-primary:hover{background:var(--secondary-color)}.btn-secondary{background:var(--bg-color);color:var(--text-color);border:1px solid var(--border-color)}.btn-secondary:hover{background:var(--border-color)}.table-container{overflow-x:auto;max-height:600px;overflow-y:auto}table{width:100%;border-collapse:collapse;font-size:0.875rem}th{background:var(--bg-color);padding:12px 16px;text-align:left;font-weight:600;color:var(--text-color);border-bottom:2px solid var(--border-color);position:sticky;top:0;cursor:pointer;user-select:none;white-space:nowrap}th:hover{background:var(--border-color)}th::after{content:'';margin-left:8px}th.sort-asc::after{content:'\25B2'}th.sort-desc::after{content:'\25BC'}td{padding:12px 16px;border-bottom:1px solid var(--border-color);max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}tr:hover{background:rgba(0,120,212,0.04)}.badge{display:inline-block;padding:4px 8px;border-radius:4px;font-size:0.75rem;font-weight:600;text-transform:uppercase}.badge-deliver{background:#dff6dd;color:#107c10}.badge-receive{background:#deecf9;color:#0078d4}.badge-send{background:#fff4ce;color:#d29200}.badge-fail{background:#fde7e9;color:#d13438}.badge-incoming{background:#e1dfdd;color:#323130}.badge-outgoing{background:#f3f2f1;color:#605e5c}.pagination{display:flex;justify-content:space-between;align-items:center;padding:16px 24px;background:var(--bg-color);border-top:1px solid var(--border-color)}.pagination-info{color:var(--text-muted);font-size:0.875rem}.pagination-controls{display:flex;gap:8px}.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:1000;justify-content:center;align-items:center;backdrop-filter:blur(4px);animation:modalFadeIn 0.2s ease-out}@keyframes modalFadeIn{from{opacity:0}to{opacity:1}}@keyframes modalSlideIn{from{transform:translateY(-20px);opacity:0}to{transform:translateY(0);opacity:1}}.modal.active{display:flex}.modal-content{background:var(--card-bg);border-radius:16px;max-width:900px;width:95%;max-height:85vh;overflow-y:auto;box-shadow:0 25px 50px rgba(0,0,0,0.25);animation:modalSlideIn 0.3s ease-out}.modal-header{padding:0;border-bottom:none;position:relative}.modal-header-bg{padding:24px 28px;background:linear-gradient(135deg,var(--primary-color),var(--secondary-color));color:#fff;border-radius:16px 16px 0 0}.modal-header-bg.event-deliver{background:linear-gradient(135deg,#107c10,#0e6b0e)}.modal-header-bg.event-receive{background:linear-gradient(135deg,#0078d4,#106ebe)}.modal-header-bg.event-send{background:linear-gradient(135deg,#ff8c00,#d27500)}.modal-header-bg.event-fail{background:linear-gradient(135deg,#d13438,#a52a2d)}.modal-header h3{font-weight:600;font-size:1.1rem;margin-bottom:4px}.modal-subject{font-size:1.25rem;font-weight:700;margin-top:8px;line-height:1.3;word-break:break-word}.modal-meta{display:flex;gap:16px;margin-top:12px;flex-wrap:wrap}.modal-meta-item{display:flex;align-items:center;gap:6px;font-size:0.875rem;opacity:0.9}.modal-close{position:absolute;top:16px;right:16px;background:rgba(255,255,255,0.2);border:none;font-size:1.25rem;cursor:pointer;color:#fff;padding:8px 12px;border-radius:8px;transition:background 0.2s}.modal-close:hover{background:rgba(255,255,255,0.3)}.modal-body{padding:24px 28px;padding-bottom:80px}.modal-footer{padding:16px 28px;border-top:1px solid var(--border-color);display:flex;justify-content:center;background:var(--card-bg);position:sticky;bottom:0;box-shadow:0 -4px 12px rgba(0,0,0,0.1)}.detail-section{margin-bottom:20px;border:1px solid var(--border-color);border-radius:12px;background:var(--card-bg);box-shadow:0 2px 8px rgba(0,0,0,0.06);overflow:hidden}.detail-section:last-child{margin-bottom:0}.detail-section-header{display:flex;align-items:center;gap:10px;padding:16px 20px;background:linear-gradient(135deg,var(--primary-color),var(--secondary-color));border-bottom:1px solid var(--border-color);color:#fff}.detail-section-header.section-subject{background:linear-gradient(135deg,#8764b8,#6b4c9a)}.detail-section-header.section-participants{background:linear-gradient(135deg,#107c10,#0e6b0e)}.detail-section-header.section-message{background:linear-gradient(135deg,#0078d4,#106ebe)}.detail-section-header.section-delivery{background:linear-gradient(135deg,#ff8c00,#d27500)}.detail-section-header.section-technical{background:linear-gradient(135deg,#00b7c3,#009ca6)}.detail-section-header.section-custom{background:linear-gradient(135deg,#d13438,#a52a2d)}.detail-section.collapsed .detail-section-header{border-bottom:none}.detail-section-icon{font-size:1.25rem}.detail-section-title{font-weight:600;font-size:1rem;flex-grow:1;color:#fff}.detail-toggle-btn{padding:6px 14px;border:1px solid #fff;background:rgba(255,255,255,0.2);color:#fff;border-radius:6px;font-size:0.75rem;font-weight:600;cursor:pointer;transition:all 0.2s;display:flex;align-items:center;gap:6px}.detail-toggle-btn:hover{background:rgba(255,255,255,0.3);color:#fff}.detail-toggle-btn.hide-btn{background:rgba(255,255,255,0.3);color:#fff}.detail-toggle-btn.hide-btn:hover{background:rgba(255,255,255,0.4);border-color:#fff}.detail-section.collapsed .detail-section-content{display:none}.detail-section-content{padding:16px 20px;background:#fafbfc}.detail-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px}.detail-item{padding:14px 16px;background:#fff;border-radius:10px;border:1px solid var(--border-color);transition:all 0.2s;position:relative}.detail-item:hover{border-color:var(--primary-color);box-shadow:0 2px 8px rgba(0,120,212,0.1)}.detail-item.full-width{grid-column:1/-1}.detail-item-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}.detail-label{display:flex;align-items:center;gap:6px;font-size:0.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;font-weight:600}.detail-label-icon{font-size:0.875rem}.detail-value{font-size:0.9rem;word-break:break-all;color:var(--text-color);line-height:1.5}.detail-value.monospace{font-family:'Consolas','Monaco',monospace;font-size:0.8rem;background:#e8e8e8;padding:8px 10px;border-radius:6px;margin-top:4px}.copy-btn{background:none;border:1px solid var(--border-color);padding:4px 8px;border-radius:4px;cursor:pointer;font-size:0.7rem;color:var(--text-muted);transition:all 0.2s;display:flex;align-items:center;gap:4px}.copy-btn:hover{background:var(--primary-color);color:#fff;border-color:var(--primary-color)}.copy-btn.copied{background:var(--success-color);color:#fff;border-color:var(--success-color)}.status-badge{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:20px;font-size:0.8rem;font-weight:600}.status-badge.deliver{background:#dff6dd;color:#107c10}.status-badge.receive{background:#deecf9;color:#0078d4}.status-badge.send{background:#fff4ce;color:#d29200}.status-badge.fail{background:#fde7e9;color:#d13438}.direction-badge{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:20px;font-size:0.8rem;font-weight:600;background:#e8e8e8}.direction-badge.incoming{background:#deecf9;color:#0078d4}.direction-badge.outgoing{background:#fff4ce;color:#d29200}.custom-data-section{margin-top:20px;border-top:1px solid var(--border-color)}.custom-data-container{display:grid;grid-template-columns:1fr;gap:16px}.custom-data-block{background:var(--bg-color);border-radius:8px;padding:12px;border-left:4px solid var(--primary-color)}.custom-data-header{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap}.custom-data-key{font-weight:600;color:var(--primary-color);cursor:help}.custom-data-code{font-family:'Consolas',monospace;font-size:0.75rem;color:var(--text-muted);background:#e0e0e0;padding:2px 6px;border-radius:4px}.custom-data-value{font-size:0.875rem;padding:8px;background:#fff;border-radius:4px;word-break:break-word}.custom-data-table{width:100%;font-size:0.8rem;border-collapse:collapse}.custom-data-table td{padding:6px 8px;border-bottom:1px solid var(--border-color);vertical-align:top}.custom-data-table tr:last-child td{border-bottom:none}.custom-data-table .sub-key{font-weight:500;color:var(--text-color);width:40%;background:rgba(0,120,212,0.05);cursor:help}.custom-data-table .sub-value{color:var(--text-muted);word-break:break-all}.code-hint{font-family:'Consolas',monospace;font-size:0.7rem;color:#888}.raw-data-details{margin-top:16px}.raw-data-details summary{cursor:pointer;color:var(--primary-color);font-size:0.875rem;padding:8px;background:var(--bg-color);border-radius:4px}.raw-data-details summary:hover{background:#e0e0e0}.raw-data-pre{margin-top:8px;padding:12px;background:#2d2d2d;color:#f0f0f0;border-radius:6px;font-family:'Consolas',monospace;font-size:0.75rem;overflow-x:auto;white-space:pre-wrap;word-break:break-all;max-height:300px;overflow-y:auto}.loading{text-align:center;padding:40px;color:var(--text-muted)}.export-buttons{display:flex;gap:8px}.nav-menu{position:sticky;top:0;z-index:100;background:var(--card-bg);padding:12px 20px;border-radius:8px;margin-bottom:20px;box-shadow:0 2px 12px rgba(0,0,0,0.1);display:flex;gap:8px;flex-wrap:wrap;justify-content:center}.nav-btn{padding:10px 20px;border:none;border-radius:6px;font-size:0.9rem;font-weight:600;cursor:pointer;transition:all 0.2s;display:flex;align-items:center;gap:8px;background:var(--bg-color);color:var(--text-color);border:1px solid var(--border-color)}.nav-btn:hover{background:var(--primary-color);color:#fff;border-color:var(--primary-color);transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,120,212,0.3)}.nav-btn.active{background:var(--primary-color);color:#fff;border-color:var(--primary-color)}.nav-btn .nav-icon{font-size:1.1rem}.journey-section{background:var(--card-bg);border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,0.1);margin-bottom:24px;overflow:hidden}.journey-header{padding:24px;background:linear-gradient(135deg,#667eea,#764ba2);border-bottom:none}.journey-header h2{font-weight:600;font-size:1.4rem;margin-bottom:16px;color:#fff;display:flex;align-items:center;gap:12px}.journey-search{display:flex;gap:12px;flex-wrap:wrap;align-items:center}.journey-search input{flex:1;min-width:300px;padding:14px 20px;border:none;border-radius:10px;font-size:1rem;background:rgba(255,255,255,0.95);box-shadow:0 2px 10px rgba(0,0,0,0.1);transition:all 0.3s}.journey-search input:focus{outline:none;box-shadow:0 4px 20px rgba(0,0,0,0.15);transform:translateY(-1px)}.journey-search input::placeholder{color:#999}.journey-search button{padding:14px 28px;font-size:1rem;border-radius:10px;font-weight:600;transition:all 0.3s}.journey-search .btn-primary{background:rgba(255,255,255,0.2);border:2px solid #fff;color:#fff}.journey-search .btn-primary:hover{background:#fff;color:#667eea}.journey-search .btn-secondary{background:transparent;border:2px solid rgba(255,255,255,0.5);color:#fff}.journey-search .btn-secondary:hover{background:rgba(255,255,255,0.1);border-color:#fff}.journey-results{padding:24px;background:#f8f9fa}.journey-empty{text-align:center;padding:80px 20px;color:var(--text-muted);background:linear-gradient(180deg,#fff,#f8f9fa);border-radius:16px;border:2px dashed var(--border-color)}.journey-empty-icon{font-size:5rem;margin-bottom:20px;opacity:0.4;animation:float 3s ease-in-out infinite}@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(-10px)}}@keyframes slideIn{from{opacity:0;transform:translateX(-20px)}to{opacity:1;transform:translateX(0)}}@keyframes pulse{0%,100%{transform:scale(1)}50%{transform:scale(1.1)}}.journey-card{background:#fff;border-radius:16px;padding:0;margin-bottom:24px;box-shadow:0 4px 20px rgba(0,0,0,0.08);overflow:hidden;animation:slideIn 0.4s ease-out;border:1px solid rgba(0,0,0,0.05)}.journey-card-header{padding:24px;background:linear-gradient(135deg,#f8f9fa,#fff);border-bottom:1px solid var(--border-color)}.journey-card-title{font-size:1.2rem;font-weight:700;color:var(--text-color);word-break:break-word;margin-bottom:12px;line-height:1.4}.journey-card-meta{display:flex;gap:20px;flex-wrap:wrap;font-size:0.9rem;color:var(--text-muted)}.journey-card-meta span{display:flex;align-items:center;gap:6px;background:#f0f0f0;padding:6px 12px;border-radius:20px}.journey-summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:16px;padding:20px 24px;background:linear-gradient(135deg,#667eea,#764ba2);margin:0}.journey-summary-item{text-align:center;color:#fff}.journey-summary-value{font-size:1.8rem;font-weight:700;display:block}.journey-summary-label{font-size:0.8rem;opacity:0.9;text-transform:uppercase;letter-spacing:1px}.journey-flow{padding:24px;background:#fff}.journey-flow-visual{display:flex;align-items:center;justify-content:center;gap:8px;padding:20px;background:linear-gradient(135deg,#f8f9fa,#fff);border-radius:12px;margin-bottom:24px;flex-wrap:wrap}.journey-flow-node{display:flex;flex-direction:column;align-items:center;gap:8px;padding:16px 20px;background:#fff;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.08);transition:all 0.3s;cursor:pointer;min-width:100px;border:2px solid transparent}.journey-flow-node:hover{transform:translateY(-4px);box-shadow:0 8px 25px rgba(0,0,0,0.15)}.journey-flow-node.active{border-color:var(--primary-color);background:rgba(0,120,212,0.05)}.journey-flow-node.deliver{border-color:var(--success-color)}.journey-flow-node.fail{border-color:var(--danger-color)}.journey-flow-node-icon{font-size:2rem}.journey-flow-node-label{font-size:0.75rem;font-weight:600;text-transform:uppercase;color:var(--text-muted)}.journey-flow-arrow{font-size:1.5rem;color:var(--border-color);animation:arrowMove 1s ease-in-out infinite}@keyframes arrowMove{0%,100%{transform:translateX(0);opacity:0.5}50%{transform:translateX(5px);opacity:1}}.journey-timeline{position:relative;padding:0 24px 24px}.journey-timeline-line{position:absolute;left:47px;top:0;bottom:24px;width:3px;background:linear-gradient(180deg,var(--primary-color),var(--success-color));border-radius:3px}.journey-step{position:relative;padding-left:60px;padding-bottom:24px;animation:slideIn 0.4s ease-out backwards}.journey-step:nth-child(1){animation-delay:0.1s}.journey-step:nth-child(2){animation-delay:0.2s}.journey-step:nth-child(3){animation-delay:0.3s}.journey-step:nth-child(4){animation-delay:0.4s}.journey-step:nth-child(5){animation-delay:0.5s}.journey-step:last-child{padding-bottom:0}.journey-step-dot{position:absolute;left:0;top:0;width:40px;height:40px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:1rem;color:#fff;z-index:2;box-shadow:0 4px 15px rgba(0,0,0,0.2);transition:all 0.3s;cursor:pointer}.journey-step-dot:hover{transform:scale(1.15)}.journey-step-dot.deliver{background:linear-gradient(135deg,#10b981,#059669)}.journey-step-dot.receive{background:linear-gradient(135deg,#3b82f6,#2563eb)}.journey-step-dot.send{background:linear-gradient(135deg,#f59e0b,#d97706)}.journey-step-dot.fail,.journey-step-dot.defer{background:linear-gradient(135deg,#ef4444,#dc2626)}.journey-step-dot.default{background:linear-gradient(135deg,#6b7280,#4b5563)}.journey-step-connector{position:absolute;left:18px;top:40px;bottom:-24px;width:4px;background:linear-gradient(180deg,currentColor,var(--border-color))}.journey-step:last-child .journey-step-connector{display:none}.journey-step-content{background:#fff;border-radius:12px;padding:20px;border:1px solid var(--border-color);transition:all 0.3s;cursor:pointer;position:relative;overflow:hidden}.journey-step-content::before{content:'';position:absolute;top:0;left:0;width:4px;height:100%;background:var(--primary-color);opacity:0;transition:opacity 0.3s}.journey-step-content:hover{box-shadow:0 8px 25px rgba(0,0,0,0.1);transform:translateX(4px)}.journey-step-content:hover::before{opacity:1}.journey-step-content.expanded{background:#f8f9fa}.journey-step-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:12px}.journey-step-event{font-weight:700;font-size:0.95rem;padding:8px 16px;border-radius:20px;display:inline-flex;align-items:center;gap:8px}.journey-step-event.deliver{background:#d1fae5;color:#065f46}.journey-step-event.receive{background:#dbeafe;color:#1e40af}.journey-step-event.send{background:#fef3c7;color:#92400e}.journey-step-event.fail,.journey-step-event.defer{background:#fee2e2;color:#991b1b}.journey-step-event.default{background:#f3f4f6;color:#374151}.journey-step-time{font-size:0.85rem;color:var(--text-muted);display:flex;align-items:center;gap:6px;background:#f3f4f6;padding:6px 12px;border-radius:6px}.journey-step-details{display:none;margin-top:16px;padding-top:16px;border-top:1px dashed var(--border-color);animation:slideIn 0.3s ease-out}.journey-step-details.show{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}.journey-step-detail{background:#f8f9fa;padding:12px;border-radius:8px;transition:all 0.2s}.journey-step-detail:hover{background:#f0f0f0}.journey-step-detail-label{font-size:0.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;font-weight:600}.journey-step-detail-value{font-size:0.9rem;color:var(--text-color);word-break:break-all;font-weight:500}.journey-step-expand{font-size:0.8rem;color:var(--primary-color);cursor:pointer;display:flex;align-items:center;gap:4px;margin-top:12px;font-weight:600;transition:all 0.2s}.journey-step-expand:hover{color:var(--secondary-color)}.journey-ids{padding:20px 24px;background:#f8f9fa;border-top:1px solid var(--border-color)}.journey-ids-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:16px}.journey-id-item{background:#fff;padding:16px;border-radius:10px;border:1px solid var(--border-color)}.journey-id-label{font-size:0.75rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;font-weight:600}.journey-id-value{font-family:'Consolas',monospace;font-size:0.85rem;color:var(--text-color);word-break:break-all;background:#f8f9fa;padding:10px;border-radius:6px;cursor:pointer;transition:all 0.2s;display:flex;justify-content:space-between;align-items:center;gap:8px}.journey-id-value:hover{background:#e8e8e8}.journey-id-copy{font-size:0.7rem;color:var(--primary-color);white-space:nowrap}.journey-no-results{text-align:center;padding:60px;color:var(--text-muted);background:#fff;border-radius:16px;border:2px dashed var(--border-color)}.journey-no-results .journey-empty-icon{animation:float 3s ease-in-out infinite}.compliance-section{background:var(--card-bg);border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,0.1);margin-bottom:24px;overflow:hidden}.compliance-header{padding:24px;background:linear-gradient(135deg,#e74c3c,#c0392b);border-bottom:none}.compliance-header h2{font-weight:600;font-size:1.4rem;margin-bottom:8px;color:#fff;display:flex;align-items:center;gap:12px}.compliance-header p{color:rgba(255,255,255,0.9);font-size:0.9rem}.compliance-tabs{display:flex;gap:0;background:#f8f9fa;border-bottom:1px solid var(--border-color)}.compliance-tab{flex:1;padding:16px 24px;background:transparent;border:none;font-size:0.95rem;font-weight:600;cursor:pointer;transition:all 0.3s;color:var(--text-muted);border-bottom:3px solid transparent;display:flex;align-items:center;justify-content:center;gap:8px}.compliance-tab:hover{background:#fff;color:var(--text-color)}.compliance-tab.active{background:#fff;color:var(--primary-color);border-bottom-color:var(--primary-color)}.compliance-tab-icon{font-size:1.2rem}.compliance-tab-count{background:var(--primary-color);color:#fff;padding:2px 8px;border-radius:12px;font-size:0.75rem;min-width:24px;text-align:center}.compliance-content{padding:24px}.compliance-panel{display:none;animation:slideIn 0.3s ease-out}.compliance-panel.active{display:block}.compliance-summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:24px}.compliance-stat{background:linear-gradient(135deg,#f8f9fa,#fff);padding:20px;border-radius:12px;border:1px solid var(--border-color);text-align:center;transition:all 0.3s}.compliance-stat:hover{transform:translateY(-2px);box-shadow:0 4px 15px rgba(0,0,0,0.1)}.compliance-stat.clickable{cursor:pointer}.compliance-stat.clickable:hover{background:linear-gradient(135deg,rgba(0,120,212,0.1),rgba(0,120,212,0.15))}.compliance-stat.clickable.active{background:linear-gradient(135deg,rgba(0,120,212,0.15),rgba(0,120,212,0.2));box-shadow:0 4px 20px rgba(0,120,212,0.25);border-color:var(--primary-color)}.compliance-stat-value{font-size:2rem;font-weight:700;color:var(--primary-color);display:block}.compliance-stat-label{font-size:0.8rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:0.5px}.compliance-stat.warning .compliance-stat-value{color:var(--warning-color)}.compliance-stat.danger .compliance-stat-value{color:var(--danger-color)}.compliance-stat.success .compliance-stat-value{color:var(--success-color)}.compliance-table-container{background:#fff;border-radius:12px;border:1px solid var(--border-color);overflow:hidden}.compliance-table{width:100%;border-collapse:collapse}.compliance-table th{background:linear-gradient(135deg,#f8f9fa,#fff);padding:14px 16px;text-align:left;font-weight:600;font-size:0.85rem;color:var(--text-color);border-bottom:2px solid var(--border-color)}.compliance-table td{padding:14px 16px;border-bottom:1px solid var(--border-color);font-size:0.9rem}.compliance-table tr:hover{background:rgba(0,120,212,0.04)}.compliance-table tr:last-child td{border-bottom:none}.confidence-badge{display:inline-flex;align-items:center;gap:6px;padding:4px 12px;border-radius:20px;font-size:0.8rem;font-weight:600}.confidence-high{background:#d1fae5;color:#065f46}.confidence-medium{background:#fef3c7;color:#92400e}.confidence-low{background:#fee2e2;color:#991b1b}.label-badge{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:8px;font-size:0.85rem;font-weight:500;background:linear-gradient(135deg,#667eea,#764ba2);color:#fff}.content-bits-badge{display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:6px;font-size:0.75rem;font-weight:600;background:#f3f4f6;color:#374151;margin:2px}.content-bits-badge.encryption{background:#fee2e2;color:#991b1b}.content-bits-badge.watermark{background:#dbeafe;color:#1e40af}.content-bits-badge.header{background:#d1fae5;color:#065f46}.content-bits-badge.footer{background:#fef3c7;color:#92400e}.sit-id{font-family:'Consolas',monospace;font-size:0.8rem;color:var(--text-muted);word-break:break-all}.compliance-detail-row{cursor:pointer;transition:all 0.2s}.compliance-detail-row:hover{background:rgba(0,120,212,0.08)}.compliance-expand-icon{transition:transform 0.3s}.compliance-detail-row.expanded .compliance-expand-icon{transform:rotate(90deg)}.compliance-detail-content{display:none;background:#f8f9fa;padding:16px;border-bottom:1px solid var(--border-color)}.compliance-detail-content.show{display:block}.compliance-detail-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:12px}.compliance-detail-item{background:#fff;padding:12px;border-radius:8px;border:1px solid var(--border-color);overflow:hidden;min-width:0}.compliance-detail-label{font-size:0.7rem;color:var(--text-muted);text-transform:uppercase;margin-bottom:4px;font-weight:600}.compliance-detail-value{font-size:0.9rem;color:var(--text-color);word-break:break-word;overflow-wrap:break-word;white-space:normal}.compliance-empty{text-align:center;padding:60px 20px;color:var(--text-muted)}.compliance-empty-icon{font-size:4rem;margin-bottom:16px;opacity:0.4}.dlp-rule-status{display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:20px;font-size:0.8rem;font-weight:600}.dlp-rule-status.matched{background:#d1fae5;color:#065f46}.dlp-rule-status.not-matched{background:#fee2e2;color:#991b1b}.dlp-actions-list,.dlp-predicates-list{max-height:100px;overflow-y:auto;font-size:0.8rem;display:flex;flex-wrap:wrap}.dlp-predicates-list{max-width:200px}.dlp-action-item,.dlp-predicate-item{display:inline-flex;align-items:center;gap:6px;padding:4px 8px;margin:2px 4px 2px 0;background:#f3f4f6;border-radius:4px;font-family:'Consolas',monospace;font-size:0.75rem;white-space:nowrap;width:fit-content}.dlp-action-item{background:#d1fae5;color:#065f46}.dlp-predicate-item{background:#dbeafe;color:#1e40af}.dlp-predicate-item.compact{font-size:0.7rem;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:3px 6px}.dlp-rule-detail-row{cursor:pointer;transition:all 0.2s}.dlp-rule-detail-row:hover{background:rgba(0,120,212,0.08)}.dlp-rule-expand-icon{transition:transform 0.3s}.dlp-rule-detail-row.expanded .dlp-rule-expand-icon{transform:rotate(90deg)}.dlp-rule-detail-content{display:none;background:#f8f9fa;padding:16px;border-bottom:1px solid var(--border-color)}.dlp-rule-detail-content.show{display:block}@media(max-width:1400px){.charts-grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:768px){.charts-grid{grid-template-columns:1fr}.filters{flex-direction:column;width:100%}.filter-input,.filter-select{width:100%}.nav-menu{flex-direction:column}.nav-btn{width:100%;justify-content:center}.journey-search{flex-direction:column}.journey-search input{min-width:100%}.journey-flow-visual{flex-direction:column}.journey-flow-arrow{transform:rotate(90deg)}.journey-summary{grid-template-columns:repeat(2,1fr)}}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>&#9993; Message Trace Report</h1>
            <p>Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") | Date Range: $startDate to $endDate UTC</p>
        </header>
        
        <nav class="nav-menu">
            <button class="nav-btn" onclick="scrollToSection('stats-section')">
                <span class="nav-icon">&#128202;</span> Statistics & Charts
            </button>
            <button class="nav-btn" onclick="scrollToSection('data-section')">
                <span class="nav-icon">&#128203;</span> Message Details
            </button>
            <button class="nav-btn" onclick="scrollToSection('journey-section')">
                <span class="nav-icon">&#128269;</span> Message Journey
            </button>
            <button class="nav-btn" onclick="scrollToSection('compliance-section')">
                <span class="nav-icon">&#128274;</span> Compliance
            </button>
        </nav>
        
        <div id="stats-section" class="stats-grid">
            <div class="stat-card clickable" onclick="filterByStatCard('all')" title="Click to show all records">
                <div class="stat-value" id="totalRecords">$totalMessages</div>
                <div class="stat-label">Total Records</div>
            </div>
            <div class="stat-card success clickable" onclick="filterByStatCard('senders')" title="Click to see unique senders">
                <div class="stat-value">$uniqueSenders</div>
                <div class="stat-label">Unique Senders</div>
            </div>
            <div class="stat-card warning clickable" onclick="filterByStatCard('recipients')" title="Click to see unique recipients">
                <div class="stat-value">$uniqueRecipients</div>
                <div class="stat-label">Unique Recipients</div>
            </div>
            <div class="stat-card success clickable" onclick="filterByStatCard('delivered')" title="Click to see delivered messages">
                <div class="stat-value">$deliveredMessages</div>
                <div class="stat-label">Delivered Messages</div>
            </div>
            <div class="stat-card danger clickable" onclick="filterByStatCard('failed')" title="Click to see failed/deferred messages">
                <div class="stat-value">$failedMessages</div>
                <div class="stat-label">Failed/Deferred</div>
            </div>
            <div class="stat-card danger clickable" onclick="filterByStatCard('spam')" title="Click to see spam messages">
                <div class="stat-value">$spamMessages</div>
                <div class="stat-label">&#128683; Spam Detected</div>
            </div>
            <div class="stat-card warning clickable" onclick="filterByStatCard('highscl')" title="Click to see high SCL messages">
                <div class="stat-value">$highSCLMessages</div>
                <div class="stat-label">&#9888; High SCL (5+)</div>
            </div>
            <div class="stat-card clickable" onclick="filterByStatCard('bulk')" title="Click to see bulk mail">
                <div class="stat-value">$bulkMessages</div>
                <div class="stat-label">&#128231; Bulk Mail</div>
            </div>
            <div class="stat-card danger clickable" onclick="filterByStatCard('phish')" title="Click to see phishing messages">
                <div class="stat-value">$phishMessages</div>
                <div class="stat-label">&#127907; Phishing</div>
            </div>
            <div class="stat-card clickable" onclick="filterByStatCard('inbound')" title="Click to see inbound messages">
                <div class="stat-value">$inboundMessages</div>
                <div class="stat-label">&#128229; Inbound Messages</div>
            </div>
            <div class="stat-card clickable" onclick="filterByStatCard('outbound')" title="Click to see outbound messages">
                <div class="stat-value">$outboundMessages</div>
                <div class="stat-label">&#128228; Outbound Messages</div>
            </div>
            <div class="stat-card danger clickable" onclick="filterByStatCard('dkim')" title="Click to see DKIM failures">
                <div class="stat-value">$dkimFailures</div>
                <div class="stat-label">&#128274; DKIM Failures</div>
            </div>
            <div class="stat-card danger clickable" onclick="filterByStatCard('spf')" title="Click to see SPF failures">
                <div class="stat-value">$spfFailures</div>
                <div class="stat-label">&#128272; SPF Failures</div>
            </div>
            <div class="stat-card warning clickable" onclick="filterByStatCard('quarantined')" title="Click to see quarantined messages">
                <div class="stat-value">$quarantinedMessages</div>
                <div class="stat-label">&#128451; Quarantined</div>
            </div>
            <div class="stat-card danger clickable" onclick="filterByStatCard('malware')" title="Click to see malware messages">
                <div class="stat-value">$malwareMessages</div>
                <div class="stat-label">&#128026; Malware Detected</div>
            </div>
            <div class="stat-card clickable" onclick="filterByStatCard('domains')" title="Click to see messages by domain">
                <div class="stat-value">$uniqueDomains</div>
                <div class="stat-label">&#127968; Unique Domains</div>
            </div>
        </div>
        
        <div class="charts-grid">
            <div class="chart-card">
                <h3>&#128202; Events by Type</h3>
                <div class="chart-container">
                    <canvas id="eventChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#8644; Message Direction</h3>
                <div class="chart-container">
                    <canvas id="directionChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#9881; Events by Source</h3>
                <div class="chart-container">
                    <canvas id="sourceChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128100; Top Senders</h3>
                <div class="chart-container">
                    <canvas id="senderChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128229; Top Recipients</h3>
                <div class="chart-container">
                    <canvas id="recipientChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#10004; Delivery Success Rate</h3>
                <div class="chart-container">
                    <canvas id="deliveryChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128737; Threat Overview</h3>
                <div class="chart-container">
                    <canvas id="threatChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128274; Authentication Failures</h3>
                <div class="chart-container">
                    <canvas id="authChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128200; SCL Distribution</h3>
                <div class="chart-container">
                    <canvas id="sclChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128197; Daily Volume Trend</h3>
                <div class="chart-container">
                    <canvas id="dailyChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128198; Weekday Distribution</h3>
                <div class="chart-container">
                    <canvas id="weekdayChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128506; Inbound vs Outbound Trend</h3>
                <div class="chart-container">
                    <canvas id="directionTrendChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#127760; Top Sender Domains</h3>
                <div class="chart-container">
                    <canvas id="domainChart"></canvas>
                </div>
            </div>
            <div class="chart-card">
                <h3>&#128683; Failed Messages by Hour</h3>
                <div class="chart-container">
                    <canvas id="hourlyFailuresChart"></canvas>
                </div>
            </div>
        </div>
        
        <div id="data-section" class="data-section">
            <div class="section-header">
                <h2>Message Details</h2>
                <div class="filters">
                    <input type="text" class="filter-input" id="searchInput" placeholder="Search all fields...">
                    <select class="filter-select" id="eventFilter">
                        <option value="">All Events</option>
                    </select>
                    <select class="filter-select" id="directionFilter">
                        <option value="">All Directions</option>
                    </select>
                    <select class="filter-select" id="sourceFilter">
                        <option value="">All Sources</option>
                    </select>
                    <div class="export-buttons">
                        <button class="btn btn-secondary" onclick="exportToCSV()">Export CSV</button>
                        <button class="btn btn-secondary" onclick="resetFilters()">Reset Filters</button>
                    </div>
                </div>
            </div>
            
            <div class="table-container">
                <table id="dataTable">
                    <thead>
                        <tr>
                            <th data-sort="date_time">Date/Time (UTC)</th>
                            <th data-sort="sender">Sender</th>
                            <th data-sort="recipient">Recipient</th>
                            <th data-sort="subject">Subject</th>
                            <th data-sort="event_id">Event</th>
                            <th data-sort="source">Source</th>
                            <th data-sort="directionality">Direction</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                    </tbody>
                </table>
            </div>
            
            <div class="pagination">
                <div class="pagination-info">
                    Showing <span id="showingStart">0</span> to <span id="showingEnd">0</span> of <span id="totalFiltered">0</span> entries
                </div>
                <div class="pagination-controls">
                    <button class="btn btn-secondary" id="prevBtn" onclick="previousPage()">Previous</button>
                    <span id="pageInfo" style="padding: 8px 12px;">Page 1</span>
                    <button class="btn btn-secondary" id="nextBtn" onclick="nextPage()">Next</button>
                </div>
            </div>
        </div>
        
        <div id="journey-section" class="journey-section">
            <div class="journey-header">
                <h2>&#128269; Message Journey Tracker</h2>
                <div class="journey-search">
                    <input type="text" id="journeySearchInput" placeholder="Enter Message ID or Network Message ID...">
                    <button class="btn btn-primary" onclick="searchMessageJourney()">&#128270; Track Message</button>
                    <button class="btn btn-secondary" onclick="clearJourneySearch()">&#10005; Clear</button>
                </div>
            </div>
            <div class="journey-results" id="journeyResults">
                <div class="journey-empty">
                    <div class="journey-empty-icon">&#128488;</div>
                    <h3>Track Email Journey</h3>
                    <p>Enter a Message ID or Network Message ID above to see the complete routing path and status of an email.</p>
                </div>
            </div>
        </div>
        
        <div id="compliance-section" class="compliance-section">
            <div class="compliance-header">
                <h2>&#128274; Compliance Investigation</h2>
                <p>Data Loss Prevention, Sensitive Information Types (SIT), and Sensitivity Labels detected in messages</p>
            </div>
            <div class="compliance-tabs">
                <button class="compliance-tab active" onclick="switchComplianceTab('dlpRules', this)">
                    <span class="compliance-tab-icon">&#128737;</span>
                    Data Loss Prevention
                    <span class="compliance-tab-count" id="dlpRulesCount">0</span>
                </button>
                <button class="compliance-tab" onclick="switchComplianceTab('dlp', this)">
                    <span class="compliance-tab-icon">&#128373;</span>
                    Sensitive Information Types
                    <span class="compliance-tab-count" id="dlpCount">0</span>
                </button>
                <button class="compliance-tab" onclick="switchComplianceTab('labels', this)">
                    <span class="compliance-tab-icon">&#127991;</span>
                    Sensitivity Labels
                    <span class="compliance-tab-count" id="labelsCount">0</span>
                </button>
            </div>
            <div class="compliance-content">
                <div id="dlpRulesPanel" class="compliance-panel active">
                    <div class="compliance-summary" id="dlpRulesSummary"></div>
                    <div class="section-header" style="background:#fff;padding:16px 20px;">
                        <div class="filters">
                            <input type="text" class="filter-input" id="dlpRulesFilterAll" placeholder="&#128269; Search all fields..." oninput="filterDLPRulesTable()" style="min-width:250px;">
                            <input type="text" class="filter-input" id="dlpRulesFilterSubject" placeholder="Filter by Subject..." oninput="filterDLPRulesTable()">
                            <input type="text" class="filter-input" id="dlpRulesFilterSender" placeholder="Filter by Sender..." oninput="filterDLPRulesTable()">
                            <select class="filter-select" id="dlpRulesFilterMatch" onchange="filterDLPRulesTable()">
                                <option value="">All Rules</option>
                                <option value="matched">Matched (Has Actions)</option>
                                <option value="notmatched">Not Matched</option>
                            </select>
                        </div>
                        <div class="export-buttons">
                            <button class="btn btn-secondary" onclick="resetDLPRulesFilters()">&#10005; Reset</button>
                            <button class="btn btn-primary" onclick="exportDLPRulesToCSV()">&#128190; Export CSV</button>
                        </div>
                    </div>
                    <div class="compliance-table-container">
                        <table class="compliance-table">
                            <thead>
                                <tr>
                                    <th style="width:30px"></th>
                                    <th>Subject</th>
                                    <th>Sender</th>
                                    <th>Recipient</th>
                                    <th>Rule ID</th>
                                    <th>Status</th>
                                    <th>Actions</th>
                                    <th>Predicates</th>
                                    <th>Date/Time</th>
                                </tr>
                            </thead>
                            <tbody id="dlpRulesTableBody"></tbody>
                        </table>
                    </div>
                    <div class="pagination">
                        <div class="pagination-info">
                            Showing <span id="dlpRulesShowingStart">0</span> to <span id="dlpRulesShowingEnd">0</span> of <span id="dlpRulesTotalFiltered">0</span> entries
                        </div>
                        <div class="pagination-controls">
                            <button class="btn btn-secondary" id="dlpRulesPrevBtn" onclick="dlpRulesPreviousPage()">Previous</button>
                            <span id="dlpRulesPageInfo" style="padding: 8px 12px;">Page 1</span>
                            <button class="btn btn-secondary" id="dlpRulesNextBtn" onclick="dlpRulesNextPage()">Next</button>
                        </div>
                    </div>
                </div>
                <div id="dlpPanel" class="compliance-panel">
                    <div class="compliance-summary" id="dlpSummary"></div>
                    <div class="section-header" style="background:#fff;padding:16px 20px;">
                        <div class="filters">
                            <input type="text" class="filter-input" id="dlpFilterAll" placeholder="&#128269; Search all fields..." oninput="filterDLPTable()" style="min-width:250px;">
                            <input type="text" class="filter-input" id="dlpFilterSubject" placeholder="Filter by Subject..." oninput="filterDLPTable()">
                            <input type="text" class="filter-input" id="dlpFilterSender" placeholder="Filter by Sender..." oninput="filterDLPTable()">
                            <select class="filter-select" id="dlpFilterType" onchange="filterDLPTable()">
                                <option value="">All Types</option>
                                <option value="DLP Data Classification">DLP Data Classification</option>
                                <option value="Server Side Auto Labeling">Server Side Auto Labeling</option>
                            </select>
                        </div>
                        <div class="export-buttons">
                            <button class="btn btn-secondary" onclick="resetDLPFilters()">&#10005; Reset</button>
                            <button class="btn btn-primary" onclick="exportDLPToCSV()">&#128190; Export CSV</button>
                        </div>
                    </div>
                    <div class="compliance-table-container">
                        <table class="compliance-table">
                            <thead>
                                <tr>
                                    <th style="width:30px"></th>
                                    <th>Subject</th>
                                    <th>Sender</th>
                                    <th>Recipient</th>
                                    <th>Rule ID</th>
                                    <th>Type</th>
                                    <th>Predicate</th>
                                    <th>Date/Time</th>
                                </tr>
                            </thead>
                            <tbody id="dlpTableBody"></tbody>
                        </table>
                    </div>
                    <div class="pagination">
                        <div class="pagination-info">
                            Showing <span id="dlpShowingStart">0</span> to <span id="dlpShowingEnd">0</span> of <span id="dlpTotalFiltered">0</span> entries
                        </div>
                        <div class="pagination-controls">
                            <button class="btn btn-secondary" id="dlpPrevBtn" onclick="dlpPreviousPage()">Previous</button>
                            <span id="dlpPageInfo" style="padding: 8px 12px;">Page 1</span>
                            <button class="btn btn-secondary" id="dlpNextBtn" onclick="dlpNextPage()">Next</button>
                        </div>
                    </div>
                </div>
                <div id="labelsPanel" class="compliance-panel">
                    <div class="compliance-summary" id="labelsSummary"></div>
                    <div class="section-header" style="background:#fff;padding:16px 20px;">
                        <div class="filters">
                            <input type="text" class="filter-input" id="labelsFilterAll" placeholder="&#128269; Search all fields..." oninput="filterLabelsTable()" style="min-width:250px;">
                            <input type="text" class="filter-input" id="labelsFilterSubject" placeholder="Filter by Subject..." oninput="filterLabelsTable()">
                            <input type="text" class="filter-input" id="labelsFilterSender" placeholder="Filter by Sender..." oninput="filterLabelsTable()">
                            <select class="filter-select" id="labelsFilterType" onchange="filterLabelsTable()">
                                <option value="">All Label Types</option>
                                <option value="Custom Label">Custom Label</option>
                                <option value="default">Default Labels</option>
                            </select>
                        </div>
                        <div class="export-buttons">
                            <button class="btn btn-secondary" onclick="resetLabelsFilters()">&#10005; Reset</button>
                            <button class="btn btn-primary" onclick="exportLabelsToCSV()">&#128190; Export CSV</button>
                        </div>
                    </div>
                    <div class="compliance-table-container">
                        <table class="compliance-table">
                            <thead>
                                <tr>
                                    <th style="width:30px"></th>
                                    <th>Subject</th>
                                    <th>Sender</th>
                                    <th>Recipient</th>
                                    <th>Label ID</th>
                                    <th>Label Type</th>
                                    <th>Content Bits</th>
                                    <th>Date/Time</th>
                                </tr>
                            </thead>
                            <tbody id="labelsTableBody"></tbody>
                        </table>
                    </div>
                    <div class="pagination">
                        <div class="pagination-info">
                            Showing <span id="labelsShowingStart">0</span> to <span id="labelsShowingEnd">0</span> of <span id="labelsTotalFiltered">0</span> entries
                        </div>
                        <div class="pagination-controls">
                            <button class="btn btn-secondary" id="labelsPrevBtn" onclick="labelsPreviousPage()">Previous</button>
                            <span id="labelsPageInfo" style="padding: 8px 12px;">Page 1</span>
                            <button class="btn btn-secondary" id="labelsNextBtn" onclick="labelsNextPage()">Next</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Detail Modal -->
    <div class="modal" id="detailModal">
        <div class="modal-content">
            <div class="modal-header" id="modalHeader">
            </div>
            <div class="modal-body" id="modalBody">
            </div>
        </div>
    </div>
    
    <script>
        // Data
        let rawData = [];
        let eventCounts = [];
        let directionCounts = [];
        let sourceCounts = [];
        let topSenders = [];
        let topRecipients = [];
        let hourlyCounts = [];
        let sizeDistribution = [];
        let sclDistribution = [];
        let dailyCounts = [];
        let weekdayCounts = [];
        let topSenderDomains = [];
        let threatOverview = [];
        let authFailures = [];
        let hourlyFailures = [];
        let dailyDirectionTrend = [];
        let quarantineReasons = [];
        
        // Helper to ensure array
        function ensureArray(data) {
            if (!data) return [];
            if (Array.isArray(data)) return data;
            return [data];
        }
        
        try {
            rawData = ensureArray(%%JSONDATA%%);
            eventCounts = ensureArray(%%EVENTCOUNTS%%);
            directionCounts = ensureArray(%%DIRECTIONCOUNTS%%);
            sourceCounts = ensureArray(%%SOURCECOUNTS%%);
            topSenders = ensureArray(%%TOPSENDERS%%);
            topRecipients = ensureArray(%%TOPRECIPIENTS%%);
            hourlyCounts = ensureArray(%%HOURLYCOUNTS%%);
            sizeDistribution = ensureArray(%%SIZEDISTRIBUTION%%);
            sclDistribution = ensureArray(%%SCLDISTRIBUTION%%);
            dailyCounts = ensureArray(%%DAILYCOUNTS%%);
            weekdayCounts = ensureArray(%%WEEKDAYCOUNTS%%);
            topSenderDomains = ensureArray(%%TOPSENDERDOMAINS%%);
            threatOverview = ensureArray(%%THREATOVERVIEW%%);
            authFailures = ensureArray(%%AUTHFAILURES%%);
            hourlyFailures = ensureArray(%%HOURLYFAILURES%%);
            dailyDirectionTrend = ensureArray(%%DAILYDIRECTIONTREND%%);
            quarantineReasons = ensureArray(%%QUARANTINEREASONS%%);
            console.log('Data loaded successfully. Records:', rawData.length);
        } catch(e) {
            console.error('Error parsing data:', e);
            alert('Error loading data: ' + e.message);
        }
        
        let filteredData = Array.isArray(rawData) ? [...rawData] : [];
        let currentPage = 1;
        const pageSize = 10;
        let sortColumn = 'date_time';
        let sortDirection = 'desc';
        
        // Initialize
        document.addEventListener('DOMContentLoaded', function() {
            try {
                console.log('Initializing with', rawData.length, 'records');
                if (!rawData || rawData.length === 0) {
                    console.warn('No data loaded');
                    document.getElementById('tableBody').innerHTML = '<tr><td colspan="8" style="text-align:center;">No data available</td></tr>';
                    return;
                }
                populateFilters();
                renderTable();
                setupEventListeners();
                initializeCharts();
                initializeCompliance();
                console.log('Initialization complete');
            } catch(e) {
                console.error('Initialization error:', e);
                alert('Error initializing: ' + e.message);
            }
        });
        
        // Chart color palette
        const chartColors = [
            '#0078d4', '#106ebe', '#107c10', '#ff8c00', '#d13438',
            '#8764b8', '#00b7c3', '#e3008c', '#4f6bed', '#69797e',
            '#ca5010', '#498205', '#038387', '#8e562e', '#847545'
        ];
        
        function createChart(id, type, labels, datasets, opts) {
            const el = document.getElementById(id); if (!el) return;
            opts = opts || {};
            const baseOpts = {responsive:true,maintainAspectRatio:false};
            const scales = type==='bar'||type==='line' ? {[opts.horizontal?'x':'y']:{beginAtZero:true}} : {};
            const legend = (type==='doughnut'||type==='pie') ? {position:'bottom',labels:{padding:15}} : {display:false};
            // Add percentage tooltip for doughnut/pie charts
            const tooltipCallback = (type==='doughnut'||type==='pie') ? {
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const total = context.dataset.data.reduce((a,b) => a + b, 0);
                            const value = context.raw;
                            const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                            return context.label + ': ' + value.toLocaleString() + ' (' + percentage + '%)';
                        }
                    }
                }
            } : {};
            const finalOpts = {...baseOpts, plugins:{legend,...tooltipCallback,...(opts.plugins||{})}, scales, ...(opts.horizontal?{indexAxis:'y'}:{})};
            new Chart(el.getContext('2d'), {type, data:{labels, datasets}, options:finalOpts});
        }
        function createChartWithPercentage(id, type, labels, datasets) {
            const el = document.getElementById(id); if (!el) return;
            const ctx = el.getContext('2d');
            const data = datasets[0].data;
            const total = data.reduce((a, b) => a + b, 0);
            
            // Create labels with percentages
            const labelsWithPct = labels.map((label, i) => {
                const pct = total > 0 ? ((data[i] / total) * 100).toFixed(1) : 0;
                return label + ' (' + pct + '%)';
            });
            
            new Chart(ctx, {
                type: type,
                data: { labels: labelsWithPct, datasets: datasets },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        legend: {
                            position: 'bottom',
                            labels: { padding: 15 }
                        },
                        tooltip: {
                            callbacks: {
                                label: function(context) {
                                    const value = context.raw;
                                    const pct = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
                                    return value.toLocaleString() + ' (' + pct + '%)';
                                }
                            }
                        }
                    }
                }
            });
        }
        function initializeCharts() {
            // Event, Direction, Source charts
            createChart('eventChart','bar',eventCounts.map(e=>e.Event),[{label:'Count',data:eventCounts.map(e=>e.Count),backgroundColor:chartColors.slice(0,eventCounts.length),borderWidth:1}]);
            createChartWithPercentage('directionChart','doughnut',directionCounts.map(d=>d.Direction),[{data:directionCounts.map(d=>d.Count),backgroundColor:['#0078d4','#107c10','#ff8c00','#d13438'],borderWidth:2,borderColor:'#fff'}]);
            createChartWithPercentage('sourceChart','pie',sourceCounts.map(s=>s.Source),[{data:sourceCounts.map(s=>s.Count),backgroundColor:chartColors.slice(0,sourceCounts.length),borderWidth:2,borderColor:'#fff'}]);
            // Top Senders/Recipients (horizontal bars)
            createChart('senderChart','bar',topSenders.slice(0,10).map(s=>truncateEmail(s.Sender)),[{label:'Messages Sent',data:topSenders.slice(0,10).map(s=>s.Count),backgroundColor:'#0078d4',borderColor:'#106ebe',borderWidth:1}],{horizontal:true,plugins:{legend:{display:false},tooltip:{callbacks:{title:c=>topSenders[c[0].dataIndex].Sender}}}});
            createChart('recipientChart','bar',topRecipients.slice(0,10).map(r=>truncateEmail(r.Recipient)),[{label:'Messages Received',data:topRecipients.slice(0,10).map(r=>r.Count),backgroundColor:'#107c10',borderColor:'#0e6b0e',borderWidth:1}],{horizontal:true,plugins:{legend:{display:false},tooltip:{callbacks:{title:c=>topRecipients[c[0].dataIndex].Recipient}}}});
            // Delivery Status
            const dc=rawData.filter(r=>r.event_id&&r.event_id.toUpperCase().includes('DELIVER')).length, fc=rawData.filter(r=>r.event_id&&/FAIL|DEFER/i.test(r.event_id)).length;
            createChartWithPercentage('deliveryChart','doughnut',['Delivered','Failed/Deferred','Other'],[{data:[dc,fc,rawData.length-dc-fc],backgroundColor:['#107c10','#d13438','#ff8c00'],borderWidth:2,borderColor:'#fff'}]);
            // Threat & Auth charts
            createChartWithPercentage('threatChart','doughnut',threatOverview.map(t=>t.Category),[{data:threatOverview.map(t=>t.Count),backgroundColor:['#107c10','#d13438','#ff8c00','#8764b8','#00b7c3'],borderWidth:2,borderColor:'#fff'}]);
            createChart('authChart','bar',authFailures.map(a=>a.Type),[{label:'Count',data:authFailures.map(a=>a.Count),backgroundColor:['#d13438','#ff8c00'],borderWidth:1}]);
            // SCL Distribution
            createChart('sclChart','bar',sclDistribution.map(s=>'SCL '+s.SCL),[{label:'Count',data:sclDistribution.map(s=>s.Count),backgroundColor:sclDistribution.map(s=>s.SCL<5?'#107c10':s.SCL<7?'#ff8c00':'#d13438'),borderWidth:1}]);
            // Daily & Weekday charts
            createChart('dailyChart','line',dailyCounts.map(d=>d.Date),[{label:'Messages',data:dailyCounts.map(d=>d.Count),borderColor:'#0078d4',backgroundColor:'rgba(0,120,212,0.1)',fill:true,tension:0.3,pointRadius:3,pointBackgroundColor:'#0078d4'}]);
            createChart('weekdayChart','bar',weekdayCounts.map(w=>w.Weekday.substring(0,3)),[{label:'Messages',data:weekdayCounts.map(w=>w.Count),backgroundColor:['#ff8c00','#0078d4','#0078d4','#0078d4','#0078d4','#0078d4','#ff8c00'],borderWidth:1}]);
            // Direction Trend (multi-dataset)
            createChart('directionTrendChart','line',dailyDirectionTrend.map(d=>d.Date),[{label:'Inbound',data:dailyDirectionTrend.map(d=>d.Inbound),borderColor:'#0078d4',backgroundColor:'rgba(0,120,212,0.1)',fill:true,tension:0.3},{label:'Outbound',data:dailyDirectionTrend.map(d=>d.Outbound),borderColor:'#107c10',backgroundColor:'rgba(16,124,16,0.1)',fill:true,tension:0.3}],{plugins:{legend:{position:'bottom'}}});
            // Domain chart (horizontal)
            createChart('domainChart','bar',topSenderDomains.slice(0,10).map(d=>truncateEmail(d.Domain)),[{label:'Messages',data:topSenderDomains.slice(0,10).map(d=>d.Count),backgroundColor:'#8764b8',borderColor:'#6b4c9a',borderWidth:1}],{horizontal:true,plugins:{legend:{display:false},tooltip:{callbacks:{title:c=>topSenderDomains[c[0].dataIndex].Domain}}}});
            // Hourly Failures
            const hfLabels=Array.from({length:24},(_,i)=>i.toString().padStart(2,'0')+':00'), hfData=hfLabels.map((_,i)=>(hourlyFailures.find(h=>h.Hour===i)||{Count:0}).Count);
            createChart('hourlyFailuresChart','line',hfLabels,[{label:'Failed Messages',data:hfData,borderColor:'#d13438',backgroundColor:'rgba(209,52,56,0.1)',fill:true,tension:0.3,pointRadius:3,pointBackgroundColor:'#d13438'}]);
        }
        
        function searchMessageJourney() {
            const searchInput = document.getElementById('journeySearchInput').value.trim();
            const resultsDiv = document.getElementById('journeyResults');
            
            if (!searchInput) {
                resultsDiv.innerHTML = '<div class="journey-empty"><div class="journey-empty-icon">&#128488;</div><h3>Track Email Journey</h3><p>Enter a Message ID or Network Message ID above to see the complete routing path and status of an email.</p></div>';
                return;
            }
            
            // Search by message_id or network_message_id
            const matchingMessages = rawData.filter(msg => 
                (msg.message_id && msg.message_id.toLowerCase().includes(searchInput.toLowerCase())) ||
                (msg.network_message_id && msg.network_message_id.toLowerCase().includes(searchInput.toLowerCase()))
            );
            
            if (matchingMessages.length === 0) {
                resultsDiv.innerHTML = '<div class="journey-no-results"><div class="journey-empty-icon">&#128533;</div><h3>No Messages Found</h3><p>No messages found matching the provided ID. Please check the ID and try again.</p></div>';
                return;
            }
            
            // Group messages by network_message_id or message_id
            const groupedMessages = {};
            matchingMessages.forEach(msg => {
                const groupKey = msg.network_message_id || msg.message_id || 'unknown';
                if (!groupedMessages[groupKey]) {
                    groupedMessages[groupKey] = [];
                }
                groupedMessages[groupKey].push(msg);
            });
            
            // Sort each group by date
            Object.keys(groupedMessages).forEach(key => {
                groupedMessages[key].sort((a, b) => new Date(a.date_time) - new Date(b.date_time));
            });
            
            // Build HTML
            let html = '';
            Object.keys(groupedMessages).forEach((groupKey, groupIndex) => {
                const messages = groupedMessages[groupKey];
                const firstMsg = messages[0];
                const lastMsg = messages[messages.length - 1];
                const finalStatus = getJourneyFinalStatus(messages);
                const duration = getJourneyDuration(messages);
                
                html += '<div class="journey-card" style="animation-delay:' + (groupIndex * 0.1) + 's">';
                
                // Header
                html += '<div class="journey-card-header">';
                html += '<div class="journey-card-title">&#128231; ' + escapeHtml(firstMsg.subject || 'No Subject') + '</div>';
                html += '<div class="journey-card-meta">';
                html += '<span>&#128100; ' + escapeHtml(truncateEmail(firstMsg.sender) || 'Unknown') + '</span>';
                html += '<span>&#128101; ' + escapeHtml(truncateEmail(firstMsg.recipient) || 'Unknown') + '</span>';
                html += '<span>&#128337; ' + messages.length + ' events</span>';
                html += '</div>';
                html += '</div>';
                
                // Summary Stats
                html += '<div class="journey-summary">';
                html += '<div class="journey-summary-item"><span class="journey-summary-value">' + finalStatus.icon + '</span><span class="journey-summary-label">' + finalStatus.text + '</span></div>';
                html += '<div class="journey-summary-item"><span class="journey-summary-value">' + messages.length + '</span><span class="journey-summary-label">Total Events</span></div>';
                html += '<div class="journey-summary-item"><span class="journey-summary-value">' + duration + '</span><span class="journey-summary-label">Duration</span></div>';
                html += '<div class="journey-summary-item"><span class="journey-summary-value">' + formatBytes(parseInt(firstMsg.total_bytes) || 0) + '</span><span class="journey-summary-label">Size</span></div>';
                html += '</div>';
                
                // Visual Flow
                html += '<div class="journey-flow">';
                html += '<div class="journey-flow-visual">';
                messages.forEach((msg, idx) => {
                    const eventClass = getJourneyEventClass(msg.event_id);
                    const eventIcon = getJourneyEventIcon(msg.event_id);
                    html += '<div class="journey-flow-node ' + eventClass + '" onclick="highlightStep(\'' + groupKey + '\', ' + idx + ')" title="' + escapeHtml(msg.event_id) + ' - Click to view details">';
                    html += '<span class="journey-flow-node-icon">' + eventIcon + '</span>';
                    html += '<span class="journey-flow-node-label">' + escapeHtml((msg.event_id || 'Event').substring(0, 8)) + '</span>';
                    html += '</div>';
                    if (idx < messages.length - 1) {
                        html += '<span class="journey-flow-arrow">&#10132;</span>';
                    }
                });
                html += '</div>';
                html += '</div>';
                
                // Timeline
                html += '<div class="journey-timeline">';
                html += '<div class="journey-timeline-line"></div>';
                messages.forEach((msg, index) => {
                    const eventClass = getJourneyEventClass(msg.event_id);
                    const eventIcon = getJourneyEventIcon(msg.event_id);
                    const stepId = 'step-' + groupKey.replace(/[^a-zA-Z0-9]/g, '') + '-' + index;
                    
                    html += '<div class="journey-step" id="' + stepId + '">';
                    html += '<div class="journey-step-dot ' + eventClass + '" title="' + escapeHtml(msg.event_id) + '">' + eventIcon + '</div>';
                    html += '<div class="journey-step-content" onclick="toggleStepDetails(this)">';
                    html += '<div class="journey-step-header">';
                    html += '<span class="journey-step-event ' + eventClass + '">' + eventIcon + ' ' + escapeHtml(msg.event_id || 'Unknown Event') + '</span>';
                    html += '<span class="journey-step-time">&#128337; ' + formatJourneyDate(msg.date_time) + '</span>';
                    html += '</div>';
                    
                    // Quick info row (always visible)
                    html += '<div style="display:flex;gap:16px;flex-wrap:wrap;margin-top:8px;font-size:0.85rem;">';
                    if (msg.sender) html += '<span style="display:flex;align-items:center;gap:4px;color:var(--text-muted);">&#128100; <strong style="color:var(--text-color)">' + escapeHtml(truncateEmail(msg.sender)) + '</strong></span>';
                    if (msg.recipient) html += '<span style="display:flex;align-items:center;gap:4px;color:var(--text-muted);">&#128101; <strong style="color:var(--text-color)">' + escapeHtml(truncateEmail(msg.recipient)) + '</strong></span>';
                    if (msg.directionality) html += '<span style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:12px;font-size:0.75rem;font-weight:600;background:' + (msg.directionality.toLowerCase().includes('incoming') || msg.directionality.toLowerCase().includes('inbound') ? '#deecf9;color:#0078d4' : '#fff4ce;color:#d29200') + ';">' + (msg.directionality.toLowerCase().includes('incoming') || msg.directionality.toLowerCase().includes('inbound') ? '&#128229;' : '&#128228;') + ' ' + escapeHtml(msg.directionality) + '</span>';
                    html += '</div>';
                    
                    // Expandable details
                    html += '<div class="journey-step-details" id="details-' + stepId + '">';
                    if (msg.sender) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#128100; Sender</div><div class="journey-step-detail-value">' + escapeHtml(msg.sender) + '</div></div>';
                    if (msg.recipient) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#128101; Recipient</div><div class="journey-step-detail-value">' + escapeHtml(msg.recipient) + '</div></div>';
                    if (msg.directionality) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#8644; Direction</div><div class="journey-step-detail-value">' + escapeHtml(msg.directionality) + '</div></div>';
                    if (msg.source) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#128640; Source</div><div class="journey-step-detail-value">' + escapeHtml(msg.source) + '</div></div>';
                    if (msg.server_hostname) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#128421; Server</div><div class="journey-step-detail-value">' + escapeHtml(msg.server_hostname) + '</div></div>';
                    if (msg.client_ip) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#127760; Client IP</div><div class="journey-step-detail-value">' + escapeHtml(msg.client_ip) + '</div></div>';
                    if (msg.recipient_status) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#128221; Status</div><div class="journey-step-detail-value">' + escapeHtml(msg.recipient_status) + '</div></div>';
                    if (msg.source_context) html += '<div class="journey-step-detail"><div class="journey-step-detail-label">&#128196; Context</div><div class="journey-step-detail-value">' + escapeHtml(msg.source_context) + '</div></div>';
                    html += '</div>';
                    
                    html += '<div class="journey-step-expand">&#128269; Click to view details</div>';
                    html += '</div>';
                    html += '</div>';
                });
                html += '</div>';
                
                // IDs section
                html += '<div class="journey-ids">';
                html += '<div class="journey-ids-grid">';
                html += '<div class="journey-id-item"><div class="journey-id-label">&#128273; Message ID</div><div class="journey-id-value" onclick="copyJourneyId(this, \'' + escapeHtml(firstMsg.message_id || 'N/A').replace(/'/g, "\\'") + '\')"><span>' + escapeHtml(firstMsg.message_id || 'N/A') + '</span><span class="journey-id-copy">&#128203; Copy</span></div></div>';
                html += '<div class="journey-id-item"><div class="journey-id-label">&#128279; Network Message ID</div><div class="journey-id-value" onclick="copyJourneyId(this, \'' + escapeHtml(firstMsg.network_message_id || 'N/A').replace(/'/g, "\\'") + '\')"><span>' + escapeHtml(firstMsg.network_message_id || 'N/A') + '</span><span class="journey-id-copy">&#128203; Copy</span></div></div>';
                html += '</div>';
                html += '</div>';
                
                html += '</div>';
            });
            
            resultsDiv.innerHTML = html;
        }
        
        function toggleStepDetails(element) {
            const details = element.querySelector('.journey-step-details');
            const expand = element.querySelector('.journey-step-expand');
            if (details) {
                details.classList.toggle('show');
                element.classList.toggle('expanded');
                if (expand) {
                    expand.innerHTML = details.classList.contains('show') ? '&#128317; Click to hide details' : '&#128269; Click to view details';
                }
            }
        }
        
        function highlightStep(groupKey, index) {
            const stepId = 'step-' + groupKey.replace(/[^a-zA-Z0-9]/g, '') + '-' + index;
            const step = document.getElementById(stepId);
            if (step) {
                // Remove previous highlights
                document.querySelectorAll('.journey-step-content.expanded').forEach(el => {
                    if (!el.closest('#' + stepId)) {
                        el.classList.remove('expanded');
                        const details = el.querySelector('.journey-step-details');
                        if (details) details.classList.remove('show');
                    }
                });
                
                // Scroll to and highlight
                step.scrollIntoView({ behavior: 'smooth', block: 'center' });
                const content = step.querySelector('.journey-step-content');
                if (content) {
                    content.classList.add('expanded');
                    const details = content.querySelector('.journey-step-details');
                    if (details) details.classList.add('show');
                    const expand = content.querySelector('.journey-step-expand');
                    if (expand) expand.innerHTML = '&#128317; Click to hide details';
                }
                
                // Flash effect
                step.style.transform = 'scale(1.02)';
                setTimeout(() => { step.style.transform = ''; }, 300);
            }
        }
        
        function copyJourneyId(element, text) {
            navigator.clipboard.writeText(text).then(() => {
                const copySpan = element.querySelector('.journey-id-copy');
                const originalText = copySpan.innerHTML;
                copySpan.innerHTML = '&#10004; Copied!';
                copySpan.style.color = 'var(--success-color)';
                setTimeout(() => {
                    copySpan.innerHTML = originalText;
                    copySpan.style.color = '';
                }, 1500);
            });
        }
        
        function getJourneyFinalStatus(messages) {
            const lastEvent = messages[messages.length - 1].event_id || '';
            const upper = lastEvent.toUpperCase();
            if (upper.includes('DELIVER')) return { icon: '&#10004;', text: 'Delivered', class: 'deliver' };
            if (upper.includes('FAIL') || upper.includes('DROP')) return { icon: '&#10060;', text: 'Failed', class: 'fail' };
            if (upper.includes('DEFER')) return { icon: '&#9888;', text: 'Deferred', class: 'defer' };
            if (upper.includes('RECEIVE')) return { icon: '&#128229;', text: 'Received', class: 'receive' };
            if (upper.includes('SEND')) return { icon: '&#128228;', text: 'Sent', class: 'send' };
            return { icon: '&#9679;', text: 'In Transit', class: 'default' };
        }
        
        function getJourneyDuration(messages) {
            if (messages.length < 2) return '< 1s';
            try {
                const start = new Date(messages[0].date_time);
                const end = new Date(messages[messages.length - 1].date_time);
                const diffMs = end - start;
                if (diffMs < 1000) return '< 1s';
                if (diffMs < 60000) return Math.round(diffMs / 1000) + 's';
                if (diffMs < 3600000) return Math.round(diffMs / 60000) + 'm';
                return Math.round(diffMs / 3600000) + 'h';
            } catch(e) {
                return 'N/A';
            }
        }
        
        function clearJourneySearch() {
            document.getElementById('journeySearchInput').value = '';
            document.getElementById('journeyResults').innerHTML = '<div class="journey-empty"><div class="journey-empty-icon">&#128488;</div><h3>Track Email Journey</h3><p>Enter a Message ID or Network Message ID above to see the complete routing path and status of an email.</p></div>';
        }
        
        // Compliance Investigation Functions
        let allDLPRulesData = [];
        let filteredDLPRulesData = [];
        let dlpRulesCurrentPage = 1;
        const dlpRulesPageSize = 10;
        
        let allDLPData = [];
        let filteredDLPData = [];
        let dlpCurrentPage = 1;
        const dlpPageSize = 10;
        
        let allLabelsData = [];
        let filteredLabelsData = [];
        let labelsCurrentPage = 1;
        const labelsPageSize = 10;
        
        function initializeCompliance() {
            allDLPRulesData = parseAllDLPRulesData();
            allDLPData = parseAllDLPData();
            allLabelsData = parseAllLabelData();
            
            document.getElementById('dlpRulesCount').textContent = allDLPRulesData.length;
            document.getElementById('dlpCount').textContent = allDLPData.length;
            document.getElementById('labelsCount').textContent = allLabelsData.length;
            
            renderDLPRulesSummary(allDLPRulesData);
            renderDLPSummary(allDLPData);
            renderLabelsSummary(allLabelsData);
            
            filterDLPRulesTable();
            filterDLPTable();
            filterLabelsTable();
        }
        
        // Data Loss Prevention Tab Functions
        function parseAllDLPRulesData() {
            const results = [];
            rawData.forEach((msg, idx) => {
                if (!msg.custom_data) return;
                const ruleMatches = parseDLPRulesFromCustomData(msg.custom_data);
                ruleMatches.forEach(rule => {
                    results.push({
                        index: idx,
                        subject: msg.subject || 'No Subject',
                        sender: msg.sender || 'Unknown',
                        recipient: msg.recipient || 'Unknown',
                        dateTime: msg.date_time,
                        ruleId: rule.ruleId,
                        mgtRuleId: rule.mgtRuleId,
                        policyId: rule.policyId,
                        timestamp: rule.timestamp,
                        predicates: rule.predicates,
                        actions: rule.actions,
                        matched: rule.actions.length > 0,
                        totalTimeSpent: rule.totalTimeSpent,
                        raw: rule.raw
                    });
                });
            });
            return results;
        }
        
        function parseDLPRulesFromCustomData(customData) {
            const results = [];
            // Match full S:DPA=DPR entries - they can span until next S: or semicolon/end
            // Format: S:DPA=DPR|ruleId=<GUID>|[mgtRuleId=<GUID>|][policyId=<GUID>|]st=<timestamp>|predicate=...|timeSpent=...|[action=...|timeSpent=...]...
            // Use a more robust regex that captures until the next S: section or semicolon
            const dprEntries = customData.match(/S:DPA=DPR\|[^;]*(?:;|$)/g) || [];
            
            dprEntries.forEach(entry => {
                // Extract ruleId
                const ruleIdMatch = entry.match(/ruleId=([a-f0-9-]+)/i);
                const mgtRuleIdMatch = entry.match(/mgtRuleId=([a-f0-9-]+)/i);
                const policyIdMatch = entry.match(/policyId=([a-f0-9-]+)/i);
                const stMatch = entry.match(/st=([^|;]+)/);
                
                // Extract all predicates (predicate=Name|timeSpent=N)
                const predicates = [];
                const predicateRegex = /predicate=([^|;]+)\|timeSpent=(-?\d+)/g;
                let predMatch;
                while ((predMatch = predicateRegex.exec(entry)) !== null) {
                    // Skip AndCondition as it's just a container
                    if (predMatch[1] !== 'AndCondition') {
                        predicates.push({
                            name: predMatch[1],
                            timeSpent: parseInt(predMatch[2])
                        });
                    }
                }
                
                // Extract all actions (action=Name|timeSpent=N)
                const actions = [];
                const actionRegex = /action=([^|;]+)\|timeSpent=(-?\d+)/g;
                let actMatch;
                while ((actMatch = actionRegex.exec(entry)) !== null) {
                    actions.push({
                        name: actMatch[1],
                        timeSpent: parseInt(actMatch[2])  // -1 means executed later
                    });
                }
                
                // Calculate total time spent
                let totalTimeSpent = 0;
                predicates.forEach(p => { if (p.timeSpent > 0) totalTimeSpent += p.timeSpent; });
                actions.forEach(a => { if (a.timeSpent > 0) totalTimeSpent += a.timeSpent; });
                
                if (ruleIdMatch) {
                    results.push({
                        ruleId: ruleIdMatch[1],
                        mgtRuleId: mgtRuleIdMatch ? mgtRuleIdMatch[1] : '',
                        policyId: policyIdMatch ? policyIdMatch[1] : '',
                        timestamp: stMatch ? stMatch[1] : '',
                        predicates: predicates,
                        actions: actions,
                        totalTimeSpent: totalTimeSpent,
                        raw: entry.trim()
                    });
                }
            });
            return results;
        }
        
        function filterDLPRulesTable() {
            const allFilter = document.getElementById('dlpRulesFilterAll').value.toLowerCase();
            const subjectFilter = document.getElementById('dlpRulesFilterSubject').value.toLowerCase();
            const senderFilter = document.getElementById('dlpRulesFilterSender').value.toLowerCase();
            const matchFilter = document.getElementById('dlpRulesFilterMatch').value;
            
            filteredDLPRulesData = allDLPRulesData.filter(function(item) {
                const predicatesStr = item.predicates.map(p => p.name).join(' ').toLowerCase();
                const actionsStr = item.actions.map(a => a.name).join(' ').toLowerCase();
                
                const matchAll = !allFilter || (
                    (item.subject && item.subject.toLowerCase().includes(allFilter)) ||
                    (item.sender && item.sender.toLowerCase().includes(allFilter)) ||
                    (item.recipient && item.recipient.toLowerCase().includes(allFilter)) ||
                    (item.ruleId && item.ruleId.toLowerCase().includes(allFilter)) ||
                    (item.policyId && item.policyId.toLowerCase().includes(allFilter)) ||
                    predicatesStr.includes(allFilter) ||
                    actionsStr.includes(allFilter) ||
                    (item.dateTime && item.dateTime.toLowerCase().includes(allFilter))
                );
                const matchSubject = !subjectFilter || (item.subject && item.subject.toLowerCase().includes(subjectFilter));
                const matchSender = !senderFilter || (item.sender && item.sender.toLowerCase().includes(senderFilter));
                let matchStatus = true;
                if (matchFilter === 'matched') {
                    matchStatus = item.matched === true;
                } else if (matchFilter === 'notmatched') {
                    matchStatus = item.matched === false;
                }
                return matchAll && matchSubject && matchSender && matchStatus;
            });
            
            dlpRulesCurrentPage = 1;
            renderDLPRulesTable();
            updateDLPRulesPagination();
        }
        
        function resetDLPRulesFilters() {
            document.getElementById('dlpRulesFilterAll').value = '';
            document.getElementById('dlpRulesFilterSubject').value = '';
            document.getElementById('dlpRulesFilterSender').value = '';
            document.getElementById('dlpRulesFilterMatch').value = '';
            filterDLPRulesTable();
        }
        
        function dlpRulesPreviousPage() {
            if (dlpRulesCurrentPage > 1) {
                dlpRulesCurrentPage--;
                renderDLPRulesTable();
                updateDLPRulesPagination();
            }
        }
        
        function dlpRulesNextPage() {
            const totalPages = Math.ceil(filteredDLPRulesData.length / dlpRulesPageSize);
            if (dlpRulesCurrentPage < totalPages) {
                dlpRulesCurrentPage++;
                renderDLPRulesTable();
                updateDLPRulesPagination();
            }
        }
        
        function updateDLPRulesPagination() {
            const totalPages = Math.ceil(filteredDLPRulesData.length / dlpRulesPageSize) || 1;
            const start = filteredDLPRulesData.length === 0 ? 0 : (dlpRulesCurrentPage - 1) * dlpRulesPageSize + 1;
            const end = Math.min(dlpRulesCurrentPage * dlpRulesPageSize, filteredDLPRulesData.length);
            
            document.getElementById('dlpRulesShowingStart').textContent = start;
            document.getElementById('dlpRulesShowingEnd').textContent = end;
            document.getElementById('dlpRulesTotalFiltered').textContent = filteredDLPRulesData.length;
            document.getElementById('dlpRulesPageInfo').textContent = 'Page ' + dlpRulesCurrentPage + ' of ' + totalPages;
            document.getElementById('dlpRulesPrevBtn').disabled = dlpRulesCurrentPage <= 1;
            document.getElementById('dlpRulesNextBtn').disabled = dlpRulesCurrentPage >= totalPages;
        }
        
        function exportDLPRulesToCSV() {
            if (filteredDLPRulesData.length === 0) {
                alert('No data to export');
                return;
            }
            const nl = String.fromCharCode(10);
            let csv = 'Subject,Sender,Recipient,Rule ID,Policy ID,Status,Actions,Predicates,Total Time (ms),Timestamp,Date/Time' + nl;
            filteredDLPRulesData.forEach(function(item) {
                const actionsStr = item.actions.map(a => expandActionName(a.name) + '(' + a.timeSpent + 'ms)').join('; ');
                const predicatesStr = item.predicates.map(p => expandPredicateName(p.name) + '(' + p.timeSpent + 'ms)').join('; ');
                csv += '"' + (item.subject || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.sender || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.recipient || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.ruleId || '') + '",';
                csv += '"' + (item.policyId || '') + '",';
                csv += '"' + (item.matched ? 'Matched' : 'Not Matched') + '",';
                csv += '"' + actionsStr.replace(/"/g, '""') + '",';
                csv += '"' + predicatesStr.replace(/"/g, '""') + '",';
                csv += '"' + (item.totalTimeSpent || '') + '",';
                csv += '"' + (item.timestamp || '') + '",';
                csv += '"' + (item.dateTime || '') + '"' + nl;
            });
            downloadCSV(csv, 'DLP_Policy_Rules_Report.csv');
        }
        
        // Consolidated rendering helpers
        let activeComplianceStatFilter = {};
        function renderSummaryStats(containerId, stats, panelType) {
            document.getElementById(containerId).innerHTML = stats.map(s => {
                const clickable = s.filter ? ' clickable' : '';
                const onclick = s.filter ? ' onclick="filterByComplianceStat(\'' + panelType + '\', \'' + s.filter + '\', this)"' : '';
                return '<div class="compliance-stat' + (s.type ? ' ' + s.type : '') + clickable + '"' + onclick + '><span class="compliance-stat-value">' + s.value + '</span><span class="compliance-stat-label">' + s.label + '</span></div>';
            }).join('');
        }
        
        function filterByComplianceStat(panelType, filterType, element) {
            const container = element.parentElement;
            container.querySelectorAll('.compliance-stat').forEach(card => card.classList.remove('active'));
            
            // If clicking the same filter, clear it
            if (activeComplianceStatFilter[panelType] === filterType) {
                activeComplianceStatFilter[panelType] = null;
                if (panelType === 'dlpRules') {
                    document.getElementById('dlpRulesFilterMatch').value = '';
                    filterDLPRulesTable();
                } else if (panelType === 'dlp') {
                    document.getElementById('dlpFilterType').value = '';
                    filterDLPTable();
                } else if (panelType === 'labels') {
                    document.getElementById('labelsFilterType').value = '';
                    filterLabelsTable();
                }
                return;
            }
            
            activeComplianceStatFilter[panelType] = filterType;
            element.classList.add('active');
            
            // Apply filter based on panel type
            if (panelType === 'dlpRules') {
                switch(filterType) {
                    case 'all':
                        filteredDLPRulesData = [...allDLPRulesData];
                        break;
                    case 'matched':
                        filteredDLPRulesData = allDLPRulesData.filter(d => d.matched === true);
                        break;
                    case 'notmatched':
                        filteredDLPRulesData = allDLPRulesData.filter(d => d.matched === false);
                        break;
                    case 'uniquerules':
                        const seenRules = {};
                        filteredDLPRulesData = allDLPRulesData.filter(d => {
                            if (d.ruleId && !seenRules[d.ruleId]) { seenRules[d.ruleId] = true; return true; }
                            return false;
                        });
                        break;
                    case 'uniquepolicies':
                        const seenPolicies = {};
                        filteredDLPRulesData = allDLPRulesData.filter(d => {
                            if (d.policyId && !seenPolicies[d.policyId]) { seenPolicies[d.policyId] = true; return true; }
                            return false;
                        });
                        break;
                    default:
                        filteredDLPRulesData = [...allDLPRulesData];
                }
                dlpRulesCurrentPage = 1;
                renderDLPRulesTable();
                updateDLPRulesPagination();
            } else if (panelType === 'dlp') {
                switch(filterType) {
                    case 'all':
                        filteredDLPData = [...allDLPData];
                        break;
                    case 'classification':
                        filteredDLPData = allDLPData.filter(d => d.type === 'DLP Data Classification');
                        break;
                    case 'autolabel':
                        filteredDLPData = allDLPData.filter(d => d.type === 'Server Side Auto Labeling');
                        break;
                    case 'sit':
                        filteredDLPData = allDLPData.filter(d => d.predicate && d.predicate.includes('SensitiveInformation'));
                        break;
                    case 'uniquerules':
                        const seenDLPRules = {};
                        filteredDLPData = allDLPData.filter(d => {
                            if (d.ruleId && !seenDLPRules[d.ruleId]) { seenDLPRules[d.ruleId] = true; return true; }
                            return false;
                        });
                        break;
                    default:
                        filteredDLPData = [...allDLPData];
                }
                dlpCurrentPage = 1;
                renderDLPTable();
                updateDLPPagination();
            } else if (panelType === 'labels') {
                switch(filterType) {
                    case 'all':
                        filteredLabelsData = [...allLabelsData];
                        break;
                    case 'labeledmessages':
                        const seenMessages = {};
                        filteredLabelsData = allLabelsData.filter(d => {
                            if (d.index && !seenMessages[d.index]) { seenMessages[d.index] = true; return true; }
                            return false;
                        });
                        break;
                    case 'uniquelabels':
                        const seenLabels = {};
                        filteredLabelsData = allLabelsData.filter(d => {
                            if (d.labelId && !seenLabels[d.labelId]) { seenLabels[d.labelId] = true; return true; }
                            return false;
                        });
                        break;
                    case 'default':
                        filteredLabelsData = allLabelsData.filter(d => d.labelType !== 'Custom Label');
                        break;
                    case 'custom':
                        filteredLabelsData = allLabelsData.filter(d => d.labelType === 'Custom Label');
                        break;
                    default:
                        filteredLabelsData = [...allLabelsData];
                }
                labelsCurrentPage = 1;
                renderLabelsTable();
                updateLabelsPagination();
            }
        }
        
        function renderComplianceTable(config) {
            const tbody = document.getElementById(config.tbodyId);
            const data = config.data;
            const page = config.page;
            const pageSize = config.pageSize;
            
            if (data.length === 0) {
                tbody.innerHTML = '<tr><td colspan="' + config.colspan + '" class="compliance-empty"><div class="compliance-empty-icon">' + config.emptyIcon + '</div><p>' + config.emptyMsg + '</p></td></tr>';
                return;
            }
            
            const start = (page - 1) * pageSize;
            const end = Math.min(start + pageSize, data.length);
            const pageData = data.slice(start, end);
            
            tbody.innerHTML = pageData.map((item, idx) => config.rowRenderer(item, start + idx)).join('');
        }
        
        function renderDLPRulesSummary(data) {
            renderSummaryStats('dlpRulesSummary', [
                {value: data.length, label: 'Total Rule Evaluations', filter: 'all'},
                {value: new Set(data.map(d => d.ruleId)).size, label: 'Unique Rules', filter: 'uniquerules'},
                {value: data.filter(d => d.matched).length, label: 'Rules Matched', type: 'success', filter: 'matched'},
                {value: data.filter(d => !d.matched).length, label: 'Rules Not Matched', type: 'warning', filter: 'notmatched'},
                {value: new Set(data.filter(d => d.policyId).map(d => d.policyId)).size, label: 'Unique Policies', filter: 'uniquepolicies'}
            ], 'dlpRules');
        }
        
        // Helper functions for table rendering
        function detailItem(icon, label, value, opts) {
            opts = opts || {};
            const style = opts.mono ? ' style="font-family:monospace;font-size:0.85rem;"' : '';
            const fullWidth = opts.full ? ' style="grid-column:1/-1;"' : '';
            return '<div class="compliance-detail-item"' + fullWidth + '><div class="compliance-detail-label">' + icon + ' ' + label + '</div><div class="compliance-detail-value"' + style + '>' + value + '</div></div>';
        }
        function listPreview(items, itemClass, maxShow, expandFn) {
            if (!items || items.length === 0) return '-';
            let html = items.slice(0, maxShow).map(i => {
                const name = i.name || i;
                const displayName = expandFn ? expandFn(name) : name;
                return '<span class="' + itemClass + '">' + displayName + '</span>';
            }).join('');
            if (items.length > maxShow) html += '<span class="content-bits-badge">+' + (items.length - maxShow) + ' more</span>';
            return html;
        }
        function tableRow(rowClass, expandId, cells, toggleFn) {
            return '<tr class="' + rowClass + '" onclick="' + toggleFn + '(\'' + expandId + '\', this)"><td><span class="' + rowClass.replace('-row','') + '-expand-icon">&#9658;</span></td>' + cells.map(c => '<td' + (c.title ? ' title="' + c.title + '"' : '') + '>' + c.html + '</td>').join('') + '</tr>';
        }
        
        function renderDLPRulesTable() {
            const tbody = document.getElementById('dlpRulesTableBody');
            if (filteredDLPRulesData.length === 0) {
                tbody.innerHTML = '<tr><td colspan="9" class="compliance-empty"><div class="compliance-empty-icon">&#128737;</div><p>No Data Loss Prevention events found matching your criteria.</p></td></tr>';
                return;
            }
            const start = (dlpRulesCurrentPage - 1) * dlpRulesPageSize;
            const end = Math.min(start + dlpRulesPageSize, filteredDLPRulesData.length);
            let html = '';
            filteredDLPRulesData.slice(start, end).forEach(function(item, idx) {
                const gIdx = start + idx;
                const sc = item.matched ? 'matched' : 'not-matched', st = item.matched ? 'Matched' : 'Not Matched', si = item.matched ? '&#10004;' : '&#10060;';
                html += tableRow('dlp-rule-detail-row', 'dlprule-' + gIdx, [
                    {html: escapeHtml(truncateText(item.subject, 35)), title: escapeHtml(item.subject)},
                    {html: escapeHtml(truncateEmail(item.sender)), title: escapeHtml(item.sender)},
                    {html: escapeHtml(truncateEmail(item.recipient)), title: escapeHtml(item.recipient)},
                    {html: '<span class="sit-id" title="' + item.ruleId + '">' + truncateText(item.ruleId, 20) + '</span>'},
                    {html: '<span class="dlp-rule-status ' + sc + '">' + si + ' ' + st + '</span>'},
                    {html: '<div class="dlp-actions-list">' + listPreview(item.actions, 'dlp-action-item', 2) + '</div>'},
                    {html: '<div class="dlp-predicates-list">' + listPreview(item.predicates, 'dlp-predicate-item', 2) + '</div>'},
                    {html: formatJourneyDate(item.dateTime)}
                ], 'toggleDLPRuleDetail');
                // Expanded detail
                html += '<tr><td colspan="9" style="padding:0;"><div class="dlp-rule-detail-content" id="dlprule-' + gIdx + '"><div class="compliance-detail-grid">';
                html += detailItem('&#128737;', 'Rule Status', '<span class="dlp-rule-status ' + sc + '">' + si + ' ' + st + '</span>');
                html += detailItem('&#128373;', 'Rule ID', item.ruleId, {mono:true});
                if (item.policyId) html += detailItem('&#128203;', 'Policy ID', item.policyId, {mono:true});
                if (item.mgtRuleId) html += detailItem('&#128203;', 'Management Rule ID', item.mgtRuleId, {mono:true});
                html += detailItem('&#9202;', 'Total Processing Time', item.totalTimeSpent + ' ms');
                if (item.timestamp) html += detailItem('&#128197;', 'Rule Timestamp', item.timestamp);
                html += detailItem('&#128100;', 'Sender', escapeHtml(item.sender));
                html += detailItem('&#128101;', 'Recipient', escapeHtml(item.recipient));
                html += '<div class="compliance-detail-item" style="grid-column:1/-1;"><div class="compliance-detail-label">&#128221; Subject</div><div class="compliance-detail-value">' + escapeHtml(item.subject) + '</div></div>';
                if (item.actions.length > 0) {
                    html += '<div class="compliance-detail-item" style="grid-column:1/-1;"><div class="compliance-detail-label">&#128736; Actions Taken (' + item.actions.length + ')</div><div class="compliance-detail-value"><div class="dlp-actions-list" style="max-height:none;">';
                    item.actions.forEach(a => { html += '<span class="dlp-action-item">' + expandActionName(a.name) + (a.timeSpent === -1 ? ' (executed later)' : ' (' + a.timeSpent + 'ms)') + '</span>'; });
                    html += '</div></div></div>';
                }
                if (item.predicates.length > 0) {
                    html += '<div class="compliance-detail-item" style="grid-column:1/-1;"><div class="compliance-detail-label">&#127919; Predicates Evaluated (' + item.predicates.length + ')</div><div class="compliance-detail-value"><div class="dlp-predicates-list" style="max-height:none;max-width:none;">';
                    item.predicates.forEach((p, pIdx) => {
                        let note = ' (' + p.timeSpent + 'ms)';
                        if (!item.matched && pIdx === item.predicates.length - 1 && item.predicates.length > 1) note += ' <span style="color:#991b1b;font-size:0.7rem;">(condition not met)</span>';
                        html += '<span class="dlp-predicate-item">' + expandPredicateName(p.name) + note + '</span>';
                    });
                    html += '</div></div></div>';
                }
                html += '<div class="compliance-detail-item" style="grid-column:1/-1;"><div class="compliance-detail-label">&#128196; Raw Data</div><div class="compliance-detail-value" style="font-family:monospace;font-size:0.75rem;background:#2d2d2d;color:#f0f0f0;padding:10px;border-radius:6px;word-break:break-all;">' + escapeHtml(item.raw) + '</div></div>';
                html += '</div></div></td></tr>';
            });
            tbody.innerHTML = html;
        }
        
        function toggleDLPRuleDetail(id, row) {
            const detail = document.getElementById(id);
            const isShown = detail.classList.contains('show');
            
            // Close all other details
            document.querySelectorAll('.dlp-rule-detail-content.show').forEach(d => d.classList.remove('show'));
            document.querySelectorAll('.dlp-rule-detail-row.expanded').forEach(r => r.classList.remove('expanded'));
            
            if (!isShown) {
                detail.classList.add('show');
                row.classList.add('expanded');
            }
        }
        
        // SIT/DLP Tab Functions
        function filterDLPTable() {
            const allFilter = document.getElementById('dlpFilterAll').value.toLowerCase();
            const subjectFilter = document.getElementById('dlpFilterSubject').value.toLowerCase();
            const senderFilter = document.getElementById('dlpFilterSender').value.toLowerCase();
            const typeFilter = document.getElementById('dlpFilterType').value;
            
            filteredDLPData = allDLPData.filter(function(item) {
                const matchAll = !allFilter || (
                    (item.subject && item.subject.toLowerCase().includes(allFilter)) ||
                    (item.sender && item.sender.toLowerCase().includes(allFilter)) ||
                    (item.recipient && item.recipient.toLowerCase().includes(allFilter)) ||
                    (item.ruleId && item.ruleId.toLowerCase().includes(allFilter)) ||
                    (item.mgtRuleId && item.mgtRuleId.toLowerCase().includes(allFilter)) ||
                    (item.type && item.type.toLowerCase().includes(allFilter)) ||
                    (item.predicate && item.predicate.toLowerCase().includes(allFilter)) ||
                    (item.dateTime && item.dateTime.toLowerCase().includes(allFilter))
                );
                const matchSubject = !subjectFilter || (item.subject && item.subject.toLowerCase().includes(subjectFilter));
                const matchSender = !senderFilter || (item.sender && item.sender.toLowerCase().includes(senderFilter));
                const matchType = !typeFilter || item.type === typeFilter;
                return matchAll && matchSubject && matchSender && matchType;
            });
            
            dlpCurrentPage = 1;
            renderDLPTable();
            updateDLPPagination();
        }
        
        function resetDLPFilters() {
            document.getElementById('dlpFilterAll').value = '';
            document.getElementById('dlpFilterSubject').value = '';
            document.getElementById('dlpFilterSender').value = '';
            document.getElementById('dlpFilterType').value = '';
            filterDLPTable();
        }
        
        function dlpPreviousPage() {
            if (dlpCurrentPage > 1) {
                dlpCurrentPage--;
                renderDLPTable();
                updateDLPPagination();
            }
        }
        
        function dlpNextPage() {
            const totalPages = Math.ceil(filteredDLPData.length / dlpPageSize);
            if (dlpCurrentPage < totalPages) {
                dlpCurrentPage++;
                renderDLPTable();
                updateDLPPagination();
            }
        }
        
        function updateDLPPagination() {
            const totalPages = Math.ceil(filteredDLPData.length / dlpPageSize) || 1;
            const start = filteredDLPData.length === 0 ? 0 : (dlpCurrentPage - 1) * dlpPageSize + 1;
            const end = Math.min(dlpCurrentPage * dlpPageSize, filteredDLPData.length);
            
            document.getElementById('dlpShowingStart').textContent = start;
            document.getElementById('dlpShowingEnd').textContent = end;
            document.getElementById('dlpTotalFiltered').textContent = filteredDLPData.length;
            document.getElementById('dlpPageInfo').textContent = 'Page ' + dlpCurrentPage + ' of ' + totalPages;
            document.getElementById('dlpPrevBtn').disabled = dlpCurrentPage <= 1;
            document.getElementById('dlpNextBtn').disabled = dlpCurrentPage >= totalPages;
        }
        
        function exportDLPToCSV() {
            if (filteredDLPData.length === 0) {
                alert('No data to export');
                return;
            }
            const nl = String.fromCharCode(10);
            let csv = 'Subject,Sender,Recipient,Rule ID,Mgt Rule ID,Type,Predicate,Time Spent (ms),Rule Timestamp,Date/Time' + nl;
            filteredDLPData.forEach(function(item) {
                csv += '"' + (item.subject || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.sender || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.recipient || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.ruleId || '') + '",';
                csv += '"' + (item.mgtRuleId || '') + '",';
                csv += '"' + (item.type || '') + '",';
                csv += '"' + expandPredicateName(item.predicate || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.timeSpent || '') + '",';
                csv += '"' + (item.timestamp || '') + '",';
                csv += '"' + (item.dateTime || '') + '"' + nl;
            });
            downloadCSV(csv, 'DLP_SIT_Report.csv');
        }
        
        function filterLabelsTable() {
            const allFilter = document.getElementById('labelsFilterAll').value.toLowerCase();
            const subjectFilter = document.getElementById('labelsFilterSubject').value.toLowerCase();
            const senderFilter = document.getElementById('labelsFilterSender').value.toLowerCase();
            const typeFilter = document.getElementById('labelsFilterType').value;
            
            filteredLabelsData = allLabelsData.filter(function(item) {
                const matchAll = !allFilter || (
                    (item.subject && item.subject.toLowerCase().includes(allFilter)) ||
                    (item.sender && item.sender.toLowerCase().includes(allFilter)) ||
                    (item.recipient && item.recipient.toLowerCase().includes(allFilter)) ||
                    (item.labelId && item.labelId.toLowerCase().includes(allFilter)) ||
                    (item.labelType && item.labelType.toLowerCase().includes(allFilter)) ||
                    (item.dateTime && item.dateTime.toLowerCase().includes(allFilter))
                );
                const matchSubject = !subjectFilter || (item.subject && item.subject.toLowerCase().includes(subjectFilter));
                const matchSender = !senderFilter || (item.sender && item.sender.toLowerCase().includes(senderFilter));
                let matchType = true;
                if (typeFilter === 'Custom Label') {
                    matchType = item.labelType === 'Custom Label';
                } else if (typeFilter === 'default') {
                    matchType = item.labelType !== 'Custom Label';
                }
                return matchAll && matchSubject && matchSender && matchType;
            });
            
            labelsCurrentPage = 1;
            renderLabelsTable();
            updateLabelsPagination();
        }
        
        function resetLabelsFilters() {
            document.getElementById('labelsFilterAll').value = '';
            document.getElementById('labelsFilterSubject').value = '';
            document.getElementById('labelsFilterSender').value = '';
            document.getElementById('labelsFilterType').value = '';
            filterLabelsTable();
        }
        
        function labelsPreviousPage() {
            if (labelsCurrentPage > 1) {
                labelsCurrentPage--;
                renderLabelsTable();
                updateLabelsPagination();
            }
        }
        
        function labelsNextPage() {
            const totalPages = Math.ceil(filteredLabelsData.length / labelsPageSize);
            if (labelsCurrentPage < totalPages) {
                labelsCurrentPage++;
                renderLabelsTable();
                updateLabelsPagination();
            }
        }
        
        function updateLabelsPagination() {
            const totalPages = Math.ceil(filteredLabelsData.length / labelsPageSize) || 1;
            const start = filteredLabelsData.length === 0 ? 0 : (labelsCurrentPage - 1) * labelsPageSize + 1;
            const end = Math.min(labelsCurrentPage * labelsPageSize, filteredLabelsData.length);
            
            document.getElementById('labelsShowingStart').textContent = start;
            document.getElementById('labelsShowingEnd').textContent = end;
            document.getElementById('labelsTotalFiltered').textContent = filteredLabelsData.length;
            document.getElementById('labelsPageInfo').textContent = 'Page ' + labelsCurrentPage + ' of ' + totalPages;
            document.getElementById('labelsPrevBtn').disabled = labelsCurrentPage <= 1;
            document.getElementById('labelsNextBtn').disabled = labelsCurrentPage >= totalPages;
        }
        
        function exportLabelsToCSV() {
            if (filteredLabelsData.length === 0) {
                alert('No data to export');
                return;
            }
            const nl = String.fromCharCode(10);
            let csv = 'Subject,Sender,Recipient,Label ID,Label Type,Content Bits,Content Bits Actions,Date/Time' + nl;
            filteredLabelsData.forEach(function(item) {
                const cbInfo = decodeContentBits(item.contentBits);
                csv += '"' + (item.subject || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.sender || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.recipient || '').replace(/"/g, '""') + '",';
                csv += '"' + (item.labelId || '') + '",';
                csv += '"' + (item.labelType || '') + '",';
                csv += '"' + (cbInfo.value || '') + '",';
                csv += '"' + cbInfo.actions.join('; ') + '",';
                csv += '"' + (item.dateTime || '') + '"' + nl;
            });
            downloadCSV(csv, 'Sensitivity_Labels_Report.csv');
        }
        
        function downloadCSV(csv, filename) {
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            link.click();
        }
        
        function parseAllDLPData() {
            const results = [];
            rawData.forEach((msg, idx) => {
                if (!msg.custom_data) return;
                const dlpMatches = parseDLPFromCustomData(msg.custom_data);
                dlpMatches.forEach(dlp => {
                    results.push({
                        index: idx,
                        subject: msg.subject || 'No Subject',
                        sender: msg.sender || 'Unknown',
                        recipient: msg.recipient || 'Unknown',
                        dateTime: msg.date_time,
                        ruleId: dlp.ruleId,
                        mgtRuleId: dlp.mgtRuleId || '',
                        predicate: dlp.predicate,
                        timeSpent: dlp.timeSpent,
                        timestamp: dlp.timestamp,
                        type: dlp.type,
                        raw: dlp.raw
                    });
                });
            });
            return results;
        }
        
        function parseDLPFromCustomData(customData) {
            const results = [];
            
            // Match S:DPA=DA (DLP Data Classification) entries with dcid and confidence
            const daRegex = /S:DPA=DA\|([^;]+)/gi;
            let match;
            while ((match = daRegex.exec(customData)) !== null) {
                const entry = match[1];
                const dcidMatch = entry.match(/dcid=([a-f0-9-]+)/i);
                const ruleIdMatch = entry.match(/ruleId=([a-f0-9-]+)/i);
                const confMatch = entry.match(/confidence=(\d+)/i);
                const sevMatch = entry.match(/sev=(\w+)/i);
                const stMatch = entry.match(/st=([^|]+)/);
                results.push({
                    type: 'DLP Data Classification',
                    ruleId: ruleIdMatch ? ruleIdMatch[1] : '',
                    dcid: dcidMatch ? dcidMatch[1] : '',
                    confidence: confMatch ? parseInt(confMatch[1]) : 0,
                    severity: sevMatch ? sevMatch[1] : '',
                    timestamp: stMatch ? stMatch[1] : '',
                    predicate: 'DLP Detection',
                    timeSpent: 0,
                    raw: match[0]
                });
            }
            // Match S:DPA=DPR (DLP Data Classification) entries
            const dprRegex = /S:DPA=DPR\|ruleId=([a-f0-9-]+)\|st=([^|]+)\|predicate=([^|]+)\|timeSpent=(\d+)(?:\|predicate=([^|]+)\|timeSpent=(\d+))?;?/gi;
            while ((match = dprRegex.exec(customData)) !== null) {
                results.push({
                    type: 'DLP Data Classification',
                    ruleId: match[1],
                    dcid: '',
                    confidence: 0,
                    severity: '',
                    timestamp: match[2],
                    predicate: match[3] + (match[5] ? ', ' + match[5] : ''),
                    timeSpent: parseInt(match[4]) + (match[6] ? parseInt(match[6]) : 0),
                    raw: match[0]
                });
            }
            // Match S:MLA=MLR (Server Side Auto Labeling) entries with SIT detection
            const mlaRegex = /S:MLA=MLR\|ruleId=([a-f0-9-]+)\|mgtRuleId=([a-f0-9-]+)\|st=([^|]+)\|predicate=([^|]+)\|timeSpent=(\d+)(?:\|predicate=([^|]+)\|timeSpent=(\d+))?;?/gi;
            while ((match = mlaRegex.exec(customData)) !== null) {
                results.push({
                    type: 'Server Side Auto Labeling',
                    ruleId: match[1],
                    mgtRuleId: match[2],
                    dcid: '',
                    confidence: 0,
                    severity: '',
                    timestamp: match[3],
                    predicate: match[4] + (match[6] ? ', ' + match[6] : ''),
                    timeSpent: parseInt(match[5]) + (match[7] ? parseInt(match[7]) : 0),
                    raw: match[0]
                });
            }
            return results;
        }
        
        function getPredicateIcon(predicate) {
            if (predicate.includes('SensitiveInformation')) return { icon: '&#128373;', class: 'predicate-sit', name: 'Sensitive Info' };
            if (predicate.includes('EEP')) return { icon: '&#128274;', class: 'predicate-eep', name: 'EEP' };
            if (predicate.includes('ECCSIP')) return { icon: '&#128203;', class: 'predicate-eccsip', name: 'ECCSIP' };
            if (predicate.includes('AndCondition')) return { icon: '&#128279;', class: 'predicate-and', name: 'Compound' };
            return { icon: '&#128196;', class: 'predicate-other', name: predicate.substring(0, 20) };
        }
        
        function expandPredicateName(predicate) {
            if (!predicate) return predicate;
            const predicateMap = {
                // Exchange DLP Predicates
                'ECCSIP': 'ECCSIP (ExContentContainsSensitiveInformationPredicate)',
                'EEP': 'EEP (ExEqualPredicate)',
                'EIECR': 'EIECR (ExIsEncryptionChangeRequested)',
                'EIERR': 'EIERR (ExIsEncryptionRemoveRequested)',
                'EILCR': 'EILCR (ExIsLabelChangeRequested)',
                'EIMOCGP': 'EIMOCGP (ExIsMemberOfCustomGroupsPredicate)',
                'EIMOP': 'EIMOP (ExIsMemberOfPredicate)',
                'EIP': 'EIP (ExIsPredicate)',
                'ENBASP': 'ENBASP (ExNonBifurcatingAccessScopePredicate)',
                'ERADAP': 'ERADAP (ExRecipientADAttributePredicate)',
                'ETSP': 'ETSP (ExTextScanPredicate)',
                // Classification Predicates
                'CPE': 'CPE (CcsiPredicateEvaluators)',
                'CC': 'CC (ClassificationConfigurations)',
                'ALP': 'ALP (ApplyLabelPredicate)',
                'AOP': 'AOP (AuditOperationsPredicate)',
                'CCP': 'CCP (ContainsClassificationPredicate)',
                'CMLP': 'CMLP (ContainsMachineLearningPredicate)',
                'CCSIP': 'CCSIP (ContentContainsSensitiveInformationPredicate)',
                'CMCP': 'CMCP (ContentMetadataContainsPredicate)',
                'EP': 'EP (EqualPredicate/ExistsPredicate)',
                'GTOEP': 'GTOEP (GreaterThanOrEqualPredicate)',
                'GTP': 'GTP (GreaterThanPredicate)',
                'IAP': 'IAP (IsAllPredicate)',
                'IEP': 'IEP (IsEmptyPredicate)',
                'IMOCGP': 'IMOCGP (IsMemberOfCustomGroupsPredicate)',
                'IMOP': 'IMOP (IsMemberOfPredicate)',
                'IP': 'IP (IsPredicate)',
                'LTOEP': 'LTOEP (LessThanOrEqualPredicate)',
                'LTP': 'LTP (LessThanPredicate)',
                'NVPCP': 'NVPCP (NameValuesPairConfigurationPredicate)',
                'NEP': 'NEP (NotEqualPredicate/NotExistsPredicate)',
                'NMP': 'NMP (NumericMatchPredicate)',
                'PC': 'PC (PredicateCondition)',
                'PCC': 'PCC (PredicateConditionCommon)',
                'PTSP': 'PTSP (ProtectionTextScanPredicate)',
                'QP': 'QP (QueryPredicate)',
                'TQP': 'TQP (TextQueryPredicate)',
                'TSP': 'TSP (TextScanPredicate)',
                // Other common predicates
                'AC': 'AC (AndCondition)',
                'AndCondition': 'AndCondition (Compound Condition)',
                'SensitiveInformationType': 'SensitiveInformationType (Sensitive Information Type Detection)',
                'ContentContainsSensitiveInformation': 'ContentContainsSensitiveInformation (Content Contains SIT)',
                'FromMemberOf': 'FromMemberOf (Sender Group Membership)',
                'SentTo': 'SentTo (Recipient Check)',
                'SubjectContainsWords': 'SubjectContainsWords (Subject Keyword Match)'
            };
            // Exact match for single predicates
            if (predicateMap[predicate]) {
                return predicateMap[predicate];
            }
            // If predicate contains multiple parts or comma-separated values
            let expanded = predicate;
            Object.keys(predicateMap).forEach(function(key) {
                if (expanded.indexOf(key) !== -1) {
                    expanded = expanded.split(key).join(predicateMap[key]);
                }
            });
            return expanded;
        }
        
        function expandActionName(action) {
            if (!action) return action;
            const actionMap = {
                // Exchange/Transport DLP Actions
                'BA': 'BA (BlockAccess)',
                'EAMARE': 'EAMARE (ExAddManagerAsRecipientExecutor)',
                'EAR': 'EAR (ExAddRecipients)',
                'EAREB': 'EAREB (ExAddRecipientsExecutorBase)',
                'EATRE': 'EATRE (ExAddToRecipientsExecutor)',
                'EABT': 'EABT (ExApplyBrandingTemplate)',
                'EACM': 'EACM (ExApplyContentMarking)',
                'EAHD': 'EAHD (ExApplyHtmlDisclaimer)',
                'EBCTRE': 'EBCTRE (ExBlindCopyToRecipientsExecutor)',
                'ECTRE': 'ECTRE (ExCopyToRecipientsExecutor)',
                'EE': 'EE (ExEncrypt)',
                'EGAA': 'EGAA (ExGenerateAlertAction)',
                'ELE': 'ELE (ExLabelEncrypt)',
                'EM': 'EM (ExModerate)',
                'EMS': 'EMS (ExModifySubject)',
                'ENU': 'ENU (ExNotifyUser)',
                'EPS': 'EPS (ExPrependSubject)',
                'EQ': 'EQ (ExQuarantine)',
                'ERMT': 'ERMT (ExRedirectMessageTo)',
                'ERH': 'ERH (ExRemoveHeader)',
                'ERLE': 'ERLE (ExRemoveLabelEncryption)',
                'ERRMST': 'ERRMST (ExRemoveRMSTemplate)',
                'ESH': 'ESH (ExSetHeader)',
                'ESLA': 'ESLA (ExStampLabelAction)',
                // Generic DLP Actions
                'ARA': 'ARA (AddRecipientsAction)',
                'ACMA': 'ACMA (ApplyContentMarkingAction)',
                'ALA': 'ALA (ApplyLabelAction)',
                'AOA': 'AOA (ApplyOverrideAction)',
                'ATA': 'ATA (ApplyTagAction)',
                'BAA': 'BAA (BlockAccessAction)',
                'BRAA': 'BRAA (BrowserRestrictAccessAction)',
                'CMP': 'CMP (ContentMarkingParam)',
                'DCA': 'DCA (DisableConfigurationAction)',
                'EA': 'EA (EncryptAction)',
                'GAA': 'GAA (GenerateAlertAction)',
                'GIRA': 'GIRA (GenerateIncidentReportAction)',
                'GIR': 'GIR (GenerateIncidentReportAction)',
                'HA': 'HA (HaltAction/HoldAction)',
                'LA': 'LA (LabelAction)',
                'MRAA': 'MRAA (MipRestrictAccessAction)',
                'NOA': 'NOA (NoOpAction)',
                'NAB': 'NAB (NotifyActionBase)',
                'NUA': 'NUA (NotifyUserAction)',
                'NU': 'NU (NotifyUserAction)',
                'REA': 'REA (RetentionExpireAction)',
                'RRA': 'RRA (RetentionRecycleAction)',
                'SRA': 'SRA (SelectivelyRetroactiveAction)',
                'SLA': 'SLA (StampLabelAction)',
                'TPAFA': 'TPAFA (TriggerPowerAutomateFlowAction)'
            };
            // Exact match
            if (actionMap[action]) {
                return actionMap[action];
            }
            // Return original if no match found
            return action;
        }
        
        function decodeContentBits(bits) {
            if (bits === undefined || bits === null || bits === '') return { value: '', actions: [] };
            const num = parseInt(bits);
            if (isNaN(num)) return { value: bits, actions: ['Unknown'] };
            if (num === 0) return { value: '0', actions: ['No action applied'] };
            
            const actions = [];
            if (num & 1) actions.push('Header marking');
            if (num & 2) actions.push('Footer marking');
            if (num & 4) actions.push('Watermark');
            if (num & 8) actions.push('Encryption');
            
            return { value: num.toString(), actions: actions.length > 0 ? actions : ['Unknown (' + num + ')'] };
        }
        
        function parseAllLabelData() {
            const results = [];
            rawData.forEach((msg, idx) => {
                if (!msg.custom_data) return;
                const labelMatches = parseLabelsFromCustomData(msg.custom_data);
                labelMatches.forEach(label => {
                    results.push({
                        index: idx,
                        subject: msg.subject || 'No Subject',
                        sender: msg.sender || 'Unknown',
                        recipient: msg.recipient || 'Unknown',
                        dateTime: msg.date_time,
                        labelId: label.labelId,
                        labelType: getLabelType(label.labelId),
                        contentBits: label.contentBits || '',
                        raw: label.raw
                    });
                });
            });
            return results;
        }
        
        function parseLabelsFromCustomData(customData) {
            const results = [];
            // Match S:DPA=SL|labelId=<GUID> (Sensitivity Labels)
            const regex = /S:DPA=SL\|labelId=([a-f0-9-]+);?/gi;
            let match;
            while ((match = regex.exec(customData)) !== null) {
                results.push({
                    labelId: match[1],
                    raw: match[0]
                });
            }
            
            // Extract contentBits from S:DPA=DC entries and associate with labels
            const dcRegex = /S:DPA=DC\|labelId=([a-f0-9-]+)\|[^;]*contentBits=(\d+)/gi;
            let dcMatch;
            while ((dcMatch = dcRegex.exec(customData)) !== null) {
                // Find matching label and add contentBits
                const labelId = dcMatch[1];
                const contentBits = dcMatch[2];
                const existingLabel = results.find(r => r.labelId === labelId);
                if (existingLabel) {
                    existingLabel.contentBits = contentBits;
                } else {
                    // DC entry without corresponding SL entry - add it as a label
                    results.push({
                        labelId: labelId,
                        contentBits: contentBits,
                        raw: dcMatch[0]
                    });
                }
            }
            return results;
        }
        
        function getLabelType(labelId) {
            // Default sensitivity label GUIDs start with defa4170-0d19-0005-
            // These are built-in/default labels
            if (labelId.startsWith('defa4170-0d19-0005-')) {
                const suffix = labelId.split('-')[4];
                const labelMap = {
                    '0000': 'Personal', '0001': 'Public', '0002': 'General',
                    '0003': 'Confidential', '0004': 'Highly Confidential',
                    '0005': 'Internal', '0006': 'External', '0007': 'Restricted',
                    '0008': 'Secret', '0009': 'Top Secret', '000a': 'Protected',
                    '000b': 'Classified'
                };
                const shortSuffix = suffix ? suffix.substring(0, 4).toLowerCase() : '';
                return labelMap[shortSuffix] || 'Default Label';
            }
            return 'Custom Label';
        }
        
        function renderDLPSummary(data) {
            renderSummaryStats('dlpSummary', [
                {value: data.length, label: 'Total SIT Events', filter: 'all'},
                {value: new Set(data.map(d => d.ruleId)).size, label: 'Unique Rules', filter: 'uniquerules'},
                {value: data.filter(d => d.type === 'DLP Data Classification').length, label: 'DLP Data Classification', type: 'success', filter: 'classification'},
                {value: data.filter(d => d.type === 'Server Side Auto Labeling').length, label: 'Auto-Labeling', type: 'warning', filter: 'autolabel'},
                {value: data.filter(d => d.predicate && d.predicate.includes('SensitiveInformation')).length, label: 'SIT Detections', type: 'danger', filter: 'sit'}
            ], 'dlp');
        }
        
        function renderDLPTable() {
            const tbody = document.getElementById('dlpTableBody');
            if (filteredDLPData.length === 0) {
                tbody.innerHTML = '<tr><td colspan="8" class="compliance-empty"><div class="compliance-empty-icon">&#128373;</div><p>No DLP policy events found matching your criteria.</p></td></tr>';
                return;
            }
            const start = (dlpCurrentPage - 1) * dlpPageSize;
            let html = '';
            filteredDLPData.slice(start, start + dlpPageSize).forEach(function(item, idx) {
                const gIdx = start + idx;
                const tc = item.type === 'DLP Data Classification' ? 'confidence-high' : 'confidence-medium';
                const ti = item.type === 'DLP Data Classification' ? '&#128737;' : '&#127991;';
                html += tableRow('compliance-detail-row', 'dlp-' + gIdx, [
                    {html: escapeHtml(truncateText(item.subject, 40)), title: escapeHtml(item.subject)},
                    {html: escapeHtml(truncateEmail(item.sender)), title: escapeHtml(item.sender)},
                    {html: escapeHtml(truncateEmail(item.recipient)), title: escapeHtml(item.recipient)},
                    {html: '<span class="sit-id" title="' + item.ruleId + '">' + truncateText(item.ruleId, 20) + '</span>'},
                    {html: '<span class="confidence-badge ' + tc + '">' + ti + ' ' + item.type + '</span>'},
                    {html: '<div class="dlp-predicates-list"><span class="dlp-predicate-item">' + item.predicate + '</span></div>'},
                    {html: formatJourneyDate(item.dateTime)}
                ], 'toggleComplianceDetail');
                html += '<tr><td colspan="8" style="padding:0;"><div class="compliance-detail-content" id="dlp-' + gIdx + '"><div class="compliance-detail-grid">';
                html += detailItem('&#128737;', 'Rule Type', '<span class="confidence-badge ' + tc + '">' + item.type + '</span>');
                html += detailItem('&#128373;', 'Rule ID', item.ruleId || 'N/A', {mono:true});
                if (item.dcid) html += detailItem('&#128270;', 'DCID', item.dcid, {mono:true});
                if (item.confidence > 0) {
                    const cc = item.confidence >= 85 ? 'confidence-high' : (item.confidence >= 65 ? 'confidence-medium' : 'confidence-low');
                    html += detailItem('&#127919;', 'Confidence', '<span class="confidence-badge ' + cc + '">' + item.confidence + '%</span>');
                }
                if (item.severity) html += detailItem('&#9888;', 'Severity', item.severity);
                if (item.mgtRuleId) html += detailItem('&#128203;', 'Management Rule ID', item.mgtRuleId, {mono:true});
                html += '<div class="compliance-detail-item" style="grid-column:1/-1;"><div class="compliance-detail-label">&#127919; Predicate Evaluated</div><div class="compliance-detail-value"><div class="dlp-predicates-list" style="max-height:none;max-width:none;"><span class="dlp-predicate-item">' + expandPredicateName(item.predicate) + ' (' + item.timeSpent + 'ms)</span></div></div></div>';
                html += detailItem('&#128197;', 'Rule Timestamp', item.timestamp || 'N/A');
                html += detailItem('&#128100;', 'Sender', escapeHtml(item.sender));
                html += detailItem('&#128101;', 'Recipient', escapeHtml(item.recipient));
                html += detailItem('&#128221;', 'Subject', escapeHtml(item.subject));
                html += '<div class="compliance-detail-item" style="grid-column:1/-1;"><div class="compliance-detail-label">&#128196; Raw Data</div><div class="compliance-detail-value" style="font-family:monospace;font-size:0.8rem;background:#2d2d2d;color:#f0f0f0;padding:10px;border-radius:6px;">' + escapeHtml(item.raw) + '</div></div>';
                html += '</div></div></td></tr>';
            });
            tbody.innerHTML = html;
        }
        
        function renderLabelsSummary(data) {
            renderSummaryStats('labelsSummary', [
                {value: data.length, label: 'Total Label Events', filter: 'all'},
                {value: new Set(data.map(d => d.index)).size, label: 'Labeled Messages', filter: 'labeledmessages'},
                {value: new Set(data.map(d => d.labelId)).size, label: 'Unique Labels', type: 'success', filter: 'uniquelabels'},
                {value: data.filter(d => d.labelType !== 'Custom Label').length, label: 'Default Labels', type: 'warning', filter: 'default'},
                {value: data.filter(d => d.labelType === 'Custom Label').length, label: 'Custom Labels', type: 'danger', filter: 'custom'}
            ], 'labels');
        }
        
        function renderLabelsTable() {
            const tbody = document.getElementById('labelsTableBody');
            if (filteredLabelsData.length === 0) {
                tbody.innerHTML = '<tr><td colspan="8" class="compliance-empty"><div class="compliance-empty-icon">&#127991;</div><p>No Sensitivity Labels found matching your criteria.</p></td></tr>';
                return;
            }
            const start = (labelsCurrentPage - 1) * labelsPageSize;
            let html = '';
            filteredLabelsData.slice(start, start + labelsPageSize).forEach(function(item, idx) {
                const gIdx = start + idx;
                const lc = item.labelType === 'Custom Label' ? 'confidence-medium' : 'confidence-high';
                const cb = decodeContentBits(item.contentBits);
                html += tableRow('compliance-detail-row', 'label-' + gIdx, [
                    {html: escapeHtml(truncateText(item.subject, 40)), title: escapeHtml(item.subject)},
                    {html: escapeHtml(truncateEmail(item.sender)), title: escapeHtml(item.sender)},
                    {html: escapeHtml(truncateEmail(item.recipient)), title: escapeHtml(item.recipient)},
                    {html: '<span class="sit-id" title="' + item.labelId + '">' + truncateText(item.labelId, 20) + '</span>'},
                    {html: '<span class="label-badge ' + lc + '">&#127991; ' + item.labelType + '</span>'},
                    {html: '<span class="content-bits-badge" title="' + escapeHtml(cb.actions.join(', ')) + '">' + (cb.value || '-') + '</span>'},
                    {html: formatJourneyDate(item.dateTime)}
                ], 'toggleComplianceDetail');
                html += '<tr><td colspan="8" style="padding:0;"><div class="compliance-detail-content" id="label-' + gIdx + '"><div class="compliance-detail-grid">';
                html += detailItem('&#127991;', 'Label ID', item.labelId, {mono:true});
                html += detailItem('&#128196;', 'Label Classification', '<span class="label-badge ' + lc + '">&#127991; ' + item.labelType + '</span>');
                html += detailItem('&#128204;', 'Content Bits', cb.value ? cb.value + ' - ' + cb.actions.join(', ') : 'N/A');
                html += detailItem('&#128100;', 'Sender', escapeHtml(item.sender));
                html += detailItem('&#128101;', 'Recipient', escapeHtml(item.recipient));
                html += detailItem('&#128221;', 'Subject', escapeHtml(item.subject));
                html += '<div class="compliance-detail-item" style="grid-column:1/-1;"><div class="compliance-detail-label">&#128196; Raw Data</div><div class="compliance-detail-value" style="font-family:monospace;font-size:0.8rem;background:#2d2d2d;color:#f0f0f0;padding:10px;border-radius:6px;">' + escapeHtml(item.raw) + '</div></div>';
                html += '</div></div></td></tr>';
            });
            tbody.innerHTML = html;
        }
        
        function switchComplianceTab(tab, btn) {
            document.querySelectorAll('.compliance-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.compliance-panel').forEach(p => p.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById(tab + 'Panel').classList.add('active');
        }
        
        function toggleComplianceDetail(id, row) {
            const detail = document.getElementById(id);
            const isShown = detail.classList.contains('show');
            
            // Close all other details
            document.querySelectorAll('.compliance-detail-content.show').forEach(d => d.classList.remove('show'));
            document.querySelectorAll('.compliance-detail-row.expanded').forEach(r => r.classList.remove('expanded'));
            
            if (!isShown) {
                detail.classList.add('show');
                row.classList.add('expanded');
            }
        }
        
        function truncateText(text, maxLen) {
            if (!text) return '';
            return text.length > maxLen ? text.substring(0, maxLen) + '...' : text;
        }
        
        function getJourneyEventClass(eventId) {
            if (!eventId) return 'default';
            const e = eventId.toUpperCase();
            if (e.includes('DELIVER')) return 'deliver';
            if (e.includes('RECEIVE')) return 'receive';
            if (e.includes('SEND')) return 'send';
            if (e.includes('FAIL') || e.includes('DROP')) return 'fail';
            if (e.includes('DEFER')) return 'defer';
            return 'default';
        }
        
        function getJourneyEventIcon(eventId) {
            if (!eventId) return '&#9679;';
            const e = eventId.toUpperCase();
            if (e.includes('DELIVER')) return '&#10004;';
            if (e.includes('RECEIVE')) return '&#128229;';
            if (e.includes('SEND')) return '&#128228;';
            if (e.includes('FAIL') || e.includes('DROP')) return '&#10060;';
            if (e.includes('DEFER')) return '&#9888;';
            return '&#9679;';
        }
        
        function formatJourneyDate(dateStr) {
            if (!dateStr) return 'Unknown';
            try {
                const d = new Date(dateStr);
                return d.toLocaleString('en-US', { 
                    year: 'numeric', month: 'short', day: 'numeric',
                    hour: '2-digit', minute: '2-digit', second: '2-digit',
                    hour12: false
                });
            } catch(e) {
                return dateStr;
            }
        }
        
        function escapeHtml(text) {
            if (!text) return '';
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
        
        function scrollToSection(sectionId) {
            const section = document.getElementById(sectionId);
            if (section) {
                const navHeight = document.querySelector('.nav-menu').offsetHeight;
                const sectionTop = section.getBoundingClientRect().top + window.pageYOffset - navHeight - 20;
                window.scrollTo({ top: sectionTop, behavior: 'smooth' });
                document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
                event.target.closest('.nav-btn').classList.add('active');
            }
        }
        
        function truncateEmail(email) {
            if (!email) return 'Unknown';
            if (email.length > 25) {
                return email.substring(0, 22) + '...';
            }
            return email;
        }
        
        function populateFilters() {
            const eventSelect = document.getElementById('eventFilter');
            const uniqueEvents = [...new Set(rawData.map(d => d.event_id))].filter(e => e);
            uniqueEvents.forEach(event => {
                const option = document.createElement('option');
                option.value = event;
                option.textContent = event;
                eventSelect.appendChild(option);
            });
            const directionSelect = document.getElementById('directionFilter');
            [...new Set(rawData.map(d => d.directionality))].filter(d => d).forEach(dir => { const option = document.createElement('option'); option.value = dir; option.textContent = dir; directionSelect.appendChild(option); });
            const sourceSelect = document.getElementById('sourceFilter');
            [...new Set(rawData.map(d => d.source))].filter(s => s).sort().forEach(src => { const option = document.createElement('option'); option.value = src; option.textContent = src; sourceSelect.appendChild(option); });
        }
        function setupEventListeners() {
            document.getElementById('searchInput').addEventListener('input', debounce(filterData, 300));
            ['eventFilter','directionFilter','sourceFilter'].forEach(id => document.getElementById(id).addEventListener('change', filterData));
            document.querySelectorAll('th[data-sort]').forEach(th => {
                th.addEventListener('click', () => { const column = th.dataset.sort; if (sortColumn === column) { sortDirection = sortDirection === 'asc' ? 'desc' : 'asc'; } else { sortColumn = column; sortDirection = 'asc'; } updateSortIndicators(); sortData(); renderTable(); });
            });
        }
        function debounce(func, wait) { let timeout; return function(...args) { clearTimeout(timeout); timeout = setTimeout(() => func(...args), wait); }; }
        
        let activeStatFilter = null;
        function filterByStatCard(filterType) {
            // Remove active class from all stat cards
            document.querySelectorAll('.stat-card').forEach(card => card.classList.remove('active'));
            
            // If clicking the same filter, clear it
            if (activeStatFilter === filterType) {
                activeStatFilter = null;
                document.getElementById('searchInput').value = '';
                document.getElementById('eventFilter').value = '';
                document.getElementById('directionFilter').value = '';
                filterData();
                return;
            }
            
            activeStatFilter = filterType;
            // Add active class to clicked card
            event.currentTarget.classList.add('active');
            
            // Clear existing filters first
            document.getElementById('searchInput').value = '';
            document.getElementById('eventFilter').value = '';
            document.getElementById('directionFilter').value = '';
            document.getElementById('sourceFilter').value = '';
            
            // Apply specific filter based on type
            switch(filterType) {
                case 'all':
                    filteredData = [...rawData];
                    break;
                case 'delivered':
                    filteredData = rawData.filter(row => row.event_id && row.event_id.toUpperCase().includes('DELIVER'));
                    break;
                case 'failed':
                    filteredData = rawData.filter(row => row.event_id && (row.event_id.toUpperCase().includes('FAIL') || row.event_id.toUpperCase().includes('DEFER') || row.event_id.toUpperCase().includes('DROP') || row.event_id.toUpperCase().includes('REJECT')));
                    break;
                case 'inbound':
                    filteredData = rawData.filter(row => row.directionality && (row.directionality.toLowerCase().includes('inbound') || row.directionality.toLowerCase().includes('incoming')));
                    break;
                case 'outbound':
                    filteredData = rawData.filter(row => row.directionality && (row.directionality.toLowerCase().includes('outbound') || row.directionality.toLowerCase().includes('outgoing') || row.directionality.toLowerCase().includes('originating')));
                    break;
                case 'spam':
                    filteredData = rawData.filter(row => row.custom_data && /SFV=SPM|sfv=Spam/i.test(row.custom_data));
                    break;
                case 'highscl':
                    filteredData = rawData.filter(row => row.custom_data && /SCL=[5-9]/i.test(row.custom_data));
                    break;
                case 'bulk':
                    filteredData = rawData.filter(row => row.custom_data && /BCL=[7-9]/i.test(row.custom_data));
                    break;
                case 'phish':
                    filteredData = rawData.filter(row => row.custom_data && /PHSH|phish|SFTY=9/i.test(row.custom_data));
                    break;
                case 'dkim':
                    filteredData = rawData.filter(row => row.custom_data && /DKIM=fail|dkim=none|DKIM=none/i.test(row.custom_data));
                    break;
                case 'spf':
                    filteredData = rawData.filter(row => row.custom_data && /SPF=fail|spf=fail|SPF=softfail|SPF=none/i.test(row.custom_data));
                    break;
                case 'quarantined':
                    filteredData = rawData.filter(row => row.custom_data && /quarantine|SFV=SKQ|SFV=SKA/i.test(row.custom_data));
                    break;
                case 'malware':
                    filteredData = rawData.filter(row => row.custom_data && /SFV=CLN.*AMP|malware|SFTY=9\.22/i.test(row.custom_data));
                    break;
                case 'senders':
                    // Show unique senders - group by sender and show first message from each
                    const uniqueSenders = {};
                    rawData.forEach(row => { if (row.sender && !uniqueSenders[row.sender]) uniqueSenders[row.sender] = row; });
                    filteredData = Object.values(uniqueSenders);
                    break;
                case 'recipients':
                    // Show unique recipients - group by recipient and show first message to each
                    const uniqueRecipients = {};
                    rawData.forEach(row => { if (row.recipient && !uniqueRecipients[row.recipient]) uniqueRecipients[row.recipient] = row; });
                    filteredData = Object.values(uniqueRecipients);
                    break;
                case 'domains':
                    // Show unique domains - group by sender domain and show first message from each
                    const uniqueDomains = {};
                    rawData.forEach(row => { 
                        if (row.sender && row.sender.includes('@')) {
                            const domain = row.sender.split('@')[1];
                            if (!uniqueDomains[domain]) uniqueDomains[domain] = row;
                        }
                    });
                    filteredData = Object.values(uniqueDomains);
                    break;
                default:
                    filteredData = [...rawData];
            }
            
            currentPage = 1;
            sortData();
            renderTable();
            
            // Scroll to the data section
            document.getElementById('data-section').scrollIntoView({ behavior: 'smooth' });
        }
        
        function filterData() {
            // Clear stat card active state when using manual filters
            activeStatFilter = null;
            document.querySelectorAll('.stat-card').forEach(card => card.classList.remove('active'));
            
            const searchTerm = document.getElementById('searchInput').value.toLowerCase();
            const eventFilter = document.getElementById('eventFilter').value;
            const directionFilter = document.getElementById('directionFilter').value;
            const sourceFilter = document.getElementById('sourceFilter').value;
            filteredData = rawData.filter(row => { const matchesSearch = !searchTerm || Object.values(row).some(val => val && val.toString().toLowerCase().includes(searchTerm)); return matchesSearch && (!eventFilter || row.event_id === eventFilter) && (!directionFilter || row.directionality === directionFilter) && (!sourceFilter || row.source === sourceFilter); });
            currentPage = 1; sortData(); renderTable();
        }
        function sortData() { filteredData.sort((a, b) => { let aVal = a[sortColumn] || '', bVal = b[sortColumn] || ''; if (sortColumn === 'date_time') { aVal = new Date(aVal); bVal = new Date(bVal); } if (aVal < bVal) return sortDirection === 'asc' ? -1 : 1; if (aVal > bVal) return sortDirection === 'asc' ? 1 : -1; return 0; }); }
        function updateSortIndicators() { document.querySelectorAll('th[data-sort]').forEach(th => { th.classList.remove('sort-asc', 'sort-desc'); if (th.dataset.sort === sortColumn) th.classList.add(sortDirection === 'asc' ? 'sort-asc' : 'sort-desc'); }); }
        function renderTable() {
            const tbody = document.getElementById('tableBody'), start = (currentPage - 1) * pageSize, end = Math.min(start + pageSize, filteredData.length), pageData = filteredData.slice(start, end);
            tbody.innerHTML = pageData.map((row, index) => ``
                <tr>
                    <td>``+formatDate(row.date_time)+``</td>
                    <td title="``+escapeHtml(row.sender)+``">``+escapeHtml(row.sender)+``</td>
                    <td title="``+escapeHtml(row.recipient)+``">``+escapeHtml(truncateText(row.recipient, 50))+``</td>
                    <td title="``+escapeHtml(row.subject)+``">``+escapeHtml(truncateText(row.subject, 40))+``</td>
                    <td><span class="badge badge-``+getEventBadgeClass(row.event_id)+``">``+escapeHtml(row.event_id)+``</span></td>
                    <td>``+escapeHtml(row.source)+``</td>
                    <td><span class="badge badge-``+getDirectionBadgeClass(row.directionality)+``">``+escapeHtml(row.directionality)+``</span></td>
                    <td><button class="btn btn-primary" onclick="showDetails(``+(start + index)+``)">View</button></td>
                </tr>
            ``).join('');
            document.getElementById('showingStart').textContent = filteredData.length > 0 ? start + 1 : 0;
            document.getElementById('showingEnd').textContent = end;
            document.getElementById('totalFiltered').textContent = filteredData.length;
            document.getElementById('pageInfo').textContent = 'Page ' + currentPage + ' of ' + (Math.ceil(filteredData.length / pageSize) || 1);
            document.getElementById('prevBtn').disabled = currentPage === 1;
            document.getElementById('nextBtn').disabled = end >= filteredData.length;
        }
        function formatDate(dateStr) { if (!dateStr) return ''; try { return new Date(dateStr).toISOString().replace('T', ' ').substring(0, 19); } catch { return dateStr; } }
        function truncateText(text, maxLen) { return !text ? '' : text.length > maxLen ? text.substring(0, maxLen) + '...' : text; }
        function escapeHtml(text) { if (!text) return ''; const div = document.createElement('div'); div.textContent = text; return div.innerHTML; }
        function getEventBadgeClass(event) { if (!event) return 'incoming'; const e = event.toUpperCase(); if (e.includes('DELIVER')) return 'deliver'; if (e.includes('RECEIVE')) return 'receive'; if (e.includes('SEND')) return 'send'; if (e.includes('FAIL') || e.includes('DEFER')) return 'fail'; return 'incoming'; }
        function getDirectionBadgeClass(direction) { if (!direction) return 'incoming'; const d = direction.toLowerCase(); if (d.includes('incoming') || d.includes('inbound')) return 'incoming'; if (d.includes('outgoing') || d.includes('outbound') || d.includes('originating')) return 'outgoing'; return 'incoming'; }
        function previousPage() { if (currentPage > 1) { currentPage--; renderTable(); } }
        function nextPage() { if (currentPage * pageSize < filteredData.length) { currentPage++; renderTable(); } }
        function showDetails(index) {
            const row = filteredData[index], modal = document.getElementById('detailModal'), modalHeader = document.getElementById('modalHeader'), modalBody = document.getElementById('modalBody');
            const eventClass = getEventBadgeClass(row.event_id), directionClass = getDirectionBadgeClass(row.directionality);
            modalHeader.innerHTML = ``
                <div class="modal-header-bg event-``+eventClass+``">
                    <h3>&#9993; Message Details</h3>
                    <div class="modal-meta">
                        <div class="modal-meta-item">&#128197; ``+formatDate(row.date_time)+``</div>
                        <div class="modal-meta-item">&#9881; ``+escapeHtml(row.source)+``</div>
                        <div class="modal-meta-item">&#9889; ``+escapeHtml(row.event_id)+``</div>
                        <div class="modal-meta-item">&#8644; ``+escapeHtml(row.directionality)+``</div>
                    </div>
                </div>
                <button class="modal-close" onclick="closeModal()">&times;</button>
            ``;
            modalBody.innerHTML = ``
                <!-- Subject Section -->
                <div class="detail-section" id="section-subject">
                    <div class="detail-section-header section-subject">
                        <span class="detail-section-icon">&#128221;</span>
                        <span class="detail-section-title">Subject</span>
                        <button class="detail-toggle-btn hide-btn" id="toggle-subject" onclick="toggleSection('subject')">&#128065; Hide Details</button>
                    </div>
                    <div class="detail-section-content">
                        <div class="detail-grid">
                            <div class="detail-item full-width">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128172;</span> Message Subject</div>
                                    <button class="copy-btn" onclick="copyToClipboard('``+escapeHtml(row.subject)+``', this)">&#128203; Copy</button>
                                </div>
                                <div class="detail-value" style="font-size: 1rem; font-weight: 500;">``+escapeHtml(row.subject || 'No Subject')+``</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Participants Section -->
                <div class="detail-section" id="section-participants">
                    <div class="detail-section-header section-participants">
                        <span class="detail-section-icon">&#128101;</span>
                        <span class="detail-section-title">Participants</span>
                        <button class="detail-toggle-btn hide-btn" id="toggle-participants" onclick="toggleSection('participants')">&#128065; Hide Details</button>
                    </div>
                    <div class="detail-section-content">
                        <div class="detail-grid">
                            <div class="detail-item full-width">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128228;</span> Sender</div>
                                    <button class="copy-btn" onclick="copyToClipboard('``+escapeHtml(row.sender)+``', this)">&#128203; Copy</button>
                                </div>
                                <div class="detail-value">``+escapeHtml(row.sender)+``</div>
                            </div>
                            <div class="detail-item full-width">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128229;</span> Recipient(s)</div>
                                    <button class="copy-btn" onclick="copyToClipboard('``+escapeHtml(row.recipient)+``', this)">&#128203; Copy</button>
                                </div>
                                <div class="detail-value">``+escapeHtml(row.recipient)+``</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Message Info Section -->
                <div class="detail-section" id="section-message">
                    <div class="detail-section-header section-message">
                        <span class="detail-section-icon">&#128172;</span>
                        <span class="detail-section-title">Message Information</span>
                        <button class="detail-toggle-btn hide-btn" id="toggle-message" onclick="toggleSection('message')">&#128065; Hide Details</button>
                    </div>
                    <div class="detail-section-content">
                        <div class="detail-grid">
                            <div class="detail-item">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128202;</span> Size</div>
                                </div>
                                <div class="detail-value" style="font-size: 1.25rem; font-weight: 600; color: var(--primary-color);">``+(row.total_bytes ? formatBytes(row.total_bytes) : 'N/A')+``</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128101;</span> Recipients</div>
                                </div>
                                <div class="detail-value" style="font-size: 1.25rem; font-weight: 600; color: var(--primary-color);">``+escapeHtml(row.recipient_count || '1')+``</div>
                            </div>
                            <div class="detail-item full-width">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128273;</span> Message ID</div>
                                    <button class="copy-btn" onclick="copyToClipboard('``+escapeHtml(row.message_id)+``', this)">&#128203; Copy</button>
                                </div>
                                <div class="detail-value monospace">``+escapeHtml(row.message_id || 'N/A')+``</div>
                            </div>
                            <div class="detail-item full-width">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#127760;</span> Network Message ID</div>
                                    <button class="copy-btn" onclick="copyToClipboard('``+escapeHtml(row.network_message_id)+``', this)">&#128203; Copy</button>
                                </div>
                                <div class="detail-value monospace">``+escapeHtml(row.network_message_id || 'N/A')+``</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Delivery Status Section -->
                <div class="detail-section" id="section-delivery">
                    <div class="detail-section-header section-delivery">
                        <span class="detail-section-icon">&#128230;</span>
                        <span class="detail-section-title">Delivery Status</span>
                        <button class="detail-toggle-btn hide-btn" id="toggle-delivery" onclick="toggleSection('delivery')">&#128065; Hide Details</button>
                    </div>
                    <div class="detail-section-content">
                        <div class="detail-grid">
                            <div class="detail-item">
                                <div class="detail-label"><span class="detail-label-icon">&#9889;</span> Event</div>
                                <div class="detail-value" style="margin-top: 8px;"><span class="status-badge ``+eventClass+``">``+getEventIcon(row.event_id)+`` ``+escapeHtml(row.event_id)+``</span></div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label"><span class="detail-label-icon">&#8644;</span> Direction</div>
                                <div class="detail-value" style="margin-top: 8px;"><span class="direction-badge ``+directionClass+``">``+getDirectionIcon(row.directionality)+`` ``+escapeHtml(row.directionality)+``</span></div>
                            </div>
                            <div class="detail-item full-width">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128196;</span> Recipient Status</div>
                                </div>
                                <div class="detail-value monospace">``+escapeHtml(row.recipient_status || 'N/A')+``</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                <!-- Technical Details Section -->
                <div class="detail-section" id="section-technical">
                    <div class="detail-section-header section-technical">
                        <span class="detail-section-icon">&#9881;</span>
                        <span class="detail-section-title">Technical Details</span>
                        <button class="detail-toggle-btn hide-btn" id="toggle-technical" onclick="toggleSection('technical')">&#128065; Hide Details</button>
                    </div>
                    <div class="detail-section-content">
                        <div class="detail-grid">
                            <div class="detail-item">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#127760;</span> Client IP</div>
                                    ``+(row.client_ip || row.original_client_ip ? '<button class="copy-btn" onclick="copyToClipboard(\''+(row.client_ip || row.original_client_ip)+'\', this)">&#128203; Copy</button>' : '')+``
                                </div>
                                <div class="detail-value monospace">``+escapeHtml(row.client_ip || row.original_client_ip || 'N/A')+``</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128421;</span> Server</div>
                                    ``+(row.server_hostname ? '<button class="copy-btn" onclick="copyToClipboard(\''+escapeHtml(row.server_hostname)+'\', this)">&#128203; Copy</button>' : '')+``
                                </div>
                                <div class="detail-value monospace" style="font-size: 0.75rem;">``+escapeHtml(row.server_hostname || 'N/A')+``</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label"><span class="detail-label-icon">&#9881;</span> Source</div>
                                <div class="detail-value">``+escapeHtml(row.source || 'N/A')+``</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#127970;</span> Tenant ID</div>
                                    ``+(row.tenant_id ? '<button class="copy-btn" onclick="copyToClipboard(\''+escapeHtml(row.tenant_id)+'\', this)">&#128203; Copy</button>' : '')+``
                                </div>
                                <div class="detail-value monospace" style="font-size: 0.75rem;">``+escapeHtml(row.tenant_id || 'N/A')+``</div>
                            </div>
                            <div class="detail-item full-width">
                                <div class="detail-item-header">
                                    <div class="detail-label"><span class="detail-label-icon">&#128269;</span> Source Context</div>
                                    ``+(row.source_context ? '<button class="copy-btn" onclick="copyToClipboard(\''+escapeHtml(row.source_context)+'\', this)">&#128203; Copy</button>' : '')+``
                                </div>
                                <div class="detail-value monospace">``+escapeHtml(row.source_context || 'N/A')+``</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                ``+(row.custom_data ? formatCustomData(row.custom_data) : '')+``
                
                <div class="modal-footer">
                    <button class="btn btn-primary" onclick="closeModal()" style="padding: 10px 24px; font-size: 0.9rem;">&#10005; Close</button>
                </div>
            ``;
            modal.classList.add('active');
        }
        function toggleSection(sectionId) { const section = document.getElementById('section-' + sectionId), btn = document.getElementById('toggle-' + sectionId); section.classList.toggle('collapsed'); if (section.classList.contains('collapsed')) { btn.innerHTML = '&#128064; Show Details'; btn.classList.remove('hide-btn'); } else { btn.innerHTML = '&#128065; Hide Details'; btn.classList.add('hide-btn'); } }
        function copyToClipboard(text, btn) { navigator.clipboard.writeText(text).then(() => { const originalText = btn.innerHTML; btn.innerHTML = '&#10004; Copied!'; btn.classList.add('copied'); setTimeout(() => { btn.innerHTML = originalText; btn.classList.remove('copied'); }, 1500); }).catch(err => console.error('Copy failed:', err)); }
        function getEventIcon(event) { if (!event) return '&#9679;'; const e = event.toUpperCase(); if (e.includes('DELIVER')) return '&#10004;'; if (e.includes('RECEIVE')) return '&#128229;'; if (e.includes('SEND')) return '&#128228;'; if (e.includes('FAIL') || e.includes('DEFER')) return '&#10060;'; return '&#9679;'; }
        function getDirectionIcon(direction) { if (!direction) return '&#8644;'; const d = direction.toLowerCase(); if (d.includes('incoming') || d.includes('inbound')) return '&#128229;'; if (d.includes('outgoing') || d.includes('outbound') || d.includes('originating')) return '&#128228;'; return '&#8644;'; }
        function closeModal() { document.getElementById('detailModal').classList.remove('active'); }
        function formatBytes(bytes) { if (!bytes) return '0 Bytes'; bytes = parseInt(bytes); const k = 1024, sizes = ['Bytes', 'KB', 'MB', 'GB'], i = Math.floor(Math.log(bytes) / Math.log(k)); return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i]; }
        const customDataAcronyms = {
            'S:SFA': { name: 'Spam Filter Agent', desc: 'Spam filtering results from Exchange Online Protection' }, 'S:PCFA': { name: 'Pre-Categorization Filter Agent', desc: 'Pre-filtering actions before spam categorization' }, 'S:CompCost': { name: 'Compliance Cost', desc: 'Compliance processing cost metrics' }, 'S:DeliveryPriority': { name: 'Delivery Priority', desc: 'Message delivery priority level (Normal, Low, High)' }, 'S:PrioritizationReason': { name: 'Prioritization Reason', desc: 'Reason for message priority assignment' }, 'S:AccountForest': { name: 'Account Forest', desc: 'Exchange Online forest hosting the mailbox' }, 'S:MapiNotifyTime': { name: 'MAPI Notify Time', desc: 'Time when MAPI notification was sent' }, 'S:TenantEopForest': { name: 'Tenant EOP Forest', desc: 'Exchange Online Protection forest for the tenant' }, 'S:MailboxDatabaseGuid': { name: 'Mailbox Database GUID', desc: 'Unique identifier of the mailbox database' }, 'S:AMA': { name: 'Anti-Malware Agent', desc: 'Anti-malware scanning results' }, 'S:TRA': { name: 'Transport Rule Agent', desc: 'Transport/Mail flow rule processing results' },
            'SFV': { name: 'Spam Filter Verdict', desc: 'Final spam filtering decision' }, 'SFV=NSPM': { name: 'Not Spam', desc: 'Message was not identified as spam' }, 'SFV=SPM': { name: 'Spam', desc: 'Message was identified as spam' }, 'SFV=BLK': { name: 'Blocked', desc: 'Message was blocked by spam filter' }, 'SFV=SKS': { name: 'Skipped Scanning', desc: 'Spam scanning was skipped' }, 'SFV=SKN': { name: 'Skip - Allow List', desc: 'Skipped due to sender being on allow list' }, 'SFV=SKI': { name: 'Skip - Internal', desc: 'Skipped for internal message' }, 'SFV=SKA': { name: 'Skip - Admin Allow', desc: 'Skipped due to admin allow policy' }, 'SFV=SKB': { name: 'Skip - Block List', desc: 'Blocked due to sender being on block list' }, 'SFV=SKQ': { name: 'Skip - Quarantine', desc: 'Message sent to quarantine' },
            'SCL': { name: 'Spam Confidence Level', desc: 'Score indicating spam likelihood (0-9, higher = more likely spam)' }, 'SCL=-1': { name: 'SCL: Trusted', desc: 'Message from trusted sender, skip spam filtering' }, 'SCL=0': { name: 'SCL: Not Spam', desc: 'Scanning determined message is not spam' }, 'SCL=1': { name: 'SCL: Not Spam (Low)', desc: 'Very low spam probability' }, 'SCL=5': { name: 'SCL: Spam', desc: 'Spam filtering marked as spam' }, 'SCL=6': { name: 'SCL: Spam (High)', desc: 'High spam probability' }, 'SCL=9': { name: 'SCL: High Confidence Spam', desc: 'Very high confidence spam' }, 'BCL': { name: 'Bulk Complaint Level', desc: 'Bulk email complaint rating (0-9, higher = more complaints)' },
            'IPV': { name: 'IP Verdict', desc: 'IP reputation verdict' }, 'IPV=NLI': { name: 'Not Listed', desc: 'IP not on any block lists' }, 'IPV=CAL': { name: 'Customer Allow List', desc: 'IP on customer allow list' }, 'IPV=CBL': { name: 'Customer Block List', desc: 'IP on customer block list' }, 'DIR': { name: 'Direction', desc: 'Message direction' }, 'DIR=INB': { name: 'Inbound', desc: 'Message is inbound (incoming)' }, 'DIR=OUT': { name: 'Outbound', desc: 'Message is outbound (outgoing)' }, 'DIR=INT': { name: 'Internal', desc: 'Message is internal (within organization)' },
            'CTRY': { name: 'Country', desc: 'Country code of sending IP' }, 'CLTCTRY': { name: 'Client Country', desc: 'Country of the original client' }, 'SRV': { name: 'Server Verdict', desc: 'Server-level verdict' }, 'H': { name: 'HELO/EHLO', desc: 'Hostname from HELO/EHLO SMTP command' }, 'CIP': { name: 'Connecting IP', desc: 'IP address of connecting server' }, 'SFP': { name: 'Spam Filter Processing', desc: 'Spam filter processing indicator' }, 'ASF': { name: 'Advanced Spam Filtering', desc: 'Advanced spam filter result' }, 'LANG': { name: 'Language', desc: 'Detected message language' }, 'RD': { name: 'Reverse DNS', desc: 'Reverse DNS lookup result of connecting IP' },
            'LAT': { name: 'Latency', desc: 'Processing latency in milliseconds' }, 'MLAT': { name: 'Message Latency', desc: 'Total message processing latency' }, 'RLAT': { name: 'Routing Latency', desc: 'Routing processing latency' }, 'ALAT': { name: 'Agent Latency', desc: 'Agent processing latency' }, 'SFS': { name: 'Spam Filter Score', desc: 'Spam filter score/rule ID' }, 'SCORE': { name: 'Score', desc: 'Overall spam score' }, 'LIST': { name: 'List', desc: 'List membership indicator' }, 'DI': { name: 'Directory Info', desc: 'Directory lookup information' }, 'FPR': { name: 'Fingerprint', desc: 'Message fingerprint for tracking' },
            'tact': { name: 'Threat Action', desc: 'Threat protection action taken' }, 'tactcat': { name: 'Threat Action Category', desc: 'Category of threat action' }, 'FTBP': { name: 'File Type Block Policy', desc: 'Action taken by file type blocking policy' }, 'hctfp': { name: 'Hash Check True File Policy', desc: 'Hash-based file policy check' }, 'ETR': { name: 'Exchange Transport Rule', desc: 'Transport rule triggered' }, 'C': { name: 'Compliance', desc: 'Compliance processing indicator' }, 'SMS': { name: 'Size/Memory Status', desc: 'Message size relative to limit (bytes used/limit)' },
            'sfv': { name: 'Spam Filter Verdict', desc: 'Spam filter verdict code' }, 'rsk': { name: 'Risk Level', desc: 'Risk level assessment (Low/Medium/High)' }, 'scl': { name: 'Spam Confidence Level', desc: 'Spam confidence level value' }, 'bcl': { name: 'Bulk Complaint Level', desc: 'Bulk email complaint level' }, 'score': { name: 'Spam Score', desc: 'Calculated spam score' }, 'sfs': { name: 'Spam Filter Signatures', desc: 'Matched spam filter rule IDs' }, 'sfp': { name: 'Spam Filter Processing', desc: 'Spam filter processing flags' }, 'fprx': { name: 'Fingerprint Extended', desc: 'Extended message fingerprint' }, 'mlc': { name: 'Machine Learning Category', desc: 'ML-based spam category' }, 'mlv': { name: 'Machine Learning Verdict', desc: 'ML-based spam verdict' },
            'list': { name: 'List Match', desc: 'Safe/Block list match indicator' }, 'di': { name: 'Directory Info', desc: 'Directory lookup result' }, 'rd': { name: 'Reverse DNS', desc: 'PTR record of sender IP' }, 'h': { name: 'HELO Domain', desc: 'SMTP HELO/EHLO domain' }, 'ctry': { name: 'Country', desc: 'Originating country code' }, 'cltctry': { name: 'Client Country', desc: 'Client country code' }, 'lang': { name: 'Language', desc: 'Message language' }, 'cip': { name: 'Client IP', desc: 'Connecting client IP' }, 'dir': { name: 'Direction', desc: 'Message flow direction' }, 'alat': { name: 'Agent Latency', desc: 'Agent processing time (ms)' }, 'mlat': { name: 'Message Latency', desc: 'Message processing time (ms)' }, 'rlat': { name: 'Routing Latency', desc: 'Routing time (ms)' }, 'asf': { name: 'ASF Flags', desc: 'Advanced spam filtering flags' },
            'NotSpam': { name: 'Not Spam', desc: 'Message classified as legitimate' }, 'Spam': { name: 'Spam', desc: 'Message classified as spam' }, 'Low': { name: 'Low Risk', desc: 'Low risk level' }, 'Normal': { name: 'Normal Priority', desc: 'Normal delivery priority' }, 'High': { name: 'High Priority', desc: 'High delivery priority' }, 'Incoming': { name: 'Incoming', desc: 'Inbound message' }, 'Outgoing': { name: 'Outgoing', desc: 'Outbound message' }, 'Originating': { name: 'Originating', desc: 'Message originated from this tenant' }
        };
        function formatCustomData(customData) {
            if (!customData) return '';
            let html = '<div class="detail-section" id="section-custom"><div class="detail-section-header section-custom"><span class="detail-section-icon">&#128202;</span><span class="detail-section-title">Custom Data (Detailed Analysis)</span></div><div class="detail-section-content">';
            const sections = customData.split(';').filter(s => s.trim());
            html += '<div class="custom-data-container">';
            sections.forEach(section => {
                section = section.trim();
                if (!section) return;
                if (section.startsWith('S:')) {
                    const colonIndex = section.indexOf('='); let sectionKey, sectionValue;
                    if (colonIndex > -1) { sectionKey = section.substring(0, colonIndex); sectionValue = section.substring(colonIndex + 1); } else { sectionKey = section; sectionValue = ''; }
                    const sectionInfo = customDataAcronyms[sectionKey] || { name: sectionKey.replace('S:', ''), desc: 'Exchange Online Protection data' };
                    html += '<div class="custom-data-block"><div class="custom-data-header"><span class="custom-data-key" title="' + escapeHtml(sectionInfo.desc) + '">' + escapeHtml(sectionInfo.name) + '</span><span class="custom-data-code">(' + escapeHtml(sectionKey) + ')</span></div>';
                    if (sectionValue) {
                        if (sectionValue.includes('|')) {
                            const subParts = sectionValue.split('|'); html += '<table class="custom-data-table"><tbody>';
                            subParts.forEach(part => { const [subKey, subVal] = part.includes('=') ? part.split('=', 2) : [part, '']; const subInfo = customDataAcronyms[subKey.toLowerCase()] || customDataAcronyms[subKey] || { name: subKey, desc: '' }; const valInfo = customDataAcronyms[subVal] || customDataAcronyms[subKey + '=' + subVal] || null; html += '<tr><td class="sub-key" title="' + escapeHtml(subInfo.desc) + '">' + escapeHtml(subInfo.name) + ' <span class="code-hint">(' + escapeHtml(subKey) + ')</span></td><td class="sub-value">'; if (valInfo) { html += '<span title="' + escapeHtml(valInfo.desc) + '">' + escapeHtml(valInfo.name) + '</span>'; if (subVal && subVal !== valInfo.name) html += ' <span class="code-hint">(' + escapeHtml(subVal) + ')</span>'; } else { html += escapeHtml(subVal || '-'); } html += '</td></tr>'; });
                            html += '</tbody></table>';
                        } else { const valInfo = customDataAcronyms[sectionValue] || null; html += '<div class="custom-data-value">'; if (valInfo) { html += '<span title="' + escapeHtml(valInfo.desc) + '">' + escapeHtml(valInfo.name) + '</span> <span class="code-hint">(' + escapeHtml(sectionValue) + ')</span>'; } else { html += escapeHtml(sectionValue); } html += '</div>'; }
                    }
                    html += '</div>';
                }
            });
            html += '</div><details class="raw-data-details"><summary>View Raw Custom Data</summary><pre class="raw-data-pre">' + escapeHtml(customData) + '</pre></details></div></div>';
            return html;
        }
        function resetFilters() { document.getElementById('searchInput').value = ''; document.getElementById('eventFilter').value = ''; document.getElementById('directionFilter').value = ''; document.getElementById('sourceFilter').value = ''; filterData(); }
        function exportToCSV() {
            const headers = ['Date/Time', 'Sender', 'Recipient', 'Subject', 'Event', 'Source', 'Direction', 'Message ID'];
            const rows = filteredData.map(row => [row.date_time, row.sender, row.recipient, row.subject, row.event_id, row.source, row.directionality, row.message_id]);
            let csv = headers.join(',') + '\n';
            rows.forEach(row => { csv += row.map(cell => { if (cell && cell.toString().includes(',')) { return '"' + cell.toString().replace(/"/g, '""') + '"'; } return cell || ''; }).join(',') + '\n'; });
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' }); const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = 'filtered_message_trace.csv'; link.click();
        }
        document.getElementById('detailModal').addEventListener('click', function(e) { if (e.target === this) closeModal(); });
        document.addEventListener('keydown', function(e) { if (e.key === 'Escape') closeModal(); });
    </script>
</body>
</html>
"@

    return $htmlTemplate
}
#endregion

#region Main Script Execution
Invoke-MessageTraceReport -CsvFilePath $CsvPath -OutputFilePath $OutputPath
