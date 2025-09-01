# ----------------------------
# Settings
# ----------------------------

# Option 1: Define computers manually
$Computers = @("PC01","PC02","PC03")

# Option 2: Or load from a text file (one hostname per line)
# $Computers = Get-Content "C:\Temp\hosts.txt"

# Time window (adjust as needed)
$StartTime = (Get-Date).AddDays(-3)   # last 3 days
$EndTime   = Get-Date

# Output files
$DetailFile  = "C:\Temp\RebootDetails.csv"
$SummaryFile = "C:\Temp\RebootSummary.csv"

# Event ID to status mapping
$EventMap = @{
    1074 = "Planned restart/shutdown initiated"
    1076 = "Restart/shutdown canceled"
    6006 = "Clean shutdown (normal)"
    6008 = "Unexpected shutdown (crash, power loss)"
}

# ----------------------------
# Script
# ----------------------------

$Details = foreach ($Computer in $Computers) {
    Write-Host "Checking $Computer ..." -ForegroundColor Cyan

    try {
        $Events = Invoke-Command -ComputerName $Computer -ScriptBlock {
            param($StartTime, $EndTime)

            Get-WinEvent -LogName System -FilterHashtable @{
                Id = 1074,1076,6006,6008
                StartTime = $StartTime
                EndTime = $EndTime
            } | Select-Object TimeCreated, Id, Message

        } -ArgumentList $StartTime, $EndTime -ErrorAction Stop

        foreach ($Event in $Events) {
            [PSCustomObject]@{
                Computer    = $Computer
                TimeCreated = $Event.TimeCreated
                EventID     = $Event.Id
                Status      = $EventMap[$Event.Id]
                Message     = $Event.Message -replace "`r|`n"," "
            }
        }
    }
    catch {
        [PSCustomObject]@{
            Computer    = $Computer
            TimeCreated = ""
            EventID     = ""
            Status      = "Error"
            Message     = $_.Exception.Message
        }
    }
}


# ----------------------------
# Build Summary
# ----------------------------

$Summary = $Details | Group-Object Computer | ForEach-Object {
    $lastEvent = $_.Group | Where-Object {$_.TimeCreated} | Sort-Object TimeCreated -Descending | Select-Object -First 1

    if ($null -eq $lastEvent) {
        [PSCustomObject]@{
            Computer    = $_.Name
            LastEvent   = "None in timeframe"
            Status      = "No reboot/shutdown logged"
            TimeCreated = ""
        }
    }
    else {
        [PSCustomObject]@{
            Computer    = $_.Name
            LastEvent   = $lastEvent.Status
            Status      = $lastEvent.Status
            TimeCreated = $lastEvent.TimeCreated
        }
    }
}

# ----------------------------
# Output
# ----------------------------

Write-Host "`n=== Detailed Events ===" -ForegroundColor Yellow
$Details | Format-Table -AutoSize

Write-Host "`n=== Summary Per Machine ===" -ForegroundColor Yellow
$Summary | Format-Table -AutoSize

# Save to CSVs
$Details | Export-Csv -Path $DetailFile -NoTypeInformation -Encoding UTF8
$Summary | Export-Csv -Path $SummaryFile -NoTypeInformation -Encoding UTF8

Write-Host "`nReports saved to:" -ForegroundColor Green
Write-Host "  Details: $DetailFile"
Write-Host "  Summary: $SummaryFile"
