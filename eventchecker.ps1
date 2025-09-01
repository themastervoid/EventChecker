# ----------------------------
# Settings
# ----------------------------
$Computers = @("PC01","PC02","PC03")
$StartTime = (Get-Date).AddDays(-3)
$EndTime   = Get-Date

$DetailFile  = "C:\Temp\RebootDetails.csv"
$SummaryFile = "C:\Temp\RebootSummary.csv"

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
        $XmlEvents = Invoke-Command -ComputerName $Computer -ScriptBlock {
            param($StartTime, $EndTime)
            # Convert to W3C format for wevtutil query
            $StartStr = $StartTime.ToString("yyyy-MM-ddTHH:mm:ss")
            $EndStr   = $EndTime.ToString("yyyy-MM-ddTHH:mm:ss")
            # Query System log for reboot events
            wevtutil qe System /q:"*[System[(EventID=1074 or EventID=1076 or EventID=6006 or EventID=6008) and TimeCreated[@SystemTime >= '$StartStr' and @SystemTime <= '$EndStr']]]" /f:xml
        } -ArgumentList $StartTime, $EndTime -ErrorAction Stop

        # Parse XML and extract fields
        $Events = foreach ($xml in $XmlEvents) {
            [xml]$doc = $xml
            $node = $doc.Event
            if ($node) {
                $eventId = [int]$node.System.EventID
                $time    = [datetime]$node.System.TimeCreated.SystemTime
                $message = $node.EventData.Data -join " "
                [PSCustomObject]@{
                    Computer    = $Computer
                    TimeCreated = $time
                    EventID     = $eventId
                    Status      = $EventMap[$eventId]
                    Message     = $message -replace "`r|`n"," "
                }
            }
        }

        $Events
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
    } else {
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

$Details | Export-Csv -Path $DetailFile -NoTypeInformation -Encoding UTF8
$Summary | Export-Csv -Path $SummaryFile -NoTypeInformation -Encoding UTF8

Write-Host "`nReports saved to:" -ForegroundColor Green
Write-Host "  Details: $DetailFile"
Write-Host "  Summary: $SummaryFile"
