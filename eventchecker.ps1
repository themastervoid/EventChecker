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
        $XPath = @"
        *[System[
            (EventID=1074 or EventID=1076 or EventID=6006 or EventID=6008) and
            TimeCreated[@SystemTime >= '$($StartTime.ToUniversalTime().ToString("o"))' and
                        @SystemTime <= '$($EndTime.ToUniversalTime().ToString("o"))']
        ]]
"@

        $Events = Invoke-Command -ComputerName $Computer -ScriptBlock {
            param($XPath)
            Get-WinEvent -LogName System -FilterXPath $XPath
        } -ArgumentList $XPath -ErrorAction Stop

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
