# --------------------------------
# Last Reboot Status Check - PSSessionhttps://github.com/themastervoid/EventChecker/blob/main/eventchecker.ps1
# --------------------------------

# Time window: last 7 days (adjust as needed)
$StartTime = (Get-Date).AddDays(-7)
$EndTime   = Get-Date

# Event ID mapping
$EventMap = @{
    1074 = "Planned restart/shutdown initiated"
    1076 = "Restart/shutdown canceled"
    6006 = "Clean shutdown (normal)"
    6008 = "Unexpected shutdown (crash, power loss)"
}

# Build XPath filter for WinEvent
$XPath = @"
*[System[
    (EventID=1074 or EventID=1076 or EventID=6006 or EventID=6008) and
    TimeCreated[@SystemTime >= '$($StartTime.ToUniversalTime().ToString("o"))' and
                @SystemTime <= '$($EndTime.ToUniversalTime().ToString("o"))']
]]
"@

# Get events using FilterXPath (avoids parameter set errors in PSSession)
$Events = Get-WinEvent -LogName System -FilterXPath $XPath | Sort-Object TimeCreated -Descending

if ($Events.Count -eq 0) {
    Write-Host "No reboot/shutdown events found in the last 7 days." -ForegroundColor Yellow
} else {
    $lastEvent = $Events[0]
    $eventId   = $lastEvent.Id
    $time      = $lastEvent.TimeCreated
    $status    = $EventMap[$eventId]

    Write-Host "Last reboot/shutdown on this machine:" -ForegroundColor Cyan
    Write-Host "  Status    : $status"
    Write-Host "  Event ID  : $eventId"
    Write-Host "  Timestamp : $time"
    Write-Host "  Message   : $($lastEvent.Message -replace "`r|`n"," ")" -ForegroundColor Green
}
