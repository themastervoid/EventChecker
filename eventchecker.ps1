# Look at last 3 days of system shutdown/reboot events
$StartTime = (Get-Date).AddDays(-3)
$EndTime   = Get-Date

Get-WinEvent -LogName System -FilterHashtable @{
    Id = 1074,1076,6006,6008
    StartTime = $StartTime
    EndTime   = $EndTime
} | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize
