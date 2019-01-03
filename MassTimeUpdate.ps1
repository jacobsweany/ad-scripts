$Workstations = get-adcomputer -SearchBase "OU=Structure" -Filter * | Select-Object -ExpandProperty Name

$MassPing = $Workstations | ForEach-Object { Test-Connection -ComputerName $_ -Count -AsJob } | Get-Job | Receive-Job -Wait | Select-Object Address,StatusCode

$OnlineComputers = $MassPing | Where-Object StatusCode -EQ "0" | Select-Object -ExpandProperty Address
$OfflineComputers = $MassPing | Where-Object StatusCode -NE "0" | Select-Object -ExpandProperty Address

$TimeUpdate = Invoke-Command -ComputerName $OnlineComputers -ScriptBlock { w32tm /config /syncfromflags:domhier /update ; Restart-Service W32Time } -JobName "TimeUpdate" -ThrottleLimit 20 -AsJob
