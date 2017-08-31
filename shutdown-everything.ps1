# Test connection, shut down if online. Either way log to file

$Workstations = "\\path\file.txt"
$WorkstationsClean = "\\path\file-clean.txt"

$Online = "\\path\Online.txt"
$Offline = "\\path\Offline.txt"

foreach ($Computer in $Computers) {
    $test = Test-Connection -ComputerName $Computer -Count 1 -Quiet -BufferSize 16
     if ($test) {
        Write-Output "$Computer" | Out-File $Online -Append
        Stop-Computer $Computer -Force
     }
     else {
        Write-Output "$Computer" | Out-File $Offline -Append
     }
}

Start-Sleep -Seconds 120

# Test connection, shut down if online. Either way log to file
foreach ($Computer in $Computers) {
    $test = Test-Connection -ComputerName $Computer -Count 1 -Quiet -BufferSize 16
     if ($test) {
        Write-Output "$Computer" | Out-File $Online -Append
        Stop-Computer $Computer -Force
     }
     else {
        Write-Output "$Computer" | Out-File $Offline -Append
        # Remove computer name out of Online
        (Get-Content $Online) -notmatch "$Computer" | Out-File $Online
     }
}

Start-Sleep -Seconds 120

foreach ($Computer in $Computers) {
    Stop-Computer -ComputerName $Computer -Force
}

Stop-Computer -ComputerName SERVERNAME1 -Force
Write-Output "$Computer" | Out-File $Offline -Append

while ($test = Test-Connection -ComputerName SERVERNAME1 -Count 1 -Quiet) {
    Write-Output "SERVERNAME1 is online still"
    Start-Sleep -Seconds 10
}

Stop-Computer -Force
