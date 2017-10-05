# Test connection, shut down if online. Either way log to file

$Workstations = "\\path\file.txt"
$WorkstationsClean = "\\path\file-clean.txt"

$Online = "\\path\Online.txt"
$Offline = "\\path\Offline.txt"

$OnlineContent = Get-Content $Online
Get-Content $Workstations | Sort-Object | Get-Unique | Set-Content $WorkstationsClean

Clear-Content $Online
Clear-Content $Offline



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

# Test connection, shut down if online. Either way log to file. First pass.
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

Start-Sleep -Seconds 10
# Loop entire process until all online computers are no longer pingable
while ($OnlineContent -ne $null) {
    # Refresh input file
    $OnlineContent = Get-Content $Online
    # Test connection, shut down if online. Either way log to file. First pass.
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
}


foreach ($Computer in $Computers) {
    Stop-Computer -ComputerName $Computer -Force
}

Stop-Computer -ComputerName SERVERNAME1 -Force
Write-Output "$Computer" | Out-File $Offline -Append

while ($test = Test-Connection -ComputerName SERVERNAME1 -Count 1 -Quiet) {
    Write-Output "SERVERNAME1 is online still"
    Start-Sleep -Seconds 10
}

# Shut down local system
# Stop-Computer -Force
