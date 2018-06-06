$DC = "DCname"

$Credential = (Get-Credential)
Write-Host "Importing AD Session"
$ADSession = New-PSSession -ComputerName $DC -Credential $Credential
Invoke-Command -Command { Import-Module ActiveDirectory } -Session $ADSession
Import-PSSession -Session $ADSession -Module ActiveDirectory -AllowClobber


$RunBank = New-Object psobject @{}
$Machines = Get-ADComputer -SearchBase "OU-Path" -Filter * | Select-Object -ExpandProperty Name
$DnsSearchOrder = "1.1.1.1", "2.2.2.2"

$MassPing = $Machines | ForEach-Object { Test-Connection -ComputerName $_ -Count 1 -AsJob } | Get-Job | Receive-Job -Wait | Select-Object Address,StatusCode
#$OnlineComputerDetails = $MassPing | Where-Object StatusCode -EQ "0" | Select-Object Address,IPV4Address
$OnlineComputers = $MassPing | Where-Object StatusCode -EQ "0" | Select-Object -ExpandProperty Address
$OfflineComputers = $MassPing | Where-Object StatusCode -NE "0" | Select-Object -ExpandProperty Address

foreach ($Computer in $OnlineComputers) {
  Write-Output "trying $Computer..."
  $NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object DNSServerSearchOrder -Contains "3.3.3.3" -ErrorAction SilentlyContinue
  if ($NIC) {
    Write-Warning "Found machine $($NIC.PSComputerName) with incorrect DNS search order: $($NIC.DNSServerSearchOrder)."
    [array]RunBank += [PSCustomObject] @{
      Computer = "$($NIC.PSComputerName)"
      CurrentDNS = "$($NIC.DNSServerSearchOrder)"
    }
    Write-Warning "Setting correct DNS server addresses now"
    $NIC.SetDNSServerSearchOrder($DnsSearchOrder)
    }
}
