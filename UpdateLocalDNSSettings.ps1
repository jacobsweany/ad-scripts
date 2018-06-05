$DC = "DCname"

$Credential = (Get-Credential)
Write-Host "Importing AD Session"
$ADSession = New-PSSession -ComputerName $DC -Credential $Credential
Invoke-Command -Command { Import-Module ActiveDirectory } -Session $ADSession
Import-PSSession -Session $ADSession -Module ActiveDirectory -AllowClobber

$Machines = Get-ADComputer -SearchBase "OU-Path" -Filter * | Select-Object -ExpandProperty Name
$DnsSearchOrder = "1.1.1.1", "2.2.2.2"

foreach ($Computer in $Machines) {
  Write-Output "trying $Computer..."
  $NIC = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer | Where-Object DNSServerSearchOrder -Contains "3.3.3.3" -ErrorAction SilentlyContinue
  if ($NIC) {
    Write-Output "Found machine $($NIC.PSComputerName) with incorrect DNS search order: $($NIC.DNSServerSearchOrder). Changing it to the correct order."
    $NIC.SetDNSServerSearchOrder($DnsSearchOrder)
    }
}
