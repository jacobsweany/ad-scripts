# Create DHCP PSSession
$DHCPSession = New-PSSession -ComputerName "DHCPSERVERNAME"
Invoke-Command -Session $DHCPSession { Import-Module DHCPServer }
Import-PSSession -Session $DHCPSession -Module DHCPServer -AllowClobber

$DHCPServers = Get-DHCPServerInDC | Select -ExpandProperty DnsName
$Query = Get-DhcpServerv4Scope | Where-Object { $_.Name -notlike "*VoIP*" -and ( $_.Name -notlike "*Printer*") -and ($_.State -eq "Active") } | select Name, StartRange, EndRange, Description
$RunBank = New-Object psobject @{}

foreach ($server in $DHCPServers) {
    Invoke-Command -ScriptBlock {$Query}
    [array]$RunBank += $Query
}

$RunBank | select Name, StartRange, EndRange, Description | Export-Csv -Path "\\path\DHCP.csv" -Force -NoTypeInformation
