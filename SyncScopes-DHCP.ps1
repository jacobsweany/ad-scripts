# Create DHCP PSSession
$DHCPSession = New-PSSession -ComputerName "DHCPSERVERNAME"
Invoke-Command -Session $DHCPSession { Import-Module DHCPServer }
Import-PSSession -Session $DHCPSession -Module DHCPServer -AllowClobber

$DHCPServers = Get-DHCPServerInDC
$RunBank = New-Object psobject @{}

foreach ($server in $DHCPServers) {
    $query = Get-DhcpServerV4Scope -ComputerName $server.DnsName |
        Select-Object Name, Description, ScopeID, SubnetMask, EndRange, State
    [array]$RunBank += $query
}

$DataOnly = $RunBank | Where-Object { $_.Name -notlike "*VoIP*" -and ( $_.Name -notlike "*Printer*") -and ($_.State -eq "Active") } | select Name, Description, ScopeID, SubnetMask, EndRange

$DataOnly | Export-Csv -Path "\\path\DHCP.csv" -Force -NoTypeInformation
