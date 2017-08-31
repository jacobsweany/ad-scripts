# Get computer names from multiple OUs
# Created by Jacob Sweany 8/30/17

# Options
$IncludeOptionalOU = $false
$ClearFiles = $false

# Text files
$Workstations = "\\path\file.txt"
$WorkstationsClean = "\\path\file-clean.txt"

# Create AD PSSession
$ADSession = New-PSSession -ComputerName DCNAME
Invoke-Command -Session $ADSession { Import-Module activedirectory }
Import-PSSession -Session $ADSession -Module activedirectory -AllowClobber

# Clear files 
if ($ClearFiles) {
    Clear-Content $Workstations
    Clear-Content $WorkstationsClean
}

# Output computer names to text file. Duplicate this line for each OU as necessary.
Get-ADComputer -SearchBase 'OU=Path,DC=test,DC=com' -Filter '*' | select -Exp Name | Out-File $Workstations -Append

# Optional OU, filters only computers that start with "WS"
if ($IncludeOptionalOU) {
    Get-ADComputer -SearchBase 'OU=OptPath,DC=test,DC=com' -Filter 'samAccountName -like "WS*"' | select -Exp Name | Out-File $Workstations -Append
}

gc $Workstations | Sort-Object | Get-Unique | Set-Content $WorkstationsClean

# Remove AD session
Remove-PSSession -Session $ADSession
