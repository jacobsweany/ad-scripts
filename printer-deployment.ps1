# Printer logon script
# Script will add new network printers as needed based off IP subnet of the computer. It will delete old network printer connections as needed.

$test = Get-Printer | Where-Object ComputerName -Contains "newserver"

if (!$test) {
    Write-Host "New print server not found"
    
    $PrintServer = "\\PrintServer\"
    $NewPrinter1 = "PrinterName"
    $NewPrinter2 = "PrinterName2"
    
    # Old printers to delete
    $OldPrinter1 = "\\printserverold\PrinterName"
    $OldPrinter2 = "\\printserverold\PrinterName2"
    
    $ComputerIPAddress = Get-NetIPAddress -AddressFamily IPv4 | Where-Object -FilterScript { $_.InterfaceAlias -ne "Loopback Pseudo-Interface 1" } | Select-Object -ExpandProperty IPAddress
    
    # Determine subnet
    
    if ($ComputerIPAddress -like "10.10.10.*") {$Location = "Location 1"; $PrintShare = $NewPrinter1}
    elseif ($ComputerIPAddress -like "10.10.30.*") {$Location = "Location 2"; $PrintShare = $NewPrinter2; $PrintDel = $OldPrinter1; $PrintDel2 = $OldPrinter2}
    elseif ($ComputerIPAddress -like "10.10.30.*") {$Location = "Location 2"; $PrintShare = $NewPrinter2; $PrintDel = $OldPrinter1; $PrintDel2 = $OldPrinter2}
    
    $Printer = "$PrintServer$PrintShare$"
    
    # Add printer, make first one default
    Add-Printer -ConnectionName $Printer | Out-Null
    (Get-WmiObject -Class Win32_Printer -Filter "ShareName='$PrintShare'").SetDefaultPrinter()
    
    # Add second printer (if called for)
    if ($PrintShare2) {
        $Printer2 = "$PrintServer$PrintShare"
        Add-Printer -ConnectionName $Printer2 | Out-Null
    }
    
    # Delete printers (if called for)
    if ($PrintDel) {
        Get-Printer | Where-Object Name -EQ "$PrintDel" | Remove-Printer | Out-Null
    }
    if ($PrintDel2) {
        Get-Printer | Where-Object Name -EQ "$PrintDel2" | Remove-Printer | Out-Null
    }
}
