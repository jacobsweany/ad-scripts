### Active Directory/Exchange/Skype for Business Account Creation GUI
### Created by Jacob Sweany

# Make sure to fill out section below for your environment. Review the process everywhere in the script 
# that determines account naming convention (search for 5+2) and make sure it conforms to your naming convention.

# This script assumes that you are running from a machine that has RSAT installed with the Active Directory
# PowerShell module added, and that you are running the script with an account that has access to Active 
# Directory, Exchange and Skype for Business with account creation and modify rights.

# Set environment specific variables and server names
$ExchangeServer = "name"
$SkypeServer = "name"
$domainName = "domain.com"
$scriptPath = "script.bat"
$UserOUBase = "OU=Users,OU=anotherOU,DC=domain,DC=com"
$SkypePool = "pool.$domainName"
$ExchArchiveDB = "Archive Database"
$UMMailboxPolicy = "Policy"
$homeDrive = "Z:"
$scriptPath = "Logon.bat"

#region Active Directory initial commands

# Load AD module
try {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    Write-Host "Imported Active Directory module."
}
catch {
    # If module is not found, throw error message and exit script
    Write-Warning "Active Directory module is not found. Do you have RSAT installed?"
    Write-Warning "Terminating script because module was not found."
    exit
}

# Convert all user OUs into a table used later on for association.
# NOTE: If you want to use the SiteCode values (used for Skype LineURI unique values and associated 
# function Get-NextAvailableLineURI), then make sure to put the two digit site code as a description in each OU.
# For example, if the Los Angeles OU has a 2 digit sitecode of 23, its description would be SiteCode 23. 
# A new user who will be created in the Los Angeles OU will have a VOIP number starting with 23.
$RawOU = Get-ADOrganizationalUnit -SearchBase $UserOUBase -Filter {Name -like "* - *" } -Properties Name, Description, DistinguishedName | select Name, Description, DistinguishedName
foreach ($line in $RawOU) {
    # Get the clean name of each OU, this may or may not be needed
    $line.Name = ($line.Name -split "- ")[1].Substring(0)
    # Isolate SiteCode number
    $line.Description = ($line.Description -split "SiteCode ")[1].Substring(0)
    $line | Add-Member -MemberType NoteProperty "SiteCode" -Value $line.Description
}
$CleanOUList = $RawOU | select Name, SiteCode, DistinguishedName | sort Sitecode
# Use the Name field later for the Site dropdown on the form
[array]$DropDownArray = $CleanOUList | select -ExpandProperty Name
#endregion Active Directory initial commands

#region Other remote shells
# Import Exchange shell
cls
Write-host "Importing Exchange session."
try {
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer.$domainName/PowerShell/ -AllowRedirection -Authentication Kerberos
    Import-PSSession -Session $ExchangeSession -AllowClobber -DisableNameChecking
}
catch {
    Write-Warning "Exchange remote shell was not imported. Terminating!"
    Write-Warning $_.Exception.Message
    Start-Sleep -Seconds 5
    exit
}

# Import Skype session
Write-host "Importing Skype session."
try {
    $SkypeSession = New-PSSession -ConnectionUri "https://$SkypeServer.$domainName/ocspowershell" -Authentication NegotiateWithImplicitCredential
    Import-PSSession -Session $SkypeSession -AllowClobber
}
catch {
    Write-Warning "Skype remote shell was not imported. Terminating!"
    Write-Warning $_.Exception.Message
    Start-Sleep -Seconds 5
    exit
}
cls
#endregion


#region Functions

function CheckUser {
# Used to check if account name is in use already or not
    $name = "$($formLastName.text), $($formFirstName.text)"
    $lname = $formLastName.text
    $fname = $formFirstName.text
    # Sets the samAccountName to be the first 5 digits of last name + first 2 of first name. Change as needed
    $lnameTrunc = "$lname"[0..4] -join ""
    $fnameTrunc = "$fname"[0..1] -join ""
    # If alternate 5and2 is set in the form, save to variable
    $alt5and2 = $formSetNew5and2.Text
    if (!$alt5and2) {
        # Otherwise, use the above calulation to create 5+2 samAccountName
        $5and2 = "$lnameTrunc$fnameTrunc".ToLower()
    }
    else {
        # If alternate field isn't populated, populate it now
        $5and2 = $formSetNew5and2.Text
    }
    # Check if samAccountName is in use or not
    $ADCheck = Get-ADUser -Filter {SAMAccountName -eq $5and2} -ErrorAction SilentlyContinue
    if ($ADCheck -eq $null) {
        $formStatusBar.Text = "$($5and2) is available"
        $formSetNew5and2_Label.Visibility = "Hidden"
        if (!$alt5and2){
         # unknown use
        }
    }
    else {
        # If in use, make new 5+2 field visible on the form and notify user that account is in use
        $formSetNew5and2_Label.Visibility = "Visible"
        $formSetNew5and2.Visibility = "Visible"
        $formSetNew5and2.Text = $5and2
        $formStatusBar.Text = "$($5and2) exists already. Enter alternative 5+2:"
        [System.Windows.MessageBox]::Show("$($5and2) exists already. Enter alternative 5+2.", "Account Exists","OK","Exclamation")
        }
    if ((!$5and2) -and (!$alt5and2)) {
        $formStatusBar.Text = "Invalid account name"
        [System.Windows.MessageBox]::Show("Invalid account name. Enter valid account name.", "Invalide Account Name","OK","Exclamation")
    }    
}

function EvalAccountDetails {
# Goes through all entered fields, generates other fields as appropriate, returns them all to the form
# and passes them to the CreateUser function if the -Submit parameter was specified when running this function
    Param(
    [bool]$Submit = $false
    )
    $formStatusBar.Text = "Evaluating account details..."
    # Runs CheckUser function to make sure that the AD account info is still valid
    CheckUser
    $name = "$($formLastName.text), $($formFirstName.text)"
    $lname = $formLastName.text
    $fname = $formFirstName.text
    $password = $formPassword.Password.ToString()
    $site = $FormDropDownSite.Text
    $employeeId = $formMyID.Text
    if (($lname) -and ($fname) -and ($password) -and ($site)) {
        # Checks to make sure that all required fields have values
        $formCheckDetailsButton.Content = "Check Details"
        $lnameTrunc = "$lname"[0..4] -join ""
        $fnameTrunc = "$fname"[0..1] -join ""
        $alt5and2 = $formSetNew5and2.Text
        if (!$alt5and2) {
            $5and2 = "$lnameTrunc$fnameTrunc".ToLower()
        }
        else {
            $5and2 = $formSetNew5and2.Text
        }
        # Determine next available Skype lineURI at site location
        $siteCode = ($CleanOUList | where {$_.Name -EQ "$($site)"}).SiteCode 
        $siteOU = ($CleanOUList | where {$_.Name -EQ "$($site)"}).DistinguishedName 
        # Pass LineURI to function, which will return the next available LineURI for the given site/sitecode
        $LineURI = Get-NextAvailableLineURI -SiteCode $siteCode
        $homeDirectory = "\\$domainName\shares\users\$5and2"
        $upn = "$5and2@$domainName"
        # Determine which Exchange mailbox the user needs to be added to
        if ($site -eq "Palmdale") {$ExchDB = "Palmdale"}
        elseif (($site -eq "59th Street") -or ($site -eq "Tinker")) {$ExchDB = "OKC"}
        else {$ExchDB = "RemoteSites"}
        # Output all result data to the form. Data object is created in this function
        $ResultData = @{
            'Display Name'="$name"; 
            'First Name'="$fname";
            'Last Name'="$lname"; 
            'Account'="$5and2"; 
            'Site'="$site"; 
            'VoIP Number'="$LineURI"; 
            'Temporary Password'="$password"; 
            'UPN'="$upn"; 
            'Home Directory'="$homeDirectory"; 
            'Site OU'="$siteOU"; 
            'Exchange Database'="$ExchDB"; 
            'MyID'="$employeeId"
        }          
        $DataTable = New-Object System.Collections.ArrayList
        $DataTable.AddRange($ResultData)
        $formResultData.ItemsSource=@($DataTable)
        # Make the objects visible now
        $formResultData.Visibility = "Visible"
        $formStatusBar.Text = "Generated account details"
        if ($Submit) {
            # If Submit parameter was set when running this function, run CreateUser function and pass all parameters to it
            Write-Host " Submitting now!"
            $formStatusBar.Text = "Submitting now!"
            CreateUser -name $name -fname $fname -lname $lname -samAccountName $5and2 -LineURI $LineURI -homeDirectory $homeDirectory -site $site -upn $upn -siteOU $siteOU -ExchDB $ExchDB -password $password -employeeId $employeeId
        }
    }
    else {
        # If one of the required fields does not have a value, notify user and do not process anything else
        $formCheckDetailsButton.Content = "Check Details"
        $formStatusBar.Text = "Form incomplete! Check all fields."
    }
    return $name, $fname, $lname, $5and2, $site, $LineURI, $homeDirectory, $upn, $siteOU, $password, $ExchDB
}

function Get-NextAvailableLineURI {
# This function pulls the next available VOIP number in Skype for Business based off of the provided site code.
    Param(
    [int]$SiteCode
    )
    $MinLineURI = $SiteCode * 1000 
    $MaxLineURI = $MinLineURI + 999
    # Utilize CleanOUList to get DN of the site based off of the site code
    $siteOU = ($CleanOUList | where {$_.SiteCode -EQ $SiteCode}).DistinguishedName
    #Write-Output $siteOU
    # Get all active LineURIs
    $ActiveLineURIs = Get-CsUser -Filter {LineURI -ne $null} | select LineURI,samAccountName
    #$ActiveLineURIs = Get-CsUser -OU $siteOU -Filter {LineURI -ne $null} | select LineURI,samAccountName
    
    #Find the total possible range of the site LineURIs
    $Range = ($MinLineURI+50)..$MaxLineURI

    # Define array which will contain only the data we need
    $NewActiveLineURIs = @()
    foreach ($line in  $ActiveLineURIs) {
        # Remove "tel:" from LineURI columns
        $line.LineURI = ($line.LineURI -replace 'tel:','')
        # Check to see if the LineURIs are within scope set by Min/MaxLineURI variables, if true then add to new array
        if (($line.LineURI -lt $MaxLineURI) -and ($line.LineURI -gt $MinLineURI) )  {
            #Write-Output "$($line.LineURI) is within scope"
            $NewActiveLineURIs = $NewActiveLineURIs += $line.LineURI.ToString()
        }
    }
    # Compare two arrays, find the available LineURIs, select the first available
    $AvailLineURIs = $Range | Where {$NewActiveLineURIs -notcontains $_}
    $NextAvailableLineURI = $AvailLineURIs | sort | select -First 1
    # Sort list so we get the last LineURI, select the last item, convert to an integer then add 1
    #$NextAvailableLineURI = (($NewActiveLineURIs |sort | select -last 1) -as [int]) +1

    if ($NextAvailableLineURI -eq "1") {
    # If no line URIs are found, create the first one
        Write-Warning "This is the first number for the given range!"
        $SiteCode *= 1000
        $SiteCode += 50
        $NextAvailableLineURI= $SiteCode
    }
    return $NextAvailableLineURI
}

function CreateUser {
# Pulls the parameters given through EvalAccountDetails, then actually creates and modifies the new account
    Param(
    [string]$name,
    [string]$fname,
    [string]$lname,
    [string]$samAccountName,
    [string]$site,
    [int]$LineURI,
    [string]$homeDirectory,
    [string]$upn,
    [string]$siteOU,
    [string]$ExchDB,
    [string]$password,
    [string]$employeeId
    )
    Write-Host "In CreateUser function"
    # Need to convert password to SecureString in order to use in New-Mailbox command
    $pwd =  $password | ConvertTo-SecureString -AsPlainText -Force
    $Description = "$($site) User"
    try {
        # Create new mailbox which also creates AD account, matches correct Exchange database based off of what 
        # EvalAccountDetails calculated. Archive database is always the same here, so it's called at the top of the script
        New-Mailbox -Name $name -UserPrincipalName $upn -Alias $samAccountName -OrganizationalUnit $siteOU -SamAccountName $samAccountName -FirstName $fname -LastName $lname -Password $pwd -Database $ExchDB -ArchiveDatabase $ExchArchiveDB
        $formStatusBar.Text ="Mailbox provisioned. Pausing for 10 seconds for Active Directory changes to replicate..."
        Write-Host "Mailbox provisioned. Pausing for 15 seconds for Active Directory changes to replicate..."
        Start-Sleep -Seconds 15
        # Now set all of the AD attributes that we care about
        Set-ADUser -Identity $samAccountName -HomeDirectory "$homeDirectory" -HomeDrive $homeDrive
        Set-ADUser -Identity $samAccountName -City "$site"
        Set-ADUser -Identity $samAccountName -Description "$Description"
        Set-ADUser -Identity $samAccountName -Office "$site"
        Set-ADUser -Identity $samAccountName -OfficePhone "$LineURI"
        Set-ADUser -Identity $samAccountName -ScriptPath "$scriptPath"
        Set-ADUser -Identity $samAccountName -ChangePasswordAtLogon $true
        Set-ADUser -Identity $samAccountName -EmployeeID $employeeId
        $formStatusBar.Text = "Active Directory account changes made. Pausing for 5 seconds..."
        Write-Host "Active Directory account changes made. Pausing for 5 seconds..."
        Start-Sleep -Seconds 5
        $formStatusBar.Text = "Skype-enabling user and assigning VOIP number..."
        # Now Skype-enable the account and assign to pool
        Write-Host "trying: Enable-CsUser -Identity $name -RegistrarPool $SkypePool -SipAddressType SamAccountName -SipDomain $domainName"
        Enable-CsUser -Identity $name -RegistrarPool $SkypePool -SipAddressType SamAccountName -SipDomain $domainName
        Write-Host "Skype for Business enabled for account. Pausing for 15 seconds..."
        $formStatusBar.Text = "Skype for Business enabled for account. Pausing for 15 seconds..."
        Start-Sleep -Seconds 15
        # Now assign LineURI (adding "tel:" to make it dialable) and enable Enterprise Voice
        Write-Host "Trying: Set-CsUser -Identity $name -Enabled $true -LineURI 'tel:$LineURI' -EnterpriseVoiceEnabled $true"
        Set-CsUser -Identity $name -Enabled $true -LineURI "tel:$LineURI" -EnterpriseVoiceEnabled $true
        Write-Host "Skype for Business VOIP number assigned and Enterprise Voice enabled. Pausing for 5 seconds..."
        $formStatusBar.Text = "Skype for Business VOIP number assigned and Enterprise Voice enabled. Pausing for 5 seconds..."
        Start-Sleep -Seconds 5
        # Enable Unified Messaging on the account, using the LineURI parameter for the extension
        Write-Host "trying: Enable-UMMailbox -Identity $name -UMMailboxPolicy $UMMailboxPolicy -SendWelcomeMail $true -Extensions $LineURI -PinExpired $true -SIPResourceIdentifier $upn -Pin $LineURI"
        Enable-UMMailbox -Identity $name -UMMailboxPolicy $UMMailboxPolicy -SendWelcomeMail $true -Extensions $LineURI -PinExpired $true -SIPResourceIdentifier $upn -Pin $LineURI
        Write-Host "Unified Messaging enabled for account."
        $formStatusBar.Text = "Unified Messaging enabled for account."
        Start-Sleep -Seconds 5
        # Grey out create user button so we don't accidentally create multiple accounts
        $formCreateUserButton.IsEnabled = $false
        $formCreateUserButton.Text = "Account Created"
        $formStatusBar.Text = "Account $($upn) has been created."
    }
    catch {
        $formStatusBar.Text = "Could not create account"
        Write-Warning "Ran into a problem. Error message:"
        Write-Warning $_.Exception.Message
        [System.Windows.MessageBox]::Show("Error creating account. Message: $($_.Exception.Message)","Error","OK","Error")
    }
}

function CreateNewUser {
# Resets everything so a new user can be created
    $formCreateUserButton.IsEnabled = $true
    $formFirstName.Text = ""
    $formLastName.Text = ""
    $formMyID.Text = ""
    $formSetNew5and2.Text = ""
    $FormDropDownSite.SelectedItem = $null
    $formPassword.Clear()
    $formResultData.Visibility = "Hidden"
}

#endregion Functions


#region Generate form

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

[xml]$Form = @"
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Create New User" Height="433" Width="918" ResizeMode="NoResize">
    <Grid Margin="0,0,0,0">
        <Label x:Name="FirstName_Label" Content="First Name:" HorizontalAlignment="Left" Margin="33,22,0,0" VerticalAlignment="Top" Height="26" Width="70"/>
        <TextBox x:Name="FirstName" HorizontalAlignment="Left" Height="23" Margin="110,25,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label x:Name="LastName_Label" Content="Last Name:" HorizontalAlignment="Left" Margin="34,61,0,0" VerticalAlignment="Top" Height="26" Width="69"/>
        <TextBox x:Name="LastName" HorizontalAlignment="Left" Height="23" Margin="110,64,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button x:Name="CheckUserButton" Content="Check Availability" HorizontalAlignment="Left" Margin="110,104,0,0" VerticalAlignment="Top" Width="120" Height="44"/>
        <ComboBox x:Name="DropDownSite" HorizontalAlignment="Left" Margin="110,207,0,0" VerticalAlignment="Top" Width="120" Height="22"/>
        <Label x:Name="DropDownSite_Label" Content="Site:" HorizontalAlignment="Left" Margin="71,203,0,0" VerticalAlignment="Top" Height="26" Width="32"/>
		<Label x:Name="Password_Label" Content="Password:" HorizontalAlignment="Left" Margin="41,244,0,0" VerticalAlignment="Top" Height="26" Width="62"/>
        <PasswordBox x:Name="Password"  HorizontalAlignment="Left" Height="23" Margin="110,248,0,0" VerticalAlignment="Top" Width="120"/>
        <Label x:Name="MyID_Text" Content="MyID:" HorizontalAlignment="Left" Margin="62,285,0,0" VerticalAlignment="Top" Height="26" Width="41"/>
        <TextBox x:Name="MyID" HorizontalAlignment="Left" Height="23" Margin="110,289,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label HorizontalAlignment="Left" Margin="12,129,0,0" VerticalAlignment="Top" Height="56" Width="95">
            <Label.Content>
                <AccessText x:Name="SetNew5and2_Label" TextWrapping="Wrap" Text="Account name is in use! Please enter new 5+2:" Width="85" Foreground="Red" Visibility="Hidden"/>
            </Label.Content>
        </Label>
        <TextBox x:Name="SetNew5and2" HorizontalAlignment="Left" Height="23" Margin="110,164,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
        <DataGrid x:Name="ResultData" HorizontalAlignment="Left" Height="287" Margin="285,25,0,0" VerticalAlignment="Top" Width="604" Background="{x:Null}" BorderBrush="{x:Null}" CanUserResizeColumns="False" CanUserReorderColumns="False" CanUserResizeRows="False" Visibility="Hidden"/>
        <Button x:Name="CheckDetailsButton" Content="Check Details" HorizontalAlignment="Left" Margin="110,328,0,0" VerticalAlignment="Top" Width="120" Height="44"/>
        <Button x:Name="CreateUserButton" Content="Create Account" HorizontalAlignment="Left" Margin="445,328,0,0" VerticalAlignment="Top" Width="120" Height="44"/>
        <Button x:Name="NewUserButton" Content="New Account" HorizontalAlignment="Left" Margin="769,328,0,0" VerticalAlignment="Top" Width="120" Height="44"/>
        <StatusBar DockPanel.Dock="Bottom" Margin="0,383,0,0">
            <StatusBarItem>
                <TextBlock Name="StatusBarText" />
            </StatusBarItem>
        </StatusBar>
	</Grid>
</Window>
"@

$XMLReader = (New-Object System.Xml.XmlNodeReader $Form)
$XMLForm = [Windows.Markup.XamlReader]::Load($XMLReader)

# Link variables with form elements
$formFirstName = $XMLForm.FindName("FirstName")
$formLastName = $XMLForm.FindName("LastName")
$formCheckUserButton = $XMLForm.FindName("CheckUserButton")
$FormDropDownSite = $XMLForm.FindName("DropDownSite")
$formPassword = $XMLForm.FindName("Password")
$formMyID = $XMLForm.FindName("MyID")
$formSetNew5and2 = $XMLForm.FindName("SetNew5and2")
$formSetNew5and2_Label = $XMLForm.FindName("SetNew5and2_Label")
$formResultData = $XMLForm.FindName("ResultData")
$formCheckDetailsButton = $XMLForm.FindName("CheckDetailsButton")
$formCreateUserButton = $XMLForm.FindName("CreateUserButton")
$formNewUserButton = $XMLForm.FindName("NewUserButton")
$formStatusBar = $XMLForm.FindName("StatusBarText")


# Define interactions with the form
$formCheckUserButton.Add_Click({CheckUser})
foreach ($item in $DropDownArray) {[void] $FormDropDownSite.Items.Add($item)}
$formCheckDetailsButton.Add_Click({EvalAccountDetails})
$formCreateUserButton.Add_Click({EvalAccountDetails -Submit $true})
$formNewUserButton.Add_Click({CreateNewUser})
$formStatusBar.Text = "Ready"

$XMLForm.ShowDialog()

#endregion Generate form


# Remove created PSSessions
Remove-PSSession $ExchangeSession
Remove-PSSession $SkypeSession
