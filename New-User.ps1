### Active Directory/Exchange/Skype for Business Account Creation GUI

# Make sure to fill out section below for your environment. Review the process everywhere in the script 
# that determines account naming convention (search for 5+2) and make sure it conforms to your naming convention.

# This script assumes that you are running from a machine that has RSAT installed with the Active Directory
# PowerShell module added, and that you are running the script with an account that has access to Active 
# Directory, Exchange and Skype for Business with account creation and modify rights.

# Set environment specific variables and server names
$ExchangeServer = "placeholder"
$SkypeServer = "placeholder"
$domainName = "placeholder"
$scriptPath = "placeholder"
$UserOUBase = "placeholder"
$SkypePool = "placeholder.$domainName"
$ExchArchiveDB = "Archive Database"
$UMMailboxPolicy = "placeholder"
$homeDrive = "placeholder"
$scriptPath = "placeholder"

#region Active Directory initial commands

# Load AD module
if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    Write-Host "Imported Active Directory module."
}
else {
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
        $formOutputBox.ForeColor = "Green"
        $formOutputBox.Text = "$($5and2) is available"
        $formStatusBar.Text = "$($5and2) is available"
        if (!$alt5and2){
            $formSetNew5and2.Visible = $false
        }
    }
    else {
        # If in use, make new 5+2 field visible on the form and notify user that account is in use
        $formOutputBox.ForeColor = "Red"
        $formOutputBox.Text = "$($5and2) exists already. `n Enter alternative 5+2:"
        $formSetNew5and2.Visible = $true
        $formSetNew5and2.Text = $5and2
        $formStatusBar.Text = "$($5and2) exists already. Enter alternative 5+2:"
    }
    if ((!$5and2) -and (!$alt5and2)) {
        $formOutputBox.ForeColor = "Red"
        $formOutputBox.Text = "Invalid account name"
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
    $password = $formPassword.Text
    $site = $FormDropDownSite.Text
    if (($lname) -and ($fname) -and ($password) -and ($site)) {
        # Checks to make sure that all required fields have values
        $formCheckDetailsButton.Text = "Check Details"
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
        # Output all result data to the form. Label object is always the same, data object is created in this function
        $formResultData_Label.Text = " DisplayName :`n First Name :`n Last Name :`n Account :`n Site :`n VOIP Number :`n Password :`n UPN :`n Home Directory :`n SiteOU :`n ExchangeDB :"
        $formResultData.Text = "
            $name
            $fname
            $lname
            $5and2 
            $site
            $LineURI 
            $password
            $upn
            $homeDirectory
            $siteOU
            $ExchDB"
        # Make the objects visible now
        $formResultData_Label.Visible = $true
        $formResultData.Visible = $true
        $formStatusBar.Text = "Generated account details"
        if ($Submit) {
            # If Submit parameter was set when running this function, run CreateUser function and pass all parameters to it
            Write-Host " Submitting now!"
            $formStatusBar.Text = "Submitting now!"
            CreateUser -name $name -fname $fname -lname $lname -samAccountName $5and2 -LineURI $LineURI -homeDirectory $homeDirectory -site $site -upn $upn -siteOU $siteOU -ExchDB $ExchDB -password $password
        }
    }
    else {
        # If one of the required fields does not have a value, notify user and do not process anything else
        $formCheckDetailsButton.Text = "Form incomplete! Check Details"
        $formStatusBar.Text = "Form incomplete!"
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
    # Get all active LineURIs for the specified site OU
    $ActiveLineURIs = Get-CsUser -Filter {LineURI -ne $null} | select -ExpandProperty LineURI 
    
    # Define array which will contain only the data we need
    $NewActiveLineURIs = @()
    foreach ($line in $ActiveLineURIs) {
        # Remove "tel:" from LineURI columns"
        $line.LineURI = ($line.LineURI -replace 'tel:','')
        # Check to see if the LineURIs are within scope set by MinLineURI and MaxLineURI variables, if true then add to new array
        if ( ($line.LineURI -lt $MaxLineURI) -and ($line.LineURI -gt $MinLineURI) ) {
            $NewActiveLineURIs = $NewActiveLineURIs += $line.LineURI.ToString()
        }
    }
    # Sort list so we get the last LineURI, select the last item, convert to an integer then add 1
    $NextAvailableLineURI = (($NewActiveLineURIs |sort | select -last 1) -as [int]) +1
    
    # If no LineURIs are found, create the first one
    if ($NextAvailableLineURI -eq "1") {
        Write-Warning "This is the first LineURI for the given range"
        $SiteCode *= 1000
        $SiteCode += 50
        $NextAvailableLineURI = $SiteCode
    }
    return $NextAvailableLineURI
}

function CreateUser {
# Pulls the parameters given through EvalAccountDetails, then creates and modifies the new account
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
    [string]$password
    )
    # Need to convert password to SecureString in order to use in New-Mailbox command
    $pwd =  $password | ConvertTo-SecureString -AsPlainText -Force
    $Description = "$($site) User"
    try {
        # Create new mailbox which also creates AD account, matches correct Exchange database based off of what 
        # EvalAccountDetails calculated. Archive database is always the same here, so it's called at the top of the script
        New-Mailbox -Name $name -UserPrincipalName $upn -Alias $samAccountName -OrganizationalUnit $siteOU -SamAccountName $samAccountName -FirstName $fname -LastName $lname -Password $pwd -Database $ExchDB -ArchiveDatabase $ExchArchiveDB
        $formStatusBar.Text ="Mailbox provisioned. Pausing for 10 seconds for Active Directory changes to replicate..."
        Write-Host "Mailbox provisioned. Pausing for 10 seconds for Active Directory changes to replicate..."
        Start-Sleep -Seconds 10
        # Now set all of the AD attributes that we care about
        Set-ADUser -Identity "$samAccountName" -HomeDirectory "$homeDirectory"
        Set-ADUser -Identity "$samAccountName"-City "$site" -Verbose
        Set-ADUser -Identity "$samAccountName" -Description "$Description"
        Set-ADUser -Identity "$samAccountName" -Office "$site"  -Verbose
        Set-ADUser -Identity "$samAccountName" -OfficePhone "$LineURI"
        Set-ADUser -Identity "$samAccountName" -ScriptPath "$scriptPath"
        Set-ADUser -Identity $samAccountName -ChangePasswordAtLogon $true
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
        $formCreateUserButton.Enabled = $false
        $formCreateUserButton.Text = "Account Created"
        $formStatusBar.Text = "Account $($upn) has been created."
    }
    catch {
        $formStatusBar.Text = "Could not create account, make sure account name is unique!"
        Write-Warning "Could not create account, make sure account name is unique!"
    }
}

function CreateNewUser {
# Resets everything so a new user can be created
    $formCreateUserButton.Enabled = $true
    $formCreateUserButton.Text = "Create Account"
    $formFirstName.Text = ""
    $formLastName.Text = ""
    $formOutputBox.Text = ""
    $formSetNew5and2.Text = ""
    $FormDropDownSite.SelectedItem = $null
    $formPassword.Text = ""
    $formResultData_Label.Visible = $false
    $formResultData.Visible = $false
}

#endregion Functions


#region Generate form
   
$form = New-Object System.Windows.Forms.Form
$form.Width = 760
$form.Height = 350
$form.Text = "User Account Creation"
$form.StartPosition = 'CenterScreen'
# Make form layout fixed, not resizable
$form.FormBorderStyle = "Fixed3D"

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic")
[System.Windows.Forms.Application]::EnableVisualStyles()

####### Main form elements #######

### Textboxes/labels

$formFirstName_Label = New-Object System.Windows.Forms.Label
$formFirstName_Label.Text = "First Name:"
$formFirstName_Label.Width = 80
$formFirstName_Label.Height = 20
$formFirstName_Label.Location = New-Object System.Drawing.Point(10,10)
$formFirstName_Label.TextAlign = "MiddleRight"

$formFirstName = New-Object System.Windows.Forms.TextBox
$formFirstName.Multiline = $false
$formFirstName.Width = 150
$formFirstName.Height = 20
$formFirstName.Location = New-Object System.Drawing.Point(100,10)

  
$formLastName_Label = New-Object System.Windows.Forms.Label
$formLastName_Label.Text = "Last Name:"
$formLastName_Label.Width = 80
$formLastName_Label.Height = 20
$formLastName_Label.Location = New-Object System.Drawing.Point(10,40)
$formLastName_Label.TextAlign = "MiddleRight"

$formLastName = New-Object System.Windows.Forms.TextBox
$formLastName.Multiline = $false
$formLastName.Width = 150
$formLastName.Height = 20
$formLastName.Location = New-Object System.Drawing.Point(100,40)


$formCheckUserButton = New-Object System.Windows.Forms.Button
$formCheckUserButton.Text = "Check if account name is in use"
$formCheckUserButton.Size = New-Object System.Drawing.Size(110,50)
$formCheckUserButton.Location = New-Object System.Drawing.Point(10,70)
$formCheckUserButton.Add_Click({CheckUser})

$formOutputBox = New-Object System.Windows.Forms.Label
$formOutputBox.Width = 150
$formOutputBox.Height = 30
$formOutputBox.Location = New-Object System.Drawing.Point(130,70)
$formOutputBox.ForeColor = "Red"

$formSetNew5and2 = New-Object System.Windows.Forms.TextBox
$formSetNew5and2.Multiline = $false
$formSetNew5and2.Width = 50
$formSetNew5and2.Height = 20
$formSetNew5and2.Location = New-Object System.Drawing.Point(160,100)
$formSetNew5and2.Visible = $false
$formSetNew5and2.MaxLength = 7

$FormDropDownSiteLabel = New-Object System.Windows.Forms.Label
$FormDropDownSiteLabel.Width = 80
$FormDropDownSiteLabel.Height = 20
$FormDropDownSiteLabel.Location = New-Object System.Drawing.Size(10,130)
$FormDropDownSiteLabel.Text = "Site:"
$FormDropDownSiteLabel.TextAlign = "MiddleRight"

$FormDropDownSite = New-Object System.Windows.Forms.ComboBox
$FormDropDownSite.Location = New-Object System.Drawing.Point (100,130)
$FormDropDownSite.Size = New-Object System.Drawing.Size(110,20)
$FormDropDownSite.DropDownStyle = "DropDownList"
foreach ($item in $DropDownArray) {[void] $FormDropDownSite.Items.Add($item)}

$formPassword_Label = New-Object System.Windows.Forms.Label
$formPassword_Label.Text = "Password:"
$formPassword_Label.Width = 80
$formPassword_Label.Height = 20
$formPassword_Label.Location = New-Object System.Drawing.Point(10,160)
$formPassword_Label.TextAlign = "MiddleRight"

$formPassword = New-Object System.Windows.Forms.MaskedTextBox
$formPassword.PasswordChar= '*'
$formPassword.Width = 150
$formPassword.Height = 20
$formPassword.Location = New-Object System.Drawing.Point(100,160)

$formResultData_Label = New-Object System.Windows.Forms.Label
$formResultData_Label.Width = 110
$formResultData_Label.Height = 200
$formResultData_Label.Location = New-Object System.Drawing.Point(250,17)
$formResultData_Label.TextAlign = "TopRight"
$formResultData_Label.Visible = $false

$formResultData = New-Object System.Windows.Forms.Label
$formResultData.Width = 600
$formResultData.Height = 200
$formResultData.Location = New-Object System.Drawing.Point(330,5)
$formResultData.TextAlign = "TopLeft"
$formResultData.Visible = $false

### Buttons

$formCheckDetailsButton = New-Object System.Windows.Forms.Button
$formCheckDetailsButton.Text = "Check Details"
$formCheckDetailsButton.Size = New-Object System.Drawing.Size(130,40)
$formCheckDetailsButton.Location = New-Object System.Drawing.Point(10,240)
$formCheckDetailsButton.Add_Click({EvalAccountDetails})


$formCreateUserButton = New-Object System.Windows.Forms.Button
$formCreateUserButton.Text = "Create Account"
$formCreateUserButton.Size = New-Object System.Drawing.Size(130,40)
$formCreateUserButton.Location = New-Object System.Drawing.Point(300,240)
$formCreateUserButton.Add_Click({EvalAccountDetails -Submit $true})
$formCreateUserButton.Font = $buttonFonts

$formNewUserButton = New-Object System.Windows.Forms.Button
$formNewUserButton.Text = "Create Another Account"
$formNewUserButton.Size = New-Object System.Drawing.Size(130,40)
$formNewUserButton.Location = New-Object System.Drawing.Point(600,240)
$formNewUserButton.Add_Click({CreateNewUser})

### Status Bar

$formStatusBar = New-Object System.Windows.Forms.StatusBar
$formStatusBar.Width = 760
$formStatusBar.Height = 20
$formStatusBar.Location = New-Object System.Drawing.Point(5,290)
$formStatusBar.Text = "Ready"

### Add Controls

$form.Controls.AddRange((
$FormDropDownSite,
$FormDropDownSiteLabel,
$formFirstName,
$formFirstName_Label,
$formLastName,
$formLastName_Label,
$formCheckDetailsButton,
$formCheckUserButton,
$formOutputBox,
$formSetNew5and2,
$formPassword_Label,
$formPassword,
$formResultData_Label,
$formResultData,
$formCreateUserButton,
$formNewUserButton,
$formStatusBar
))

### Form icon

$form.Icon = [System.Convert]::FromBase64String('
AAABAAkAICAQAAEABADoAgAAlgAAABgYEAABAAQA6AEAAH4DAAAQEBAAAQAEACgBAABmBQAAICAA
AAEACACoCAAAjgYAABgYAAABAAgAyAYAADYPAAAQEAAAAQAIAGgFAAD+FQAAICAAAAEAIACoEAAA
ZhsAABgYAAABACAAiAkAAA4sAAAQEAAAAQAgAGgEAACWNQAAKAAAACAAAABAAAAAAQAEAAAAAAAA
AgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAIAAAACAgACAAAAAgACAAICAAACAgIAAwMDAAAAA
/wAA/wAAAP//AP8AAAD/AP8A//8AAP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAiIAAAAAAAAAAAAAAAACIiIiHMAAAAAAAAAAAAAiI//iIiHcAAAAAAAAAAIiP//+IiIiH
dwAAAAAAiI//////iIiIiHhwAAAAAIj///////iIiIiHeHAAAACIj////4iIiIiIiIiHgAAAiIj/
iIiIiIj/iIiHeIcAAIiIiIiIiIiIiP+IiIeIcACIiHd3d3d3eIeI+IiHeIhwczMRMzd3czeIeI/4
iIiIcAczN3iIiIczeId4j4iHiHAACIiP///4i7OIh4iPiHiAAAAIiI//iIu7M4h3iP+HcAAAAAiI
iIi7u7u3iHiIiHAAAAAACIeP///4i7iHeIgAAAAAAAAAiI+IiLu7t3AAAAAAAAAAAACDMzMzMzM3
AAAAAAAAAAAAAAAAAAAACIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//////////////////////////////////////////
///////8f///wB///gAP//AAA/8AAAH/AAAAfwAAAB8AAAAPAAAABwAAAAEAAAABgAAAAeAAAAH4
AAAB/gAAAf+AAAP/8AAf//wAD////+f/////////////////////KAAAABgAAAAwAAAAAQAEAAAA
AAAgAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAIAAAACAgACAAAAAgACAAICAAACAgIAAwMDA
AAAA/wAA/wAAAP//AP8AAAD/AP8A//8AAP///wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIiIAAAAAAAA
AAAIiIiIcAAAAAAAAIiI/4iIhwAAAAAIiP//+IiIh3gAAACI/////4iHiHeAAACIj//4iIiIeIiI
AACIiIiIiIiPiIh4iACIiId4iIiI+IiHiHCHczMzd3eIiPiHeIeDAzd4h3M3iI+Id4gAiIj//4iz
eIiPh4gAAIiIiIu7M3iIiHcAAACIiIiIizeIiIcAAAAAiI//iLt3gAAAAAAAAIhzMzMzMAAAAAAA
AAAAAAAAdwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8A////AP///wD///8A////AP///wD/
w/8A/gH/APAA/wCAAD8AAAAfAAAADwAAAAMAAAABAAAAAAAAAAAAwAAAAPAAAAD8AAAA/wAHAP/A
BwD///MA////AP///wAoAAAAEAAAACAAAAABAAQAAAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAgAAAgAAAAICAAIAAAACAAIAAgIAAAICAgADAwMAAAAD/AAD/AAAA//8A/wAAAP8A/wD//wAA
////AAAzMzMzMzMAAHd3d3MzMwAAd3d3czMzAACPiIiIiIMAAI+IiIiIgwAAj4iIiIiDAACPiIiI
iIMAAI+IiIiIgwAAj4iIh3eDAACPj///94MAAI+P///3gwAAj4////eDAACPiIiId4MAAHd3d3Mz
MwAAd3d3d3d3AAAAAAAAAAAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMA
AMADAADAAwAAwAMAAMADAADAAwAA//8AACgAAAAgAAAAQAAAAAEACAAAAAAAAAQAAAAAAAAAAAAA
AAEAAAABAAAAAAAAEj1BABlFSgAZSlIAIUpNAC5dXgAwV1kAMWRtAD9jaQA8aXoAVWttAFhpaQBK
bnQATH5/AF16ewAAaZMACmuTAAl0mgA/eIIAHXqgACp+owBNd4EAS32LAGF9gAAShJgALouZADyG
mQAKh6kAE4ujACWNqAA/jaMAMourAC6SrgAumbIAOJayAEmDjABEgpMAS4eQAFCHkABSiZIAVZOd
AFmUngBhh4kAaIeMAGmMjQBnlZsAaZmbAHWUkwB7lJMAcJyfAHWdnwBQkKQAX5ujAEKfugBqnaEA
UKK7AF2puwBpoaUAYqSqAGmlqQBvqq4AfaOsAHWprQB4q68Aaqm4AGqtvQB0rrIAea6xAHKxtgB5
s7cAdbS4AHy1uQBzu74Adbm+AH64vAAXo8kAAKbUABSy1QAdtdkAHrvbACGjyQA4uM4ALL3eABi7
4ABdssEATLbQAGOqwABmqsAAZq/AAGGtxQBqrcIAbrDBAHW8wQB5u8EAf73CAHy0yAB4v8oAZLfQ
ADrB3wAp0t8APNDaAD7N5wA23OQAQsPfAEvO2AB5wcYAf8DEAHzCygBuw9YAdsTUAGPQ2QB92NwA
RsTgAE3J4QBH3OsAUtDqAF7b7QBh1ecAWPz/AH3k8QB7//8AgKmqAImtsQCCsrQAhLa9AI67uwCR
ubwAmMC+AIe+xgCLvMQAgrjKAJy/xgCQv8oAgL7SAIi90ACCwcYAi8PHAITFzACNxMgAi8PNAIjH
zACEys8AiMnOAJTAwQCXxsgAm8bIAJvIzACGzdIAhs3UAI7J0ACIzNAAjs3SAJjD0QCTyNQAh9DU
AIvQ0wCI0tYAjdHVAIvS2ACP1NoAjNnfAJHS1gCd09QAk9XZAJvV2gCT2NoAlNjaAJHY3ACU2NwA
os3ZALLO0gCk09YArNHTAK7W2QCs298AstPWALHV2wCO2uAAk9vgAJXb4ACV3OEAl9/kAJzd5ACq
0eAAq9XgAKLZ4wCg3+YAqN7hAKjY5ACu2uUAqdznALDT4QC11eMAs9nlAJng5ACd4ucAneTnAJ3h
6gCe5uoAn+fsAI/s8wCd8vsAm///AKrg5ACi5+wAquToAKzl6wCl6e4AtOPpALrk6wC26+4At+zu
ALzq7QCk6/AAquvwAKjt8QCx7fEAuO3wAKrw8gCq8PYArPL3AK319gCm8fgArvP4ALXw8wC28/YA
vfHzALvy9gC09foAs/r+ALz//wDI19wAw9zoAMHq7QDF6e8AwuzuAN7u7wDD7fAAxu3wAMju8ADC
8fMAx/HzAMjx8gDJ8vUAzPP2AM329gDP+PcAxff6AMP8/wDK/PwA2PD0ANH//wDY//8A6P//APL/
/wD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC3uW0AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA6MC3ztfa
pS0JAAAAAAAAAAAAAAAAAAAAAAAAAADAwdHu8u/Z19eTjHlBAAAAAAAAAAAAAAAAAAAAAMC96fL2
8vLu6dLX15OaoC8rWQAAAAAAAAAAAADAvb3u8vLy8vLy8u7p2Nfak5OcXjJ9WwAAAAAAAAAAAMXa
4fDy8vLy8vLy8uLXw7NIaZOeoBcwlVkAAAAAAAAAw9fX4fDw8vLy8tTLpZyew8WTXImeh0V7j0GG
AAAAAADD2tfa5OLS0L2koaXD19rf5eXFa1ycsQ4sooFZAAAAALna18WzpZyJaYeJiYeHh4eVz+Xk
pVxrlUdDjI9BAAAAw7OTazolIyYnJygpNDtCSkVCw9/l14lpjQs+nH98WgBBEgcDAQEEBggMFRYk
Gh44rZA5ic/l5bNpRkpHjo8/AAAnAwIFDS57kaytq4lUIRs1qa8+PKXa5dqcSApHnkQAAAAAgWCN
zvn7+/n45Ml3c05LYbB+NofM3+XFXkSHSAAAAAAAANBrgK739+TeyHd0ZVNMUIWqMjyl1+Xlnio7
AAAAAAAAAADQjHyk13d1cXBnYlJPTVWYejaHw9fapToAAAAAAAAAAAAA6YlZmfr+/v385sp4dnJs
g3lCiJKsAAAAAAAAAAAAAAAAAACJhOfs1MNvbmhkZmNRNz0AAAAAAAAAAAAAAAAAAAAAAAAAYBkY
HB0gIh8UExEPEDMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAX4IAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAP/////////////////////////////////////////////////8f///wB///gAP//AAA/8A
AAH/AAAAfwAAAB8AAAAPAAAABwAAAAEAAAABgAAAAeAAAAH4AAAB/gAAAf+AAAP/8AAf//wAD///
/+f/////////////////////KAAAABgAAAAwAAAAAQAIAAAAAABAAgAAAAAAAAAAAAAAAQAAAAEA
AAAAAAAKNjwAH0tOACFQVgAsWV8AIFdhAClcZAA4ZGwAQGxuAEBtdQARaIMAMnePAEJ0gQBFf4YA
QnuJAAuBnwAcgp0AH5CuADOMpAA5i6EAKJSvAC2arAAsmbEAIZu7ADuuvgBEgo8AXoqMAFCQmwBV
k5wAX5CfAF+WnABjhIUAYYeKAG2XmQBmmJsAeZaVAHaZmgBIl6cARp+rAEafrABYkKEAWZWlAGOZ
pwBgnqoAaKGnAGyqrwB/oKEAeqKlAHSnqwB7o6wAe6WvAHCorQB3q68Abq61AHajswBxqrIAdKyy
AHevtQByrrgAd6+8AG+yuQB9srYAcLO5AHO0ugB5tLkAeLu/AH64vQACmsUAErjdACi52gBur8AA
d7fDAHe9wgB3vsQAe7rAAHm9wwB8uMQAfL/GAHq5zAB6vtEAM8XOADzM5gA80uYAfsHFAH7DyAB8
wcwAf8TOAF/a6wBj2u4AaNnqAHra7QBV9PUAfebwAIyvsQCArr0AgLK2AIqxsgCJsrcAkLW6AJO4
vwCMvsAAlr7CAI7CwgCBxcoAh8fJAIPHzACKx8wAg8rOAIvKzwCMys0AmMLFAJ7GxQCRxMkAnsLM
AJrFzACSyc0Als3OAJ3JyACCwNUAgsrRAITL0ACHzNcAi83SAI3N0QCLztQAjM7bAJHP0QCZyNIA
n8nRAJvJ1ACO0NUAj9PfAIzX3QCe0NAAkdXaAJTX3QCd1NgAn9bfAJTZ2wCR2NwAlNndAKLI0gCj
ytUApcrUAKzJ0wCwy9YApNDXAKbW2wCq1dwAu9beAIXV4gCB2esAhNrrAIrc6gCa1uAAl9vgAJbc
4QCZ3+MAm9vlAJjf5ACk3OAAqt3jALXd4wCb4+cAnOLlAJ7k6QCY7vcAgfPzAKfh5QCg4ukAp+fr
AKbm7ACq4eoApOrvAKft7wCq7O8AtubqALDj7QC44usAueXoALbq6wC37O8Av+nrALvp7wCm6/EA
pOzwAKju8gCq7/QAs+zwALnv9ACu8PMAq/H0AKzx9QCt9/UAo/D4AK7z+QCv9PgAsvD0ALj3/ACy
+f0Atfr9ALX8/wDK4OYAxOvuAMXt8QDB8PIAwvP0AMT09QDJ8PEAzPLzAMny9QDM8/cAyPb2AM32
9gDO+PgAyv3/ANP9+gDT//8A7vHyAPf7+wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///wAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlYF5XgAAAAAAAAAAAAAAAAAA
AAAAkXGBoK6kQh0AAAAAAAAAAAAAAAAAkI2UttHRtbuLQmYqAAAAAAAAAAAAjY2izNHW0dHLvMCL
QGsjOXYAAAAAAACewM/R0dHR0dHPqptCM1SEJC5PAAAAAAClwMXP0dHMs5NtaXybaTx5aT1gTAAA
AACqwLuvqIhzanybrrvDyZ4/SHkfZHBOAACqu5tMPzU1P0xTaX6LpcnDUz9rIV91RgBHHA0GAwQH
CQwOGRxfc4vAyYtIIjNtbzqSBQECCBouXWA5JRATMW98pMm4aS0ggkIAAHFUh7XR2dbGpllFFyli
cIe7yYssQEwAAAAAtpqhsL3CXFdRREMSMWVtnru4MB0AAAAAAACziISWmZiYWlhSGCtjhXx8bTkA
AAAAAAAAALB9ytrb2cCnW1AmOoEAAAAAAAAAAAAAAAAArFQmFRYWFBEPCgsAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAADYpAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAD///8A////AP///wD///8A////AP///wD/w/8A/gH/APAA/wCAAD8AAAAfAAAADwAAAAMA
AAABAAAAAAAAAAAAwAAAAPAAAAD8AAAA/wAHAP/ABwD///MA////AP///wAoAAAAEAAAACAAAAAB
AAgAAAAAAAABAAAAAAAAAAAAAAABAAAAAQAAAAAAAABfggAljaMALZClADaTpgA/l6gAQouhAEma
qgBTnqsAXKKtAGalrwBvqbEAeKyyADOixQBBrMsAQ63MAEWvzQBGsM4ASbLPAFC10QBZudMAW7zV
AGe/1wBowtkAdsXbAHfI3QBo0OYAcNPoAHnW6QCGzN8Ahc7kAIjP4QCC2usAjN3tAJnT4wCb1uUA
n9XlAKTY5wCq2+gAr93qAJbh7gCf5fAAqejyALLs9AC77/UAvP//AMX//wDP//8A2f//AOL//wDs
//8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAA////AAAAAQEBAQEBAQEBAQEBAAAAAAYMCwoJCAcFBAMCAQAA
AAAGDAsKCQgHBQQDAgEAAAAAHiwrKikoISAcGxoNAAAAAB4sKyopKCEgHBsaDQAAAAAeLCsqKSgh
IBwbGg0AAAAAHiwrKikoISAcGxoNAAAAAB4sKyopKCEgHBsaDQAAAAAeLCIdGBYUExEOGg0AAAAA
HiwiMjEwLy4tDhoNAAAAAB4sJTIxMC8uLQ4aDQAAAAAeLCYyMTAvLi0RGg0AAAAAHiwmIh8ZFxQT
EhoNAAAAAAYMCwoJCAcFBAMCAQAAAAAGBgYGBgYGBgYGBgYAAAAAAAAAAAAAAAAAAAAAAADAAwAA
wAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAADAAwAAwAMAAMADAAD/
/wAAKAAAACAAAABAAAAAAQAgAAAAAACAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAdKbGAYG0zTSOwtRUhLzS
r5zX4e1qv9HqCF59egAnQg4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgbHMQI2+1IaY
xtiyqNTf9Kzl6/+p7vH/qvHy/5PY2v9nlZv/ADtSww9tjTMAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAaJ/CA4y91CR/tM5JjL3U
pafU4dy45Ov3xu3x/8rx8v/I7vD/uO3w/6nt8f+o7PH/hs7T/4jHzP+Aqqr/YKq67wd5oHsAR2YK
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAbaLBPoa4z26R
wtasqtjk8cDq7v/K8/X/z/j3/8z29v/J8vP/xu3w/8Pr7f+26+7/qOzw/6jt8f+HztL/h9DU/43a
3/91lJP/aIeM/yqRsb0AX4s3AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAf7jO
nJLP3sul2+bzw+3w/8329v/O9vf/zPP2/8ry9f/I8vX/yfL1/8jx8v/H7vH/wuzu/7Ht8f+p7vL/
qvDz/4bN0P+HztP/idLW/3+9wv91nZ//jru7/2Orve0Dc5t7AE9zCQAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAACd4er/rvX2/7bz9v/D8vP/yfPz/8jy9f/I8fP/yfP1/8ny9f/J8vX/yfHz/8jw
8v+98fP/quvw/53i5/+V2+D/c7u+/3nBxv+HztP/i9PY/4zZ3/9hfYD/e5ST/47J0P81kK3BAGiQ
NgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJvf5/mp7vL/qu7z/7Xw8//B8fP/x/Hz/8rz9v/M8/f/
yvP2/8jy8/++6+7/quDk/5TY2v+M09b/j9Xb/5ng5f+f5+z/h87U/3S7wf+Ax87/i9LY/4LCx/95
srf/gbK1/5XBwv9ep7rqFH6kgQAAAAMAAAAAAAAAAAAAAAAAAAAAm97m+anw8v+p7vP/rPL3/7X2
+P+78vb/t+zu/7Tm6v+o3uH/mdXZ/5HS1v+T2Nv/neTn/6bt8v+q8Pb/r/P4/7L4/v+1/P//n+br
/3zDyP93vsP/iNLW/47a4P9denv/aYyN/53T1P+Mu8H/Mo+swQBqkzcAAAAAAAAAAAAAAACe3+b5
q/Dz/6jt8v+f5ur/l9/k/5LY3f+L0NP/hMbK/3/AxP+Dwsb/h8bL/4bHyv+ExMf/hMLG/4PAxf+F
wsb/js3S/6bp7v+0+///sPf8/5HY3P90u8H/fcbM/4jM0P98tbj/ea6x/43EyP+XwMD/Y6i58QB1
n38AVnUJAAAAAJ7e5fqT2+D/hs3U/3rCyf9hoqr/S4eQ/0mDjP9Qh5D/UIiS/1OMlf9Vk53/WZSe
/1+bo/9ppKj/dK6y/365vP95tLf/dK+0/5rc4f+v8/j/t/7//6Xs8P+Axsz/ecHH/4PMz/9YaWn/
dqqu/47Q1f+YwL7/h7e9/yuLqbMAAAAKaqm4/z94gv8xZG3/G0tT/xA7Qf8VP0L/IUpN/zBXWf8/
Y2n/Sm50/013gf9LfYv/RIKT/zyGmf8/jaP/Xam7/67W2/+Xxsj/aaGl/4XDyP+k6u7/sfj9/7L5
/v+V3OH/ecLH/3W0uP9/uLz/fLa7/4jJzv+SwML/d6uv/QAAACqCy9NpPHiA3hhKUf8ZRUr/Ll1e
/0x+f/9pmZv/g7K0/5vGyP+r0NL/rtbY/6TT1v+Hxs3/XbLB/y6Zsv8Kh6n/Qp+6/6LN2f+y09b/
dais/2+prv+U2Nz/qvD1/7X7//+r8vf/iNLW/3S6vv9Va23/fLW4/4/V2/9ur7X2AAAAKAAAAACm
7fMEh8zVZGeqtsJ5v8n/hsjP/6rk6P/J+vr/0P///9L////M/v//w/z//7T3/v+d8vv/f+bz/1LQ
6v8dtdn/F6PJ/2S30P+x1dv/kbm8/2icof+AwMT/oufs/671+P+1+///nufr/3m7wf9xsbb/hcLH
/3O3vPZNTU0qAAAAAAAAAAAAAAAAktbfAqLq9UyQ1d6tf8DL/4e+xv+s29//x/f5/8P3+/+38/n/
pvH4/4/s8/955fD/Xtvt/z7N5/8Yu+D/AKbU/yGjyf+AvtL/ss7S/3Ccn/9vq67/k9fa/6br8P+0
+f7/svz//4/U2P9hh4n/ZqSp+MHBwUEAAAAAAAAAAAAAAAAAAAAAAAAAAKvt8wKw8PUuktPeqIK+
yeuCtr3/ntXc/6Do8f9/4vD/YdXn/03J4f9GxOD/QsPf/zrB3/8svd7/Hrvb/xSy1f9MttD/mMPR
/4mtsf9snqH/gMLG/5rg5P+m7PH/rPX3/5PU2P9kpqv/x8fHHAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAn+LtKI7U4YF4v8voZq/A/5PI1P/Y8PT/8v////P////o////2P///7z///+b
////e////1j8//9H3Ov/bsPW/5y/xv9/p6r8b6qv83O3vNJorLKoY6mvhF+lq3MAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAJLi7B6C1uFpc7rI3pC/yv/I19z/3u7v
/7vq7P+Z4OT/fdjc/2PQ2f9Lztj/PNDa/zbc5P8p0t//OLjO/0yguvg6c4GnRHB1aDxiZwsAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAWsbXZUuqur8ui5n/EoSY/wmGn/QDe5vcAHud0QB5nccAb5fNAGWR1QBplOIAb5f2AGmT/wpr
k/8pd5DQMm5+eAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAVdLfA0CxxkQmjq5FKIuwNSiGriIukrgaM5vBEiqTuhYig7Ad
I4ayJwxxoDcOdKRBIIWvWBp8n5Ecf6CMHoGlKgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/////////////////////////////
////////////////////4B///4AP//AAA//AAAH/AAAAfwAAAD8AAAAPAAAABwAAAAEAAAAAAAAA
AAAAAACAAAAA4AAAAPgAAAD/AAAB/8AAB//4AAf//AAD/////////////////////ygAAAAYAAAA
MAAAAAEAIAAAAAAAYAkAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABCcHgCQnB4BQAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAgbLNAYCyzRFzpLpLfLC/g4e+zNOBytXyIXGMkQBFbBMAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAGuesySBr7xSe6m7mIe0wc6d
yND4pNzg/6ft7/+c4uX/fri9/ydpfbwUmMMxAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAQnB4CF+QozJmma1of628pY68ydCq1dz+v+nr/8zy8//K8PH/t+zv/6ru8/+R2N3/fri9
/47Cwv9Rjp3kE4OvYgB2rgEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABso7R7erHAq5PBztey3OL0
xOzw/8/29//O+Pj/zPX1/8jw8v/E6+7/s+zw/6zy9v+U2t3/erS5/4PKzf95lpX/da60+jeavJ8A
dqoaAAAAAAAAAAAAAAAAAAAAAAAAAACZ2+X4rvDz/8T09f/O9/b/y/P2/8ry9f/J8vX/yvL1/8rw
8v/B8PL/p+fr/5fb4P94u7//cKit/37DyP+M193/dpma/3+gof9csMfPB4GwRAAAAAAAAAAAAAAA
AAAAAACf5Or+q/H0/7Lw9P/C8/T/yPP1/8zz9//G7/L/ueXo/6bW2/+Mys7/g8fM/4vO1P+W3OH/
gcXL/2+yuf+CytH/g8TJ/32ytv+KsbL/c7PA7iqUuH4gj7QJAAAAAAAAAACf4un6rfL1/6nv9P+q
7O//p+Hl/53U2P+Syc3/isfM/4zN0f+X3eD/pOrv/6nt8v+u8/n/tfz//5nf4/9ztLr/d73C/4TL
0P9jhIX/jL7A/5LEyf88mbWtAYCxKgAAAACl5uz6qO/z/5bd4/98v8b/cLO5/2+vtv9trrX/c7S7
/3m9w/9+wMT/h8fJ/5HP0f+U2dv/neTp/7X6/f+v9Pj/f8LH/3Czuv+Dy8//bZeZ/4Gztv+dycj/
WaS33hOIslF3t8P/VZOc/0V/hv8pXGT/IVBW/yxZX/84ZGz/QG11/0J0gf9Ce4n/RIKP/1CQm/+A
sbb/ls3O/5HY3P+s8fX/s/r+/5TZ3v93vsT/Zpib/3err/+Lys//nsbF/2qptPButL+fH1Zg/Qo2
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
AADAAwAAwAMAAMADAADAAwAAwAMAAP//AAA=')

# Select field when form is loaded
$form.Add_Shown({$formFirstName.Select()})

# Show form
$form.ShowDialog()

#endregion

# Remove created PSSessions
Remove-PSSession $ExchangeSession
Remove-PSSession $SkypeSession
