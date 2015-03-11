#TODO Check we have connections?

Function LogThis($LogText, $Colour = "white")
{
    Write-Host $LogText -Fore $Colour
    $LogText | Out-File $LogFile -Append -NoClobber

    # Switch black / white color for html
    If ($Colour -eq "white") {$Colour = "black"}
    $Global:EmailBody += "<font face='arial' color= '$Colour'>$LogText</font></br>"
}

Function SplitOU($Line)
{
    $Temp = $Line.Split(",")
    $SplitReturn = $Temp[0]
    $SplitReturn = $SplitReturn.SubString(3)
    Return $SplitReturn
}

cls

$EmailBody = "<html><body><p>"

# Get info on who is running this and when
$User = [Environment]::Username
$Domain = [Environment]::UserDomainName
$Machine = [Environment]::MachineName
$DateRun = Get-Date -DisplayHint Date
$LogFile = "c:\scripts\logs\"
$RobocopyLog = $LogFile
$LogFile += Get-Date -Format "yyyy-MM-dd_hh-mm-ss"
$LogFile += "-$User.log"

$SMTPServer = "mail-01"
$EmailFrom = "noreply@company.com"
$EmailTo = "someone@company.com"
$EmailSubject = "Offboarding report"

$DirUserProfile = "\\file_server\profiles\"
$DirUserDesktop = "\\file_server\Desktops\"
$DirUserDocuments = "\\file_server\Documents\"
$DirArchive = "\\archive-server\Archive\"
$RobocopyLog += Get-Date -Format "yyyy-MM-dd_hh-mm-ss"
$RobocopyLog += "-$User-Robocopy.log"

$EmailAttachment = $RobocopyLog

#$RobocopyOptions = @("/R:3", "/W:30", "/S", "/MOVE", "/NP", "/Z", "/TEE", "/LOG+:$RobocopyLog")
$RobocopyOptions = @("/R:3", "/W:30", "/S", "/MOVE", "/NP", "/Z", "/LOG+:$RobocopyLog")

$RobocopyIgnoreDirs = @("/xd", '"My Music"', '"My Pictures"', '"My Videos"', "iTunes", "Recycler", "`$Recycle.bin")
$RobocopyIgnoreFiles = @("/xf", "~*", "*.tmp", "*.iso", "*.nrg", "*.img", "*.mp3", "*.mp4", "*.aac", "*.wav", "*.mpg", "*.mpeg", "*.avi", "*.mov", "$*", "*.swf", "*.vhd", "*.vmdk")

$RobocopyOptions += $RobocopyIgnoreDirs
$RobocopyOptions += $RobocopyIgnoreFiles

# Create.  Assumes it won't exist
# Which unless one person is running the same code twice in the same second, it won't
LogThis "Run Date: $DateRun"
LogThis "Run by:   $Domain\$User"
LogThis "Run on:   $Machine"

# Generate the new password (long and complicated - not meant to be used)
$Password = ""

$Alphabet = $NULL
For ($a=33; $a -le 126; $a++)
{
    $Alphabet += ,[char][byte]$a
}

For ($Loop = 1; $Loop -le 32; $Loop++)
{
    $Password += $Alphabet | Get-Random
}

# Ask for a login name and try to pull information, loop if we error when trying
Do
{
    $LoginName = Read-Host "Please enter the login name of the account to disable"
    $ADUserInfo = Get-ADUser $LoginName -Properties Manager, Title, CanonicalName, MemberOf, Company
} While (!$?)

# They should have a manager
If ($ADUserInfo.Manager -eq $null)
{
    $UsersManager = "NONE"
}
Else
{
    # Manager is CN=name, OU=bla bla bla
    $UsersManagerName = SplitOU($ADUserInfo.Manager)
    $UsersManager = Get-ADUser -Filter "Name -eq '$UsersManagerName'"
    $UsersManagerLogin = $UsersManager.SamAccountName
}

#TODO Check for blank?
$OutOfOfficeMessage = "$($ADUserInfo.Name) is no longer with $($ADUserInfo.Company).  Please contact $($UsersManager.UserPrincipalName)"

# Look up the servers
$DirArchive += Get-Date -Format "yyyy"
$DirArchive += "\"

if (!(Test-Path $DirArchive))
{
    New-Item -ItemType directory -Path $DirArchive
    LogThis "Created $DirArchive folder" 
}

$DirArchive += Get-Date -Format "MM"
$DirArchive += "\"

if (!(Test-Path $DirArchive))
{
    New-Item -ItemType directory -Path $DirArchive
    LogThis "Created $DirArchive folder"
}

$DirArchive += "$LoginName"

if (!(Test-Path $DirArchive))
{
    New-Item -ItemType directory -Path $DirArchive
    LogThis "Created $DirArchive folder"
}

$DirUserProfile += "$LoginName.<domain>.v2"
$DirUserDesktop += "$LoginName"
$DirUserDocuments += "$LoginName"


###########################################################
# Any test comands should go here

#Robocopy $DirUserProfile $DirArchive\Profile $RobocopyOptions
#Robocopy $DirUserDesktop $DirArchive\Desktop $RobocopyOptions
#Robocopy $DirUserDocuments $DirArchive\Documents $RobocopyOptions

#exit
###########################################################


LogThis "===================================================================================================================================="
LogThis "Given Name:      $($ADUserInfo.GivenName)"
LogThis "Name:            $($ADUserInfo.Name)"
LogThis "Canonical Name:  $($ADUserInfo.CanonicalName)"
LogThis "Email:           $($ADUserInfo.UserPrincipalName)"
LogThis "Job Title:       $($ADUserInfo.Title)"
LogThis "Reports To:      $($UsersManager.Name)"
LogThis "Random Password: $Password"
LogThis "===================================================================================================================================="
LogThis "Out of Ofice will read as:"
LogThis $OutOfOfficeMessage
LogThis "===================================================================================================================================="
LogThis "User Profile:    $DirUserProfile"
if (!(Test-Path $DirUserProfile)){LogThis "                     Profile directory not found" Red}
LogThis "User Desktop:    $DirUserDesktop"
if (!(Test-Path $DirUserDesktop)){LogThis "                     Desktop directory not found" Red}
LogThis "User Documents:  $DirUserDocuments"
if (!(Test-Path $DirUserDocuments)){LogThis "                     Documents directory not found" Red}
LogThis "Archive Folder:  $DirArchive"
if (!(Test-Path $DirArchive))
{
    LogThis "                     Archive directory not found [DANGER!  It should have been autocreated]" Red
    exit
}
LogThis "===================================================================================================================================="
LogThis

# Check with user, that this really IS the user we want to remove
LogThis "y will answer yes to the questions.  Any other response will assume no" Green
If ($UsersManager -ne "NONE")
{
    LogThis "Will $($UsersManager.GivenName) handle users Inbox / Calendar?" Magenta
    $ConfirmPermissions = Read-Host
    LogThis $ConfirmPermissions Blue
}
LogThis "Are you sure we want to remove this user?" Magenta
$ConfirmUser = Read-Host
LogThis $ConfirmUser Blue

If ($ConfirmUser -eq 'y')
{
    LogThis "Resetting Password for $LoginName"
	Set-ADAccountPassword $LoginName -reset -newpassword (ConvertTo-SecureString -AsPlainText $Password -Force)
    If (!$?) {LogThis $Error[0] Blue}
    LogThis "Forcing a sync"
    Invoke-Expression c:\scripts\force-sync.ps1
}
# If not, quit not
Else
{
    Exit
}

# Update the email with the user name
$EmailSubject += " $LoginName completed by $User"

# Remove any groups the user is in
ForEach ($Group in $ADUserInfo.MemberOf)
{
    $GroupName = SplitOU($Group)
    LogThis "Removing user from $GroupName"
    Remove-ADGroupMember -Identity $GroupName -Member $LoginName -Confirm:$False
    If (!$?) {LogThis $Error[0] Blue}
}

# Add the user into the 'Deny Logins'
LogThis "Adding user to 'Deny Logins'"
Add-ADGroupMember -Identity "Deny Logins" -Member $LoginName -Confirm:$False
If (!$?) {LogThis $Error[0] Blue}

# Set Account To Disabled
LogThis "Setting Sign In Status to Blocked"
Set-MsolUser -UserPrincipalName $ADUserInfo.UserPrincipalName -Blockcredential $true
If (!$?) {LogThis $Error[0] Blue}

# Block active sync
Set-CASMailbox -Identity $LoginName -ActiveSyncEnabled $false -OWAEnabled $false -PopEnabled $false -ImapEnabled $false -MAPIEnabled $false

# Block user's mobile devices
$UserMobileDevices = Get-MobileDevice -Mailbox $LoginName
If ($UserMobileDevices -ne $NULL)
{
    Remove-MobileDevice $($UserMobileDevices.Identity) -Confirm:$False
    If (!$?) {LogThis $Error[0] Blue}
    LogThis "Removed $LoginName's device $($UserMobileDevices.Identity)"
}

# Set out of office
Set-MailboxAutoReplyConfiguration -Identity $LoginName -AutoReplyState Enabled `
            -ExternalMessage $OutOfOfficeMessage `
            -InternalMessage $OutOfOfficeMessage
If (!$?) {LogThis $Error[0] Blue}				 
LogThis "Set up out of office for $LoginName"

# Do the "clever" stuff to find out if mailbox is less than 9500 MB (leaves a little room up to 10 GB)
$stat = Get-MailboxStatistics $ADUserInfo.UserPrincipalName
$tmp = $stat.TotalItemSize.Value.ToString().Split("(")[0].Replace(" ","")
$mb = Invoke-Expression $tmp/1MB

If ([int]$mb -lt 9500)
{
	# Check if user is shared
    LogThis "Checking user $LoginName's current setup"
    $TempMailboxType = Get-Mailbox -Identity $ADUserInfo.UserPrincipalName
    If ($TempMailboxType.RecipientTypeDetails -eq "SharedMailbox")
    {
        LogThis "User $LoginName is already a shared mailbox" Green
    }
    Else
    {
        # Setting the actual mailbox parameters
	    LogThis "Converting user $LoginName to shared and setting quota to 10 GB..."
	    Set-Mailbox -Identity $ADUserInfo.UserPrincipalName -Type "Shared" -ProhibitSendReceiveQuota 10GB -ProhibitSendQuota 9.75GB -IssueWarningQuota 9.5GB
    }
    If (!$?) {LogThis $Error[0] Blue}				 
}
else
{ 
	LogThis "Mailbox is $([int]$mb) MB which is too large for conversion to a nonlicensed shared mailbox, reduce size and try again"
    Exit
}
			
LogThis 'Checking to ensure conversion was successful'

$Count = 0

# Check the mailbox was converted
Do
{
    # Wait 30 seconds after the first try
    If ($Count -ge 1)
    {
        Sleep 30
    }
    $MailboxCheck = Get-Mailbox -Identity $ADUserInfo.UserPrincipalName    
    $Count++
    # And give up after a while!
    if ($Count -gt 5)
    {
        LogThis "Inbox is still $MailboxCheck" Red
        Exit
    }
} While ($MailboxCheck.RecipientTypeDetails -ne "SharedMailbox")
LogThis "Confirmed $LoginName's email is now of type 'SharedMailbox'"

LogThis "Removing Office 365 Licences..."

# Find which licenses are set
$UserLicenses = Get-MsolUser -UserPrincipalName $ADUserInfo.UserPrincipalName
# Skips if no licenses (but doesn't print them out either)
ForEach ($UserLicense in $UserLicenses.Licenses)
{
    Set-MsolUserLicense -UserPrincipalName $ADUserInfo.UserPrincipalName -RemoveLicenses $UserLicense.AccountSkuId
    LogThis "Removing license $UserLicense.AccountSkuId"
}

# Grant permissions to inbox for the manager, if requested
If ($ConfirmPermissions = 'y')
{

    # Complicated to recurse round and find exact permissions and compare to what we want.
    # Much easier to just remove all and add the right ones in

    # Pull existing permissions for manager
    $TempMailboxPermissionsAll = Get-MailboxPermission -Identity $LoginName -User $UsersManagerLogin

    # Loop through them to remove them
    ForEach ($TempMailboxPermissions in $TempMailboxPermissionsAll.AccessRights)
    {
        LogThis "Removing $TempMailboxPermissions to $LoginName for $UsersManagerLogin"
        Remove-MailboxPermission -Identity $LoginName -User $UsersManagerLogin -AccessRights $TempMailboxPermissions -Confirm:$False
        If (!$?) {LogThis $Error[0] Blue}
    }

    # However this doesn't include SendAs because reasons...
    $TempMailboxPermissionsSendAs = Get-RecipientPermission $LoginName -Trustee $UsersManagerLogin
    If ($TempMailboxPermissionsSendAs.AccessRights -ne "SendAs")
    {
        LogThis "Adding permission for $UsersManagerLogin to SendAs $LoginName"
        $Temp = Add-RecipientPermission $LoginName -Trustee $UsersManagerLogin -AccessRights SendAs -Confirm:$False
        If (!$?) {LogThis $Error[0] Blue}
    }
    Else
    {
        LogThis "$($UsersManager.Name) already has access to SendAs from $LoginName" Green
    }
    
    # Add permissions for FullAccess
    LogThis "Adding FullAcces to $LoginName for $UsersManagerLogin"
    $Temp = Add-MailboxPermission -Identity $LoginName -User $UsersManagerLogin -AccessRights FullAccess -Confirm:$False
    If (!$?) {LogThis $Error[0] Blue}

    # Add Calendar Access
    $TempCalendarPermissions = Get-MailboxFolderPermission -Identity "$($LoginName):\Calendar" -User $UsersManagerLogin
    If ($TempCalendarPermissions.AccessRights -eq "Owner")
    {
        LogThis "$($UsersManager.Name) already has access to $LoginName's Calendar" Green
    }
    Else
    {
        Add-MailboxFolderPermission -Identity "$($LoginName):\Calendar" -User $UsersManagerLogin -AccessRights Owner -Confirm:$False
        LogThis "Adding permission for $UsersManagerLogin to $LoginName's Calendar"
    }

}

# Move the files over (if paths exist - have already warned user if they don't)
LogThis "Backing up Profile"
Robocopy $DirUserProfile $DirArchive\Profile $RobocopyOptions
LogThis "Backing up Desktop"
Robocopy $DirUserDesktop $DirArchive\Desktop $RobocopyOptions
LogThis "Backing up Documents"
Robocopy $DirUserDocuments $DirArchive\Documents $RobocopyOptions

LogThis "Account Has Now Been Shutdown" Magenta

$EmailBody += "</p></html></body>"

Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Attachments $EmailAttachment -Subject $EmailSubject -Body $EmailBody -BodyAsHtml
