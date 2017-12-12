##################################################
#
# Check AD user and see if anything seems off
#
# Change the $GroupListAlert if you want it to
# warn you, and the $Domains if you have split
# admin / student domains
# 
##################################################

$Good = "green"
$Bad = "red"
$Iffy = "yellow"
$GroupListAlert = ("SPECIAL GROUP TO WARN ABOUT - ou=company, ou=.com etc")

# Ask for a login name and try to pull information, loop if we error when trying
Do
{

    $Domain = Read-Host "[a]dmin or [s]tudent (a)"
    If ($Domain -eq "s")
    {
        $Domain = "student.domain.edu"
    }
    Else
    {
        $Domain = "admin.domain.edu"
    }

    $LoginName = Read-Host "Please enter the login or name of the account to check"
    $LoginName = $LoginName.Trim()
    $LoginName = $LoginName.Replace(" ", ".") 
    $ADUserInfo = Get-ADUser $LoginName -Properties * -Server $Domain

} While (!$?)

# Do we want to check o365?
$Response = Read-Host "Check o365? (n)"

# If we do, load the modules
If ($Response -eq "y")
{
    C:\Scripts\Connect.ps1
}

# Clear the screen of the o365 rubbish
Clear-Host

$Today = Get-Date

# Show basic stuff
Write-Host "Display Name:     " $ADUserInfo.DisplayName
Write-Host "Login Name:       " $ADUserInfo.SAMAccountName
Write-Host "Created:          " $ADUserInfo.Created
Write-Host "Modified:         " $ADUserInfo.Modified
Write-Host "Info:             " $ADUserInfo.Info
Write-Host "Office:           " $ADUserInfo.Office

# Check if the account is enabled, color response
If ($ADUserInfo.Enabled)
{
    $Colour = $Good
}
Else
{
    $Colour = $Bad
}
Write-Host "Enabled?:          " -NoNewline
Write-Host $ADUserInfo.Enabled -ForegroundColor $Colour

# Check if account is locked out and color accordingly
If ($ADUserInfo.LockedOut)
{
    $Colour = $Bad
}
Else
{
    $Colour = $Good
}
Write-Host "Locked?:           " -NoNewline
Write-Host $ADUserInfo.LockedOut -ForegroundColor $Colour

# If the account is expired (also check for null as that is < today)
If ($ADUserInfo.AccountExpirationDate -lt $Today -and $ADUserInfo.AccountExpirationDate -ne $null)
{
    $Colour = $Bad
}
Else
{
    $Colour = $Good
}
Write-Host "Account Expires:   " -NoNewline
Write-Host $ADUserInfo.AccountExpirationDate -ForegroundColor $Colour

# Have they logged in in the last 7 days
If ($ADUserInfo.LastLogonDate.AddDays(7) -lt $Today)
{
    $Colour = $Iffy
}
Else
{
    $Colour = $Good
}
Write-Host "Last Logon:        " -NoNewline
Write-Host $ADUserInfo.LastLogonDate -ForegroundColor $Colour

# Check PasswordLastSet isn't null (to prevent errors)
If ($ADUserInfo.PasswordLastSet -eq $Null)
{
    $Colour = $Iffy
    Write-Host "Password Set:      " -NoNewline
    Write-Host "NULL" -ForegroundColor $Colour
}
Else
{
    # If they set their password over 90 days ago, OR in the last day
    If ($ADUserInfo.PasswordLastSet.AddDays(90) -lt $Today -or $ADUserInfo.PasswordLastSet.AddDays(1) -gt $Today)
    {
        $Colour = $Iffy
    }
    Else
    {
        $Colour = $Good
    }
    Write-Host "Password Set:      " -NoNewline
    Write-Host $ADUserInfo.PasswordLastSet -ForegroundColor $Colour

    # Is their password expired (based on GPO of 90 days)
    If ($ADUserInfo.PasswordLastSet.AddDays(90) -lt $Today)
    {
        $Colour = $Bad
    }
    Else
    {
        $Colour = $Good
    }
    Write-Host "Password Expires:  " -NoNewline
    Write-Host $ADUserInfo.PasswordLastSet.AddDays(90) -ForegroundColor $Colour
}

# If their password never expires
If ($ADUserInfo.PasswordNeverExpires)
{
    $Colour = $Bad
}
Else
{
    $Colour = $Good
}
Write-Host "Never Expires?:    " -NoNewline
Write-Host $ADUserInfo.PasswordNeverExpires -ForegroundColor $Colour
Write-Host "Account Location: " $ADUserInfo.CanonicalName
Write-Host "Member Of:        "

# Check the groups they are in
ForEach ($GroupList in $ADUserInfo.MemberOf)
{
    # If it is Deny Logins, colour it!
    If ($GroupListAlert -contains $GroupList)
    {
        Write-Host "                  " $GroupList -ForegroundColor $Bad
    }
    Else
    {
        Write-Host "                  " $GroupList
    }
}

# Check Lync
# Does not work on it tools
<#
$LyncUser = Get-CSUser $ADUserInfo.SAMAccountName
If ($LyncUser.Enabled)
{
    $Colour = $Good
}
Else
{
    $Colour = $Bad
}
Write-Host "Lync:              " -NoNewline
Write-Host $LyncUser.Enabled -ForegroundColor $Colour
#>

If ($ADUserInfo.MailNickName -eq $ADUserInfo.SAMAccountName)
{
    $Colour = $Good
}
Else
{
    $Colour = $Bad
}
Write-Host "MailNickName:      " -NoNewline
Write-Host $ADUserInfo.MailNickName -ForegroundColor $Colour

# If we wanted to check o365
If ($Response -eq "y")
{
    # Pull o365 info
    $MSUser = Get-MsolUser -UserPrincipalName $ADUserInfo.EmailAddress
    If (!$?)
    {
        Write-Host "o365:             " -NoNewLine
        Write-Host "Missing" -ForegroundColor $Bad
    }
    Else
    {

        # If they are blocked in o365
        If ($MSUser.BlockCredential)
        {
            $Colour = $Bad
        }
        Else
        {
            $Colour = $Good
        }

        Write-Host "o365 Blocked:      " -NoNewline
        Write-Host $MSUser.BlockCredential -ForegroundColor $Colour

        # If they are licensesd
        If ($MSUser.IsLicensed)
        {
            $Colour = $Good
        }
        Else
        {
            $Colour = $Bad
        }
        Write-Host "o365 Licensed:     " -NoNewline
        Write-Host $MSUser.IsLicensed -ForegroundColor $Colour

        # If they have forwarding on
        $MSUserMailBox = Get-Mailbox $ADUserInfo.EmailAddress
        Write-Host "Forwarding:        " -NoNewline
        Write-Host $MSUserMailBox.ForwardingAddress

        # Check for send quota
        If ($MSUserMailBox.ProhibitSendQuota -eq "49.5 GB (53,150,220,288 bytes)")
        {
            $Colour = $Good
        }
        Else
        {
            $Colour = $Iffy
        }
        Write-Host "Send Quota:        " -NoNewline
        Write-Host $MSUserMailBox.ProhibitSendQuota -ForegroundColor $Colour

        # And OOO
        $MSUserAutoReply = Get-MailboxAutoReplyConfiguration $ADUserInfo.EmailAddress
        Write-Host "Out Of Office:    " $MSUserAutoReply.AutoReplyState
        If ($MSUserAutoReply.InternalMessage.Length -gt 200)
        {
            Write-Host "Internal:          [Very long, skipping]"
        }
        Else
        {
            Write-Host "Internal:         " $($MSUserAutoReply.InternalMessage.Replace("<html>`n<body>`n", "")).Replace("`n</body>`n</html>`n", "")
        }
        If ($MSUserAutoReply.ExternalMessage.Length -gt 200)
        {
            Write-Host "External:          [Very long, skipping]"
        }
        Else
        {
            Write-Host "External:         " $($MSUserAutoReply.ExternalMessage.Replace("<html>`n<body>`n", "")).Replace("`n</body>`n</html>`n", "")
        }

        # If they have IMAP etc
        $MSUserCASMailBox = Get-CASMailbox -Identity $ADUserInfo.UserPrincipalName
        Write-Host "Protocol:          " -NoNewline
        If ($MSUserCASMailBox.ActiveSyncEnabled)
        {
            $Colour = $Good
        }
        Else
        {
            $Colour = $Bad
        }
        Write-Host "ActiveSync    " -ForegroundColor $Colour -NoNewline

        If ($MSUserCASMailBox.PopEnabled)
        {
            $Colour = $Good
        }
        Else
        {
            $Colour = $Bad
        }
        Write-Host "Pop    " -ForegroundColor $Colour -NoNewline

        If ($MSUserCASMailBox.ImapEnabled)
        {
            $Colour = $Good
        }
        Else
        {
            $Colour = $Bad
        }
        Write-Host "IMAP    " -ForegroundColor $Colour -NoNewline

        If ($MSUserCASMailBox.MapiEnabled)
        {
            $Colour = $Good
        }
        Else
        {
            $Colour = $Bad
        }
        Write-Host "MAPI    " -ForegroundColor $Colour -NoNewline

        If ($MSUserCASMailBox.OwaEnabled)
        {
            $Colour = $Good
        }
        Else
        {
            $Colour = $Bad
        }
        Write-Host "OWA" -ForegroundColor $Colour

    }

    $MSUser.Licenses.ServiceStatus | FT -AutoSize

    # And tidy up and close the session down
    C:\Scripts\Disconnect.ps1
}
