<#
-------------------------------------------------
DESCRIPTION
    Run various things if someone has put
    their username and password into a
    phishing site
-------------------------------------------------
EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
-------------------------------------------------
INPUTS
    • Top - Show the top sender (can use alone)
    • From - Sender User login, or email (assumes email
        if @ found in input)
    • To - Recipient User login, or email (assumes email
        if @ found in input)
    • Delivered - filter only emails that were
        delivered
    • Status - Used with From to get Status
    • Subject - Subject line of email
    • Reset - Make the user reset their password
    • Delete - delete the listed emails
               (requires -user and -subject)
-------------------------------------------------
OUTPUTS
    • Output (if any)
-------------------------------------------------
NOTES
    General notes
-------------------------------------------------
TODO 
    • Change times to PST?
    • Can we lock this script down so only certain people can run it
-------------------------------------------------
BUGS
    • Can't run delete twice on the same person
#>

Param
(
    [switch]$Help = $False,    
    [switch]$Top = $False,
    [string]$From = "",
    [string]$To = "",
    [switch]$CheckIP = $False,
    [switch]$Delivered = $False,
    [string]$Subject = "",
    [switch]$Reset = $False,
    [switch]$Delete = $False,
    [switch]$Student = $False,
    [switch]$Debug = $False
)

If ($Help)
{
    Write-Output "SpamCheck.ps1"
    Write-Output ""
    Write-Output "Usage: SpamCheck.ps1 <parameters>"
    Write-Output "Parameters: "
    Write-Output "            -Top             Shows the top people sending emails with the same subject line"
    Write-Output "            -From            Shows email from this address.  If @ in parameter assumes email,"
    Write-Output "                                 else assumes login name, and does an AD lookup"
    Write-Output "            -Delivered       Show only emails that were delivered (requires -From)"
    Write-Output "            -Reset           For when an internal user is sending spam, block their account!"
    Write-Output "            -Subject         Subject line to look for"
    Write-Output "            -Student         Toggle student domain"
    Write-Output "            -Delete          Remove emails (with -from and -subject)"
    Write-Output ""
    Write-Output "Examples:   "
    Write-Output "            .\SpamCheck.ps1 -Top"
    Write-Output "                             Show top people sending emails with the same subject line"
    Write-Output "            .\SpamCheck.ps1 -Subject"
    Write-Output "                             Search for emails from anyone with a specific subject"
    Write-Output "                                 (searches for <subject>* so can put partial subject"
    Write-Output "                                 so 'spam' would match 'spam subject line' but not 'email spam'"
    Write-Output "            .\SpamCheck.ps1 -From john.smith"
    Write-Output "                             Show emails from internal user john.smith"
    Write-Output "            .\SpamCheck.ps1 -From john.smith -Delivered"
    Write-Output "                             Show emails from internal user john.smith that are marked as delivered"
    Write-Output "            .\SpamCheck.ps1 -From john.smith@contoso.com"
    Write-Output "                             Show emails from external user john.smith@contoso.com"
    Write-Output "            .\SpamCheck.ps1 -From john.smith@contoso.com -CheckIP"
    Write-Output "                             Show emails from external user john.smith@contoso.com"
    Write-Output "                                 And then for every person who received that email"
    Write-Output "                                 look up which IPs they logged in from"
    Write-Output "            .\SpamCheck.ps1 -From john.smith -Reset"
    Write-Output "                             User john.smith is spamming.  Reset AD password, disable AD"
    Write-Output "                                 account, block o365 account and remove any mobile devices"
    Write-Output "            .\SpamCheck.ps1 -From john.smith -Subject 'spam email'"
    Write-Output "                             Look for any emails from john.smith, with the subject"
    Write-Output "                                 'spam emails' (useful to check before deleting)"
    Write-Output "            .\SpamCheck.ps1 -From john.smith -Subject 'spam email' -Delete"
    Write-Output "                             Search and delete all emails from john smith, with subject"
    Write-Output "                                 line 'spam email' (no wildcards).  Two stage process"
    Write-Output "                                 due to microsoft.  Creates discovery search, then deletes"
    Write-Output "                                 but will ask"

    Exit
}

If ($Student)
{
    $Domain = "student.contoso.com"
}
Else
{
    $Domain = "contoso.com"
}

# Connect if needed
#If ($Help -EQ $False)
#{
    $Connections = Get-PSSession | Where-Object {$_.State -eq 'Opened'} | Measure-Object

    If ($Connections.Count -eq 0)
    {
        C:\Scripts\Connect.ps1
    }

#}

# If -top, show the top senders (internally) only
If ($Top)
{ 
    $TopSenders = Get-MessageTrace -SenderAddress "*@contoso.com" -PageSize 5000 | Group-object SenderAddress, Subject | Where-object {$_.Count -GT 100} | Sort-Object Count -Descending | Select-Object Count, Name | Format-Table -AutoSize

    If (!$Debug) {C:\Scripts\Disconnect.ps1}
    Exit
}

# If subject, but nothing else search for emails with that subject
If ($Subject -NE "" -And !$Delete)
{
    Get-MessageTrace -PageSize 5000 | Where-Object {$_.Subject -Like "$Subject*"} | Select-Object Received, SenderAddress, RecipientAddress, Subject | Format-Table -AutoSize
    If (!$Debug) {C:\Scripts\Disconnect.ps1}
    Exit            
}

# If -From not empty, check for @ sign (external email) or to a get-aduser lookup (internal user)
# Doesn't check for students, but you can just use full student email if you need that
If ($From -NE "")
{
    If ($From.Contains("@"))
    {
        $FromEmail = $From
    }
    Else
    {
        $FromEmail = Get-ADUser $From -Server $Domain | Select-Object -ExpandProperty UserPrincipalName
        If (!$?)
        {
            Write-Host "Failed to find user"
            Exit
        }

    }
}

# Same lookup on To - check for @ sign for full email, or do a lookup on admin staff if not
If ($To -NE "")
{
    If ($To.Contains("@"))
    {
        $ToEmail = $To
    }
    Else
    {
        $ToEmail = Get-ADUser $To -Server $Domain | Select-Object -ExpandProperty UserPrincipalName
        If (!$?)
        {
            Write-Host "Failed to find user"
            Exit
        }

    }
}

# If just a from, no delete or reset, look up the info
If ($FromEmail -NE "" -And !$Delete -And !$Reset)
{
    If ($Delivered)
    {
        Get-MessageTrace -SenderAddress $FromEmail | Where-Object {$_.Status -Eq "Delivered"} | Format-Table -AutoSize
    }
    Else
    {
        Get-MessageTrace -SenderAddress $FromEmail | Format-Table -AutoSize
    }

    # But if they had -Check, do an IP check
    If ($CheckIP)
    {
        $Users = Get-MessageTrace -SenderAddress $FromEmail | Select-Object RecipientAddress | Sort-Object RecipientAddress -Unique

        ForEach ($User in $Users)
        {
            .\CheckIPsO365.ps1 -User $User.RecipientAddress -Age 2 -Foreign
        }

        $Users= $Null

    }

    If (!$Debug) {C:\Scripts\Disconnect.ps1}
    Exit
}

If ($ToEmail -NE "" -And !$Delete -And !$Reset)
{
    If ($Delivered)
    {
        Get-MessageTrace -RecipientAddress $ToEmail | Where-Object {$_.Status -Eq "Delivered"} | Format-Table -AutoSize
    }
    Else
    {
        Get-MessageTrace -RecipientAddress $ToEmail | Format-Table -AutoSize
    }

    If (!$Debug) {C:\Scripts\Disconnect.ps1}
    Exit
}

# If -reset used, reset password and lock account etc
If ($From -NE "" -And $Reset)
{

    # Generate the new password
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
    Set-ADAccountPassword $From -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Server $Domain
    Write-Host "Reset password to $Password for $From"

    # Disable AD Account
    Disable-ADAccount -Identity $From -Server $Domain
    Write-Host "Disabled AD Account for $From"

    # Disable o365 user (should sync, but force it blocked now)
    Set-MsolUser -UserPrincipalName $FromEmail -Blockcredential $true
    Write-Host "Disabled o365 account for $From"

    # Prevent sending emails (as the above two commands don't always immediately take effect)
    # UPDATE - now disabling the account should fix this
    #Set-Mailbox $FromEmail -IssueWarningQuota 0 -ProhibitSendQuota 0
    #Write-Host "Set sending quota to 0 (to stop emails being sent) for $From"

    # Prevent reading emails by blocking protocols
    #Set-CASMailbox -Identity $FromEmail -ActiveSyncEnabled $false -OWAEnabled $false -PopEnabled $false -ImapEnabled $false -MAPIEnabled $false
    #Write-Host "Disabled ActiveSync, OWA, POP3, IMAP, and MAPI $From"

    # Delete mobile devices
    Get-MobileDevice -Mailbox $FromEmail | Remove-MobileDevice -Confirm:$False
    Write-Host "Removed all mobile devices for $From"

    C:\Scripts\Disconnect.ps1
    Exit

}

# If they have subject and user only
If ($Delete -eq $False -And $Subject -NE "" -And $From -NE "")
{
    C:\Scripts\Connect.ps1
    Get-MessageTrace -SenderAddress $FromEmail -PageSize 5000 | Where-Object {$_.Subject -Eq $Subject} | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, FromIP | Format-Table -Autosize
    C:\Scripts\Disconnect.ps1
    Exit
}

If ($Delete -eq $True -And $Subject -NE "" -And $From -NE "")
{

# TODO Remove Inbox rules?
    # Connect
    # Get-InboxRule $FromEmail
    # Select one and remove

    C:\Scripts\ConnectSecurity.ps1

    $SearchName = (Get-Date -Format "yyyy-MM-dd")
    $SearchName += "-"
    $SearchName += $From
    #$SearchName += "2"

    #Write-Host New-ComplianceSearch -Name $SearchName -ExchangeLocation All -ContentMatchQuery "(From:$FromEmail) AND (Subject:$Subject)"

    # Search for the emails to see who they emailed
    Write-Host "Writing recipient list to C:\Scripts\$SearchName.txt"
    Get-MessageTrace -SenderAddress $FromEmail | Where-Object {$_.Subject -like "$Subject"} | Select-Object RecipientAddress | Out-File C:\Scripts\Output\$SearchName.txt
    
    # Make a new compliance search for the spam
    New-ComplianceSearch -Name $SearchName -ExchangeLocation All -ContentMatchQuery "(From:$FromEmail) AND (Subject:$Subject)" | Select-Object CreatedTime, ContentMatchQuery, Status
    
    # Must start the search for it to work
    Start-ComplianceSearch -Identity $SearchName
    
    # See if it has completed (run several times)
    $Running = $True
    
    while ($Running)
    {
        If ($(Get-ComplianceSearch -Identity $SearchName | Select-Object -ExpandProperty Status) -Eq 'Completed')
        {
            $Running = $False
        }
        Else
        {
            Get-ComplianceSearch -Identity $SearchName | Select-Object Name, Status
            Start-Sleep -Seconds 60
        }
    }

    # Show results and check we want to delete
    Get-ComplianceSearch -Identity $SearchName | Select-Object Name, Status, Items
    
    # Delete all emails found
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete | Select-Object CreatedTime, ContentMatchQuery, Status

    # Set up variable to check
    $PurgeName = $SearchName + "_Purge"
    
    # See if the delete has finished (it appends _Purge)
    $Running = $True
    
    While ($Running)
    {

        If ((Get-ComplianceSearchAction -Identity $PurgeName | Select-Object -ExpandProperty Status) -Eq 'Completed')
        {
            $Running = $False
        }
        Else
        {
            Get-ComplianceSearchAction -Identity $PurgeName | Select-Object Name, Status
            Start-Sleep -Seconds 30
        }
    }

    # Audit IPs
    Search-UnifiedAuditLog -StartDate (Get-Date).AddMonths(-1) -EndDate (Get-Date) -Operations UserLoggedIn -UserIds $FromEmail | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty ClientIP | Group-Object -NoElement

}
Else
{
    Write-Host "To delete emails, you must have -Delete, -From, and -Subject"
}

Exit

If (!$Debug) {C:\Scripts\Disconnect.ps1}
