<#
-------------------------------------------------
DESCRIPTION
    Check things for a user that are forwarding
    emails to external addresses
-------------------------------------------------
EXAMPLE
    PS C:\> ForwardCheck.ps1
    Check all users
    PS C:\> ForwardCheck.ps1 -User john.smith
    Check john.smith only
-------------------------------------------------
INPUTS
    • -User (optional) username to check
    • -Debug (optional) Display added debug text
    • -Skip (optional) skip disconnecting and email
        (very useful when testing)
-------------------------------------------------
OUTPUTS
    • Email
-------------------------------------------------
NOTES
    General notes
-------------------------------------------------
TODO 
    • Anything to do?
-------------------------------------------------
BUGS
    • Any known bugs?
#>

Param
(
    [string]$User = "",
    [switch]$Debug = $False,
    [switch]$Skip = $False
)

# Setup here
$SMTPServer = "exchange-01"              # Server to use for email
$EmailFrom = "noreply@somecollege.edu"   # Email will use this as from address
$EmailTo = "me@somecollege.edu"          # Who should receive the email
$EmailSubject = "Forwarding Check"       # Subject
$MatchEmails = "*@somecollege.edu"       # Filter out these emails (to skip students)

# See if we have an active connection already
$Connections = Get-PSSession | Where {$_.State -eq 'Opened'} | Measure-Object
If ($Connections.Count -eq 0)
{
    # If not connect then clear the screen
    C:\Scripts\Connect.ps1
    Clear-Host
}

# If a username was not passed in
If ($User -eq "")
{
    # Grab the full list of users (only enabled, only users, filtering emails as per above)
    If ($Debug) {Write-Output "No -user"}
    $Users = Get-MsolUser -EnabledFilter EnabledOnly -All | Where-Object {$_.UserType -eq "Member" -And $_.SignInName -Like $MatchEmails}
}
Else
{
    # If username was used, just pull that user
    If ($Debug) {Write-Output "-User $User"}
    $Users = Get-ADUser $User -Properties Mail | Get-MSOLUser
}

$Log = "Users forwarding emails`n"
# Keep track of how many users (just to display something on screen)
$Count = 1

If ($Debug) {Write-Output "Found $($Users.Count) Users"}

# Go through each user
ForEach ($Person in $Users)
{
    
    # Show either every line if debug, or count every 10 if not
    If ($Debug)
    {
        Write-Output "User $Count of $($Users.Count) [$($Person.UserPrincipalName)]"
    }
    Else
    {
        If ($Count % 10 -eq 0)
        {
            Write-Output "User $Count of $($Users.Count)"
        }
    }

    # Check if user is forwarding emails externally
    # Note, we have this blocked so this shouldn't happen!
    $Forwards = Get-Mailbox $Person.UserPrincipalName | Select-Object ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxAndForward

    If (($Forwards.ForwardingAddress -ne $Null -And `
        $Forwards.ForwardingAddress -like '@') -Or `
        $Forwards.ForwardingSMTPAddress -ne $Null)
    {
        # This should never happen, as forwarding address is internal only, and forwarding smtp address should be disabled
        Write-Output "$($Person.UserPrincipalName) Warning - Forwarding to External email $($Forwards.ForwardingAddress) $($Forwards.ForwardingSMTPAddress)"
        $Log += "$($Person.UserPrincipalName) Warning - Forwarding to External email`n"
    }

    If ($Debug) {Write-Output "Getting Rules"}

    # Pull the rules from the user
    $Rules = Get-InboxRule -Mailbox $Person.UserPrincipalName

    If ($Debug)
    {
        $RuleCount = $Rules.Count
        $CurrentRule = 1
    }

    # Loop through each rule to see if they are forwarding emails
    ForEach ($Rule in $Rules)
    {

        If ($Debug) {Write-Output "Rule $CurrentRule / $RuleCount"}

        # "EX:/" are internal emails, so we match ones not forwarding internally
        If ($Rule.ForwardTo -ne $null -and $Rule.ForwardTo -notmatch "EX:/")
        {
            $Temp = "$($Person.UserPrincipalName) [Enabled: $($Rule.Enabled)] '$($Rule.Name)' => $($Rule.ForwardTo)"
            Write-Output $Temp
            $Log += $Temp
            $Log += "`n"
        }

        If ($Debug) {$CurrentRule += 1}

    }

    $Count++

}

$Log += "`nChecked $($Users.Count) users`n"
$Log += "`n"
$Log += "Checked for external forwards (should be blocked)`n"
$Log += "and for rules forwarding with an @ sign in them (so not internal forward)`n"

$Users = $Null

If (!$Skip)
{
    C:\Scripts\Disconnect.ps1
    Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject $EmailSubject -Body $Log
    Write-Output "Mail Sent"
}
