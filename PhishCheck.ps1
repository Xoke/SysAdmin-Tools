<#
-------------------------------------------------
DESCRIPTION
    Check things for a user that may have been phished
    (forwarding emails externally)
-------------------------------------------------
EXAMPLE
    PS C:\> PhishCheck.ps1
    Check all users
    PS C:\> PhishCheck.ps1 -User john.smith
    Check john.smith only
-------------------------------------------------
INPUTS
    • -User (optional) username to check
    • -Force (optional) N/A currently
    • -Debug (optional) Display added debug text
    • -Skip (optional) skip connecting / disconnecting
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
    [switch]$Force = $False,
    [switch]$Debug = $False,
    [switch]$Skip = $False
)

$SMTPServer = "exchange-01"
$EmailFrom = "noreply@company.com"
$EmailTo = "me@company.com"
$EmailSubject = "Phishing Check"

If (!$Skip)
{
    $Connections = Get-PSSession | Where {$_.State -eq 'Opened'} | Measure-Object
    If ($Connections.Count -eq 0)
    {
        C:\Scripts\Connect.ps1
        Clear-Host
    }
}
Else
{
    If ($Debug) {Write-Output "Skipping connection"}
}

If ($User -eq "")
{
    If ($Debug) {Write-Output "No -user"}
    $Users = Get-ADUser -Filter {Mail -Like "*"} -Properties Mail
}
Else
{
    If ($Debug) {Write-Output "-User $User"}
    $Users = Get-ADUser $User -Properties Mail
}

$Log = "Potential Phished Users`n"

$Max = $Users.Count + 1
$Count = 1

If ($Debug) {Write-Output "Found $Max Users"}

# Go through each user
ForEach ($Person in $Users)
{

    $Mail = $Person.Mail    
    Write-Output "User $Count of $Max [$Mail]"

    # Check if they have mail set up
    If ($Mail -eq $Null)
    {
        # Don't do anything!
    }
    Else
    {
        $Forwards = Get-Mailbox $Mail | Select-Object ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxAndForward

        If (($Forwards.ForwardingAddress -ne $Null -And `
            $Forwards.ForwardingAddress -like '@') -Or `
            $Forwards.ForwardingSMTPAddress -ne $Null)
        {
            # This should never happen, as forwarding address is internal only, and forwarding smtp address should be disabled
            Write-Output "$Mail Warning - Forwarding to External email"
            $Log += "$Mail Warning - Forwarding to External email`n"
        }

        If ($Debug) {Write-Output "Getting Rules"}

        $Rules = Get-InboxRule -Mailbox $Mail

        If ($Debug)
        {
            $RuleCount = $Rules.Count
            $CurrentRule = 1
        }

        ForEach ($Rule in $Rules)
        {

            If ($Debug) {Write-Output "Rule $CurrentRule / $RuleCount"}

            If ($Rule.ForwardTo -ne $null -and $Rule.ForwardTo -notmatch "EX:/")
            {
                $Temp = "$($Person.SamAccountName) [Enabled: $($Rule.Enabled)] '$($Rule.Name)' => $($Rule.ForwardTo)"
                Write-Output $Temp
                $Log += $Temp
                $Log += "`n"
            }

            If ($Debug) {$CurrentRule += 1}

        }
    }

    $Count++

}

$Log += "`nChecked $Max users`n"
$Log += "`n"
$Log += "Checked for external forwards (should be blocked)`n"
$Log += "and for rules forwarding with an @ sign in them (so not internal forward)`n"

$Users = $Null

If (!$Skip)
{
    C:\Scripts\Disconnect.ps1
    Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject $EmailSubject -Body $Log
    Write-Output Mail Sent
}
