<#

Why am I doing this?  When you enable MFA on o365 accounts, then the user goes into the setup (aka.ms/mfasetup) and changes anything, it makes then 'enforced'
Enforced requires modern authentication, or you have to use an app password.  I don't like this, as we found it broke everything.  One day skype would work
with the regular password, the next day it would need an app password.  Every single user we tried it on demanded we turn it off.

However, if you keep 'enabled' then it does work as expected.  Yes it's a little less secure, but it was the only way I could convince management to even
do this.

To work properly, have a user set up MFA, and THEN enable them.  It won't switch them to enforced and you will have a better day

#>

Param
(
    [string]$User,
    [switch]$Update = $False
)

# Email defaults
# TODO Put real info in here
$SMTPServer = "server"
$EmailFrom = "noreply@company.com"
$EmailTo = "myemail@company.com"
$EmailSubject = "MFA Check"
$EmailBody = ""

# See if we have an active connection
$Connections = Get-PSSession | Where-Object {$_.State -eq 'Opened'} | Measure-Object
If ($Connections.Count -eq 0)
{
    C:\Scripts\Connect.ps1
}

# See if we are running on a single user
If ($User -Ne "")
{
    $Users = Get-MSolUser -UserPrincipalName $User
    $EmailBody += "Single user requested"
}
# Else pull all (matching X domain in case you have several, not matching admin accounts etc
# TODO Fix this how you want, but KEEP the 'enforced' bit
Else
{
    $Users = Get-MsolUser -EnabledFilter EnabledOnly -All | `
            Where-Object {$_.UserType -eq "Member" -And ($_.SignInName -Like "*@domain1.com" -Or `
                            $_.SignInName -Like "*@domain2.com") `
                        -And $_.SignInName -NotLike "admins*@company.com" `
                        -And $_.StrongAuthenticationRequirements.State -EQ 'Enforced'
                        }
    $EmailBody += "Number of enforced users: "
    $EmailBody += $Users.Count
}

$EmailBody += "`n`n"

# If we are updating, change back to 'enabled' not 'enforced'
If (!$Update)
{
    $Users | Select-Object DisplayName, UserPrincipalName, @{N="MFA Status"; E={ $_.StrongAuthenticationRequirements.State}}
    $EmailBody += $Users.DisplayName
}
Else
{

    $AuthenticationRequirements = New-Object "Microsoft.Online.Administration.StrongAuthenticationRequirement"
    $AuthenticationRequirements.RelyingParty = "*"
    $AuthenticationRequirements.State = "Enabled"

    $Users | Set-MSOLUser -StrongAuthenticationRequirements $AuthenticationRequirements

    $EmailBody += $Users.DisplayName

}

# If we found users, fire off the email
If ($Users.Count -GT 0)
{
    Send-MailMessage -SmtpServer $SMTPServer `
                        -From $EmailFrom `
                        -To $EmailTo `
                        -Subject "$EmailSubject" `
                        -Body "$EmailBody" `
                        -BodyAsHtml
    Write-Host Email Sent
}

C:\Scripts\Disconnect.ps1
