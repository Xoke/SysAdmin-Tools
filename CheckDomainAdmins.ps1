<#
-------------------------------------------------
DESCRIPTION
    Schedule this to run, and it checks there
    are no extra accounts with admin rights
-------------------------------------------------
EXAMPLE
    PS C:\>.\CheckDomainAdmins.ps1
    Check against good list, and email results
-------------------------------------------------
INPUTS
    • None
-------------------------------------------------
OUTPUTS
    • Email
-------------------------------------------------
NOTES
    Set up a scheduled task daily, hourly or
    whatever you are comfortable with
-------------------------------------------------
FUTURE
    • No plans
-------------------------------------------------
    • No known bugs
#>

Function GroupCheck ($GroupName, $ExpectedUsers)
{
    #Write-Host $GroupName
    #Write-Host $ExpectedUsers

    $Email += "<font face='arial' color= 'Black'>Unexpected $($GroupName):</font></br>"
    $Email += "<font face='arial' color= 'Red'>"
    $Email += Get-ADGroupMember -Identity $GroupName | `
                        Where-Object {$_.SamAccountName -NotIn $ExpectedUsers} | `
                        Sort-Object -Property SamAccountName -Unique | `
                        Select-Object -ExpandProperty SamAccountName
    $Email += "</font></br></br>"

    Return $Email

}

$SMTPServer = "mail-server"
$EmailFrom = "noreply@iffy.link"
$EmailTo = "someone@iffy.link", "another@iffy.link"
$EmailSubject = "Full admin list"
$EmailBody = "<html><body><p>"

$Ent_Admins = @("ent-admin1", "ent-admin2")				# List who is what here to remind you!
$Dom_Admins = @( "dom-admin1", "dom-admin2", `				# Split this into teams if it helps
                "washington-admin1", "washington-admin2", `		# The goal is to make it readable to YOU
                "florida-admin1", "florida-admin2", `			# The goal is to make it readable to YOU
                "service-1", "service-2", `          			# Service Accounts?
                "vendir-1", "vendor-2", `				# Vendor Accounts?
                "honey.pot" `						# Honeypot?
                )
$Sch_Admins = @("Enterprise Admins")
$HyperV_Admins = @("")
$Srv_Ops = @("Domain Admins")
$Acc_Ops = @("")
$Back_Ops = @("")
$Cert_Pub = @("CERT-SERVER$")
$Key_Admins = @("")
$Ent_Key_Admins = @("")

$EmailBody += GroupCheck 'Enterprise Admins' $Ent_Admins
$EmailBody += GroupCheck 'Domain Admins' $Dom_Admins
$EmailBody += GroupCheck 'Schema Admins' $Sch_Admins
$EmailBody += GroupCheck 'Hyper-V Administrators' $HyperV_Admins
$EmailBody += GroupCheck 'Server Operators' $Srv_Ops
$EmailBody += GroupCheck 'Account Operators' $Acc_Ops
$EmailBody += GroupCheck 'Backup Operators' $Back_Ops
$EmailBody += GroupCheck 'Cert Publishers' $Cert_Pub
$EmailBody += GroupCheck 'Key Admins' $Key_Admins
$EmailBody += GroupCheck 'Enterprise Key Admins' $Ent_Key_Admins
#$EmailBody += GroupCheck '' $

$EmailBody += "<font face='arial' color= 'Black'>Expected Enterprise Admins</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Ent_Admins</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Domain Admins:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Dom_Admins</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Schema Admins:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Sch_Admins</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Hyper-V Administrators:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$HyperV_Admins</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Server Operators:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Srv_Ops</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Account Operators:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Acc_Ops</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Backup Operators:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Back_Ops</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Cert Publishers:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Cert_Pub</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Key Admins:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Key_Admins</font></br></br>"

$EmailBody += "<font face='arial' color= 'Black'>Expected Enterprise Key Admins:</font></br>"
$EmailBody += "<font face='arial' color= 'Green'>$Ent_Key_Admins</font></br></br>"

Send-MailMessage -SmtpServer $SMTPServer -From $EmailFrom -To $EmailTo -Subject $EmailSubject -Body $EmailBody -BodyAsHtml
