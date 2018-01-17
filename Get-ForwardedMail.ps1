<#
.SYNOPSIS
    Retrieve accounts that are forwarding email

.DESCRIPTION
    Checks Exchange Online to identify all mailboxes that are forwarding email

.EXAMPLE 1: Check a single user to see if they are forwarding their email
    Get-ForwardedMail -User john.smith

    .EXAMPLE 2: Check all users to see who is forwarding their email
    Get-ForwardedMail

    .EXAMPLE 3: Check a single user to see if they are forwarding their email and send a notification
    Get-ForwardedMail -User john.smith `
                      -SMTPServer "exchange-01" `
                      -To "me@somecollege.edu" `
                      -From "noreply@somecollege.edu" `
                      -Subject "Forwarding Check" 

                      .EXAMPLE 4: Check all users to see who is forwarding their email, but with a Filter
    Get-ForwardedMail -FilterList "*@somecollege.edu" 

    .EXAMPLE 5: Check all users to see who is forwarding their email, but with a Filter, and send a notification
    Get-ForwardedMail -SMTPServer "exchange-01" `
                      -To "me@somecollege.edu" `
                      -From "noreply@somecollege.edu" `
                      -Subject "Forwarding Check" `
                      -FilterList "*@somecollege.edu" 

.INPUTS
    None
.OUTPUTS
    PSCustomObject
#>

function Get-ForwardedMail {
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1',
                   PositionalBinding=$false,
                   HelpUri = 'http://www.microsoft.com/',
                   ConfirmImpact='Medium')]
    [Alias()]
    [OutputType([String])]
    Param (
        # Specific User to search Forwarding Rules for
        [Parameter(Mandatory=$false,
                   Position=0,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("Mailbox", "Username")] 
        $User,

        # SMTP Server if you want to send a email notification
        [Parameter(Mandatory=$true,
                   Position=1,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='SendEmail')]
        [string]$SMTPServer,

        # From address if you want to send a email notification
        [Parameter(Mandatory=$true,
                   Position=2,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='SendEmail')]
        [string]$From,

        # To address if you want to send a email notification
        [Parameter(Mandatory=$true,
                   Position=3,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='SendEmail')]
        $To,

        # Subject if you want to send a email notification
        [Parameter(Mandatory=$true,
                   Position=4,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='SendEmail')]
        $Subject,

        # List to filter out email address
        [Parameter(Mandatory=$false,
                   Position=5,
                   ValueFromPipelineByPropertyName=$true,
                   ParameterSetName='Parameter Set 1')]
        $FilterList
    )
    
    Begin 
    {
        Write-Verbose 'Checking if there is an active connection already'
        $Connections = Get-PSSession | Where-Object { $_.State -eq 'Opened' } | Measure-Object
        If ($Connections.Count -eq 0)
        {
            Write-Verbose -Message 'Not active connection found.'
            # If not connect then clear the screen
            C:\Scripts\Connect.ps1
            Clear-Host
        }
        
        $AllUsers = @()
        $ReturnObject = @()
        $AlertObject = @()
    }
    
    Process 
    {
        Write-Verbose -Message 'Checking to see if a User was specified'

        if (-not($PSBoundParameters.ContainsKey('User')))
        {
            Write-Verbose -Message 'Retrieving all user accounts who are enabled'
            Write-Debug -Message 'Retrieving all user accounts since no User was specified'
            
            $AllUsers = Get-MsolUser -EnabledFilter EnabledOnly -All | `
                            Where-Object {$_.UserType -eq "Member" -And $_.SignInName -Like $MatchEmails}
        }
        elseif ($PSBoundParameters.ContainsKey(('User')))
        {
            Write-Verbose -Message "Retrieving information about {0}" -f $User
            Write-Debug -Message "Retrieving informaiton about $User"

            $AllUsers = Get-ADUser $User -Properties Mail | Get-MSOLUser
        }

        foreach ($person in $AllUsers)
        {
            $i++
            Write-Progress -Activity 'Searching for all Users' -Status 'Progress:' -PercentComplete (($i / $AllUsers.Count)  * 100)
            
            Write-Debug "Checking User [$($person.UserPrincipalName)]"

            Write-Verbose -Message "Checking if $($person.UserPrincipalName) is forwarding emails externally"

            $Forwards = Get-Mailbox $person.UserPrincipalName | `
                            Select-Object ForwardingAddress, ForwardingSMTPAddress, DeliverToMailboxAndForward

            If (($Forwards.ForwardingAddress -ne $Null -And $Forwards.ForwardingAddress -like '@') -Or $Forwards.ForwardingSMTPAddress -ne $Null)
            {
                Write-Warning -Message "$($person.UserPrincipalName) Warning - Forwarding to External email $($Forwards.ForwardingAddress) $($Forwards.ForwardingSMTPAddress)"

                Write-Verbose -Message 'Identified internal account is forwarding to an internal address'

                $props = @{
                    User                  = $person.UserPrincipalName
                    ForwardingAddress     = $Forwards.ForwardingAddress
                    ForwardingSMTPAddress = $Forwards.ForwardingSMTPAddress
                }

                $tempObject = New-Object -TypeName PSCustomObject -Property $props
                $AlertObject += $tempObject
            }

            Write-Debug -Message 'Getting Rules'

            Write-Verbose -Message "Getting Rules for $($Person.UserPrincipalName)"
            $Rules = Get-InboxRule -Mailbox $Person.UserPrincipalName

            Write-Verbose -Message "$($Person.UserPrincipalName) has $($Rules.Count) Forwarding Rules!"

            ForEach ($Rule in $Rules)
            {
                Write-Verbose -Message 'Iterating through all found rules'
                Write-Debug -Message "Looking through Rule $CurrentRule out of $RuleCount"

                # "EX:/" are internal emails, so we match ones not forwarding internally
                If ($Rule.ForwardTo -ne $null -and $Rule.ForwardTo -notmatch "EX:/")
                {
                    Write-Verbose -Message 'Identified Rule Match'
                    $props = @{
                        User          = $person.UserPrincipalName
                        EnabledStatus = $Rule.Enabled
                        RuleName      = $Rule.Name
                        ForwardTo     = $Rule.ForwardTo
                    }
                    
                    Write-Verbose "$($Person.UserPrincipalName) [Enabled: $($Rule.Enabled)] '$($Rule.Name)' => $($Rule.ForwardTo)"

                    $tempObject = New-Object -TypeName PSCustomObject -Property $props
                    $ReturnObject += $tempObject
                }
                else 
                {
                    Write-Verbose -Message "$($person.UserPrincipalName) does not have any identified rules"
                }
            } # End of Foreach Rules
        } # End of Foreach All Users
    } # End of Process Block
    
    End 
    {
        if ($PSCmdlet.ParameterSetName -eq 'SendEmail')
        {
            Write-Verbose -Message 'Checking for any warning accounts'
            if ($AlertObject)
            {
                Write-Verbose -Message 'Warning Accounts Are Present!  Sending Email Notification!'
                $MailProps = @{
                    SmtpServer = $SMTPServer
                    From       = $EmailFrom
                    To         = $EmailTo
                    Subject    = $EmailSubject
                    Body       = $AlertObject
                }
                Send-MailMessag @MailProps
                Write-Verbose -Message 'Email Sent!'

                Write-Output $AlertObject
            }
            
            Write-Verbose -Message 'Checking for accounts with Forwarded Email Rules'
            if($ReturnObject)
            {
                Write-Verbose -Message 'Accounts Identified with Forwarding Rules!  Sending Email Notification!'
                $MailProps = @{
                    SmtpServer = $SMTPServer
                    From       = $EmailFrom
                    To         = $EmailTo
                    Subject    = $EmailSubject
                    Body       = $ReturnObject
                }
                Send-MailMessag @MailProps
                Write-Verbose -Message 'Email Sent!'

                Write-Output $ReturnObject
            }
        } # End of if statement checking for SendEmail ParameterSet
        else 
        {
            Write-Verbose -Message 'Function is NOT sending an email'
            
            if($AlertObject)
            {
                Write-Verbose -Message 'Outputting Warning Accounts!'
                Write-Output -InputObject $AlertObject
            }

            if($ReturnObject)
            {
                Write-Verbose -Message 'Outputting Accounts Identified with Forwarding Rules!'
                Write-Output -InputObject $ReturnObject
            }
        } # End of writing Output to console
    } # End of End block
} # End of Get-ForwardedEmail Function
