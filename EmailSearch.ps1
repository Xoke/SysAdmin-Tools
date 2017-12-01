##################################################
#
# Search for an email
# 
##################################################
#
# To Do
#
# • Change email search to allow wildcards
# • Speed it up?
# • Make it less hacky!
# • Stop using write-hosts as per http://www.jsnover.com/blog/2013/12/07/write-host-considered-harmful/
#
##################################################
#
# Bugs
#
# • None known
#
##################################################
# Display the menu on the screen
Function Show_Menu()
{
    Clear-Host

    Write-Host
    Write-Host "--== MENU ==--" -Foreground Green
    Write-Host

    Write-Host "1: Sender    = $Sender"
    Write-Host "2: Recipient = $Recipient"
    Write-Host "3: Subject   = $Subject"

    # Print the start / end if we have a start date only
    If ($StartDate.Length -gt 0)
    {
        Write-Host "4: Time      = $StartDate to $EndDate"
    }
    Else
    {
        Write-Host "4: Time      = "
    }
    
    Write-Host "5: Status    = $Status"
    Write-Host
    Write-Host "s: Search for emails"
    Write-Host "Warning - searching for subject has to download all emails"
    Write-Host "and then filter them, so it is very slow"
    Write-Host

    # If we have results, put the option to show them again on screen
    # We wipe the results out when we exit to try to clear the memory
    If ($Results.Count -gt 0)
    {
        Write-Host "v: View results again"
        Write-Host
    }

    Write-Host "x: Exit"
    Write-Host

}

# Main search function
Function Search_Emails
{

    # Start to build up the command line
    $CommandLine = @{}

    If ($Sender.Length -gt 0)
    {
        $CommandLine += @{"SenderAddress" = "$Sender"}
    }

    If ($Recipient.Length -gt 0)
    {
        $CommandLine += @{"RecipientAddress" = "$Recipient"}
    }

    If ($StartDate.Length -gt 0)
    {
        $CommandLine += @{"StartDate" = "$StartDate"}
        $CommandLine += @{"EndDate" = "$EndDate"}
    }    

    If ($Status.Length -gt 0)
    {
        $CommandLine += @{"Status" = "$Status"}
    }  

    # If we have something in the command line use it so the subject search works correctly
    # Also pull many pages of results to get everything for subject - we pull all emails then filter for subject

    $CommandLine += @{"PageSize" = "5000"}

    $Results = $NULL

    For ($c = 1; $c -lt 1001; $c++)
    {

        Write-Host Page $c

        If ((Get-MessageTrace @CommandLine -Page $c).Count -gt 0)
        {
            If ($CommandLine.Count -eq 0)
            {
                If ($c -eq 1)
                {
                    $Results = Get-MessageTrace -Page $c
                }
                Else
                {
                    $Results += Get-MessageTrace -Page $c
                }
            }
            Else
            {
                If ($c -eq 1)
                {
                    $Results = Get-MessageTrace @CommandLine -Page $c
                }
                Else
                {
                    $Results += Get-MessageTrace @CommandLine -Page $c
                }
            }
        }
        Else
        {
            #$c = 1000
            Break
        }
    }
    
    # If we have a subject we are searching for, filter it
    If ($Subject.Length -eq 0)
    {
        #$Results | Select Received, SenderAddress, FromIP , RecipientAddress, ToIP, Subject, Status | ft -AutoSize
        $Results | Out-GridView
    }
    Else
    {
        #$Results | Where {$_.Subject -like $Subject} | Select Received, SenderAddress, FromIP , RecipientAddress, ToIP, Subject, Status | ft -AutoSize
        $Results | Where-Object {$_.Subject -like $Subject} | Out-GridView
    }

}

C:\Scripts\Connect.ps1

$Loop = $True

While ($Loop)
{

    Show_Menu

    $MenuChoice = Read-Host "Select number to change"


    Switch ($MenuChoice)
    {
        1 {$Sender = Read-Host "Sender"}
        2 {$Recipient = Read-Host "Recipient"}
        3 {$Subject = Read-Host "Subject (wildcards OK e.g. *password*)"}

        4
        {
            [int]$TimeFrom = Read-Host "From how many hours ago?"

            If ($TimeFrom -eq 0)
            {
                $StartDate = ""
                $EndDate = ""
            }
            Else
            {
                [int]$TimeTo = Read-Host "To how many hours ago?"

                If ($TimeFrom -gt $TimeTo)
                {
                    $Now = Get-Date
                    $StartDate = $Now.ToUniversalTime()
                    $EndDate = $StartDate
                    $StartDate = $StartDate.AddHours(0 - $TimeFrom)
                    $EndDate = $EndDate.AddHours(0 - $TimeTo)
                }
                Else
                {
                    Write-Host "From must be greater than To"
                    Pause
                }  
            }          

        }

        5
        {
            $Response = Read-Host "Status (D)elivered, (F)ailed, (Q)uarantined, or anything else for all"
            Switch ($Response)
            {
                d {$Status = "delivered"}
                f {$Status = "failed"}
                q {$Status = "quarantined"}
                default {$Status = ""}
            }
        }

        s
        {
            Search_Emails
            #Pause
        }

        s
        {
            $Results | Out-GridView            
        }

        x
        {
            $Loop = $False
            Write-Host "Exiting"
        }

        default {Write-Host "Unknown choice"}

    }
   

}

$Results = $NULL

C:\Scripts\Disconnect.ps1
