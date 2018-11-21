##################################################
#
# Show all quarantined items for a user
# and allow release of specific emails
# 
# Note:  It shows items already released :(
# 
# Bugs:  Doesn't work for students
#
##################################################

Param
(
    [string]$LoginName = $(If (!$Help) {Read-Host "Enter User's login"})
)

# Ask for a login name and try to pull information, loop if we error when trying
If ($Null -eq $LoginName)
{
Do
    {
        $LoginName = Read-Host "Please enter the login or name of the account to check"
        $LoginName = $LoginName.Trim()
        $LoginName = $LoginName.Replace(" ", ".") 
        $ADUserInfo = Get-ADUser $LoginName -Properties *
    } While (!$?)
}

# Get their email address
$EmailAddress = $ADUserInfo.EmailAddress

# Connect to o365
C:\Scripts\Connect.ps1

# Pull any quarantined emails to this user
$Quarantined = Get-MessageTrace -RecipientAddress $EmailAddress -Status Quarantined

# TODO Check this later - the status tag stopped working
#$Quarantined = Get-MessageTrace -RecipientAddress $EmailAddress | Where {$_.Status -eq "Quarantined"}

# Loop around
Do
{

    # Clear the screen so there is no confusion
    Clear-Host

    # Display an email number
    $Count = 0

    Write-Host "Date`t`t`t`t`t" -NoNewline -ForegroundColor Cyan
    Write-Host "No`t" -NoNewline -ForegroundColor Green
    Write-Host "Sender`t" -ForegroundColor DarkYellow

    # Loop through each email
    ForEach ($Email in $Quarantined)
    {
        # Print info out
        Write-Host $EMail.Received"`t" -NoNewline -ForegroundColor Cyan
        Write-Host $Count"`t" -NoNewline -ForegroundColor Green
        Write-Host $Email.SenderAddress"`t" -ForegroundColor DarkYellow
        $Count++
    }

    # Find out which one to release
    $Release = Read-Host "Release which messages? (return to exit)"

    # If nothing was entered exit (only way out)
    If ($Release -eq "" -or $Release -eq $null)
    {
        Break
    }

    # Release the hounds.. er.. message
    Get-QuarantineMessage -MessageID $Quarantined[$Release].MessageID | Release-QuarantineMessage -User $EmailAddress
    # Remove the released message from the list
    $Quarantined = $Quarantined -ne $Quarantined[$Release]

} While ($true)

# And disconnect
C:\Scripts\Disconnect.ps1
