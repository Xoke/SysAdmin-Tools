# Close all sessions

$AllSessions = Get-PSSession

ForEach ($CurrentSession in $AllSessions)
{
    Remove-PSSession $CurrentSession
    Write-Host Closing Session $CurrentSession
}
