$Today = Get-Date()
$Computers = Get-ADComputer -Filter * -Properties CN, LastLogonDate, OperatingSystem | Where-Object {$_.LastLogonDate -gt $Today.AddDays(-30) -and $_.OperatingSystem -like 'Windows 7*'}
Write-Host "Computer, User, User Login, Computer Login, OS"
"Computer, User, User Login, Computer Login, OS" | Out-File "Win7.csv"
ForEach ($Computer in $Computers)
{
    $User = Get-ChildItem "\\$($Computer.CN)\c$\Users" | Sort-Object -Descending LastWriteTime | Select-Object -First 1
    Write-Host "$($Computer.CN), $($User.Name), $($User.LastWriteTime), $($Computer.LastLogonDate), $($Computer.OperatingSystem)"
    "$($Computer.CN), $($User.Name), $($User.LastWriteTime), $($Computer.LastLogonDate), $($Computer.OperatingSystem)" | Out-File "Win7.csv" -Append
}
