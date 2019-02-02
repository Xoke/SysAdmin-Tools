Param
(
    [string]$User = $(Read-Host "Enter User's name or login"),
    [int]$Age = 60
)

$Connections = Get-PSSession | Where-Object {$_.State -eq 'Opened'} | Measure-Object
If ($Connections.Count -eq 0)
{
    C:\Scripts\Connect.ps1
}

# Stuff goes here!
$Results = Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(0 - $Age) -EndDate (Get-Date) -Operations UserLoggedIn -UserIds $User | Select-Object -ExpandProperty AuditData | ConvertFrom-Json | Select-Object -ExpandProperty ClientIP | Group-Object -NoElement
$Locations = @()

ForEach ($Result in $Results)
{
    $IP = $Result.Values
    $Locations += Invoke-RestMethod -Method Get -Uri https://freegeoip.app/xml/$($IP) | Select-Object -ExpandProperty Response | Select-Object @{N='User';E={$User}}, City, RegionName, CountryName, IP
}

$Locations | Format-Table

C:\Scripts\Disconnect.ps1
