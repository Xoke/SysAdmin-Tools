# Battery Life Check Script with Color Coding and Logging
# Displays current time, battery percentage, status, estimated time remaining, and logs to CSV

# Global variable to track when we started on battery
$script:batteryStartTime = $null
$script:lastPowerStatus = $null

function Get-TimeOnBattery {
    # Get power events from Windows Event Log
    try {
        # Check current battery status
        $battery = Get-WmiObject Win32_Battery
        $onBattery = $battery.BatteryStatus -eq 1 -or $battery.BatteryStatus -eq 4 -or $battery.BatteryStatus -eq 5
        
        # Track state changes
        if ($null -eq $script:lastPowerStatus) {
            $script:lastPowerStatus = $onBattery
            if ($onBattery) {
                $script:batteryStartTime = Get-Date
            }
        }
        elseif ($script:lastPowerStatus -ne $onBattery) {
            $script:lastPowerStatus = $onBattery
            if ($onBattery) {
                $script:batteryStartTime = Get-Date
            }
            else {
                $script:batteryStartTime = $null
            }
        }
        
        # Calculate time on battery
        if ($onBattery -and $null -ne $script:batteryStartTime) {
            $timeSpan = (Get-Date) - $script:batteryStartTime
            return $timeSpan
        }
        
        # Try to get last power event from event log for more accuracy
        $events = Get-WinEvent -FilterHashtable @{LogName='System'; ID=105,109} -MaxEvents 20 -ErrorAction SilentlyContinue | 
                  Where-Object {$_.Message -like "*battery*" -or $_.Message -like "*AC*"} |
                  Select-Object -First 1
        
        if ($events -and $onBattery) {
            $timeSpan = (Get-Date) - $events.TimeCreated
            if ($timeSpan.TotalMinutes -lt 1440) { # Less than 24 hours
                return $timeSpan
            }
        }
        
        if ($onBattery -and $null -eq $script:batteryStartTime) {
            # If we're on battery but don't know when it started, use current session
            $script:batteryStartTime = Get-Date
            return New-TimeSpan
        }
    }
    catch {
        # Fallback to simple tracking
        if ($null -ne $script:batteryStartTime) {
            return (Get-Date) - $script:batteryStartTime
        }
    }
    
    return $null
}

function Get-BatteryInfo {
    param(
        [bool]$ClearScreen = $true,
        [bool]$ReturnData = $false
    )
    
    if ($ClearScreen) {
        Clear-Host
    }
    
    # Get current time
    $currentTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    if (-not $ReturnData) {
        Write-Host "`n════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host "           BATTERY STATUS MONITOR" -ForegroundColor White
        Write-Host "════════════════════════════════════════════════" -ForegroundColor Cyan
        Write-Host "`n📅 Current Time: " -NoNewline -ForegroundColor Gray
        Write-Host $currentTime -ForegroundColor White
        Write-Host ""
    }
    
    # Get battery information using WMI
    $battery = Get-WmiObject Win32_Battery
    
    if ($null -eq $battery) {
        if (-not $ReturnData) {
            Write-Host "❌ No battery detected. This might be a desktop computer." -ForegroundColor Red
            Write-Host ""
        }
        return $null
    }
    
    # Battery Percentage
    $batteryPercent = $battery.EstimatedChargeRemaining
    
    if (-not $ReturnData) {
        Write-Host "🔋 Battery Level: " -NoNewline -ForegroundColor Gray
        
        # Color code the percentage
        if ($batteryPercent -ge 80) {
            $percentColor = "Green"
            $percentSymbol = "█" * ([Math]::Floor($batteryPercent / 5))
        }
        elseif ($batteryPercent -ge 50) {
            $percentColor = "Yellow"
            $percentSymbol = "█" * ([Math]::Floor($batteryPercent / 5))
        }
        elseif ($batteryPercent -ge 20) {
            $percentColor = "DarkYellow"
            $percentSymbol = "█" * ([Math]::Floor($batteryPercent / 5))
        }
        else {
            $percentColor = "Red"
            $percentSymbol = "█" * ([Math]::Floor($batteryPercent / 5))
        }
        
        Write-Host "$batteryPercent%" -ForegroundColor $percentColor -NoNewline
        Write-Host " [" -NoNewline -ForegroundColor Gray
        Write-Host $percentSymbol -ForegroundColor $percentColor -NoNewline
        $emptyBars = "░" * (20 - $percentSymbol.Length)
        Write-Host $emptyBars -ForegroundColor DarkGray -NoNewline
        Write-Host "]" -ForegroundColor Gray
    }
    
    # Battery Status
    $batteryStatus = switch ($battery.BatteryStatus) {
        1 { "Discharging"; $statusColor = "Yellow" }
        2 { "AC Power (Plugged In)"; $statusColor = "Green" }
        3 { "Fully Charged"; $statusColor = "Green" }
        4 { "Low"; $statusColor = "Red" }
        5 { "Critical"; $statusColor = "Red" }
        6 { "Charging"; $statusColor = "Cyan" }
        7 { "Charging and High"; $statusColor = "Cyan" }
        8 { "Charging and Low"; $statusColor = "Yellow" }
        9 { "Charging and Critical"; $statusColor = "Red" }
        10 { "Undefined"; $statusColor = "Gray" }
        11 { "Partially Charged"; $statusColor = "Yellow" }
        default { "Unknown"; $statusColor = "Gray" }
    }
    
    if (-not $ReturnData) {
        Write-Host "`n⚡ Status: " -NoNewline -ForegroundColor Gray
        Write-Host $batteryStatus -ForegroundColor $statusColor
    }
    
    # Time on Battery
    $timeOnBattery = Get-TimeOnBattery
    $timeOnBatteryString = "N/A"
    
    if ($null -ne $timeOnBattery) {
        $hours = [Math]::Floor($timeOnBattery.TotalHours)
        $minutes = $timeOnBattery.Minutes
        $seconds = $timeOnBattery.Seconds
        
        $timeOnBatteryString = ""
        if ($hours -gt 0) {
            $timeOnBatteryString = "$hours hr "
        }
        $timeOnBatteryString += "{0:00}:{1:00}" -f $minutes, $seconds
        
        if (-not $ReturnData) {
            Write-Host "`n🕐 Time on Battery: " -NoNewline -ForegroundColor Gray
            
            # Color code based on duration
            if ($timeOnBattery.TotalMinutes -lt 30) {
                $durationColor = "Green"
            }
            elseif ($timeOnBattery.TotalMinutes -lt 120) {
                $durationColor = "Yellow"
            }
            else {
                $durationColor = "DarkYellow"
            }
            
            Write-Host $timeOnBatteryString -ForegroundColor $durationColor
        }
    }
    elseif (-not $ReturnData) {
        if ($battery.BatteryStatus -eq 2 -or $battery.BatteryStatus -eq 6 -or $battery.BatteryStatus -eq 3) {
            Write-Host "`n🕐 Time on Battery: " -NoNewline -ForegroundColor Gray
            Write-Host "On AC Power" -ForegroundColor Green
        }
    }
    
    # Estimated Time Remaining
    $timeRemainingString = "N/A"
    
    if (-not $ReturnData) {
        Write-Host "`n⏱️  Time Remaining: " -NoNewline -ForegroundColor Gray
    }
    
    if ($battery.BatteryStatus -eq 2 -or $battery.BatteryStatus -eq 6) {
        # Charging or on AC Power
        if ($battery.EstimatedChargeRemaining -eq 100) {
            $timeRemainingString = "Fully Charged"
            if (-not $ReturnData) {
                Write-Host "Fully Charged" -ForegroundColor Green
            }
        }
        else {
            $timeRemainingString = "Charging"
            if (-not $ReturnData) {
                Write-Host "Charging..." -ForegroundColor Cyan
            }
        }
    }
    elseif ($battery.EstimatedRunTime -eq 71582788) {
        # Special value indicating unknown time
        $timeRemainingString = "Calculating"
        if (-not $ReturnData) {
            Write-Host "Calculating..." -ForegroundColor Gray
        }
    }
    elseif ($null -ne $battery.EstimatedRunTime) {
        $minutes = $battery.EstimatedRunTime
        $hours = [Math]::Floor($minutes / 60)
        $mins = $minutes % 60
        
        $timeString = ""
        if ($hours -gt 0) {
            $timeString = "$hours hour"
            if ($hours -gt 1) { $timeString += "s" }
        }
        if ($mins -gt 0) {
            if ($timeString -ne "") { $timeString += " " }
            $timeString += "$mins minute"
            if ($mins -gt 1) { $timeString += "s" }
        }
        
        $timeRemainingString = $timeString
        
        if (-not $ReturnData) {
            # Color code time remaining
            if ($minutes -ge 120) {
                $timeColor = "Green"
            }
            elseif ($minutes -ge 60) {
                $timeColor = "Yellow"
            }
            elseif ($minutes -ge 30) {
                $timeColor = "DarkYellow"
            }
            else {
                $timeColor = "Red"
            }
            
            Write-Host $timeString -ForegroundColor $timeColor
        }
    }
    else {
        if (-not $ReturnData) {
            Write-Host "Not available" -ForegroundColor Gray
        }
    }
    
    # Additional Information
    $powerSource = if ($battery.BatteryStatus -eq 2 -or $battery.BatteryStatus -eq 6 -or $battery.BatteryStatus -eq 3) {
        "AC Adapter"
    } else {
        "Battery"
    }
    
    $batteryHealth = $null
    if ($null -ne $battery.DesignCapacity -and $battery.DesignCapacity -gt 0) {
        if ($null -ne $battery.FullChargeCapacity -and $battery.FullChargeCapacity -gt 0) {
            $batteryHealth = [Math]::Round(($battery.FullChargeCapacity / $battery.DesignCapacity) * 100, 1)
        }
    }
    
    if (-not $ReturnData) {
        Write-Host "`n────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "Additional Information:" -ForegroundColor White
        
        # Power Source
        Write-Host "  💡 Power Source: " -NoNewline -ForegroundColor Gray
        if ($powerSource -eq "AC Adapter") {
            Write-Host $powerSource -ForegroundColor Green
        }
        else {
            Write-Host $powerSource -ForegroundColor Yellow
        }
        
        # Battery Health (if available)
        if ($null -ne $batteryHealth) {
            Write-Host "  🏥 Battery Health: " -NoNewline -ForegroundColor Gray
            
            if ($batteryHealth -ge 80) {
                Write-Host "$batteryHealth%" -ForegroundColor Green
            }
            elseif ($batteryHealth -ge 60) {
                Write-Host "$batteryHealth%" -ForegroundColor Yellow
            }
            else {
                Write-Host "$batteryHealth%" -ForegroundColor Red
            }
        }
        
        Write-Host "════════════════════════════════════════════════`n" -ForegroundColor Cyan
    }
    
    # Return data for logging if requested
    if ($ReturnData) {
        return @{
            Timestamp = $currentTime
            BatteryPercent = $batteryPercent
            Status = $batteryStatus
            TimeOnBattery = $timeOnBatteryString
            TimeRemaining = $timeRemainingString
            PowerSource = $powerSource
            BatteryHealth = $batteryHealth
        }
    }
}

function Export-BatteryLog {
    param(
        [string]$FilePath = "battery_log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",
        [bool]$Append = $false
    )
    
    $data = Get-BatteryInfo -ClearScreen $false -ReturnData $true
    
    if ($null -eq $data) {
        Write-Host "❌ Unable to get battery data for logging." -ForegroundColor Red
        return
    }
    
    # Create CSV object
    $csvData = [PSCustomObject]@{
        Timestamp = $data.Timestamp
        BatteryPercent = $data.BatteryPercent
        Status = $data.Status
        TimeOnBattery = $data.TimeOnBattery
        TimeRemaining = $data.TimeRemaining
        PowerSource = $data.PowerSource
        BatteryHealth = $data.BatteryHealth
    }
    
    # Check if file exists for appending
    $fileExists = Test-Path $FilePath
    
    if ($Append -and $fileExists) {
        $csvData | Export-Csv -Path $FilePath -NoTypeInformation -Append
    }
    else {
        $csvData | Export-Csv -Path $FilePath -NoTypeInformation
    }
    
    Write-Host "✅ Battery data logged to: " -NoNewline -ForegroundColor Green
    Write-Host $FilePath -ForegroundColor Cyan
}

# Function to continuously monitor battery (optional)
function Start-BatteryMonitor {
    param(
        [int]$RefreshSeconds = 30,
        [bool]$LogToCSV = $false,
        [string]$LogFile = ""
    )
    
    if ($LogToCSV) {
        if ([string]::IsNullOrEmpty($LogFile)) {
            $LogFile = "battery_monitor_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        }
        Write-Host "📝 Logging enabled. File: " -NoNewline -ForegroundColor Green
        Write-Host $LogFile -ForegroundColor Cyan
        Write-Host ""
    }
    
    Write-Host "Starting battery monitor... Press Ctrl+C to stop." -ForegroundColor Yellow
    Write-Host "Refreshing every $RefreshSeconds seconds.`n" -ForegroundColor Gray
    
    while ($true) {
        Get-BatteryInfo
        
        if ($LogToCSV) {
            Export-BatteryLog -FilePath $LogFile -Append $true
        }
        
        Start-Sleep -Seconds $RefreshSeconds
    }
}

# Function for single check with optional logging
function Invoke-SingleCheck {
    param(
        [bool]$LogToCSV = $false
    )
    
    Get-BatteryInfo
    
    if ($LogToCSV) {
        $logFile = "battery_check_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        Export-BatteryLog -FilePath $logFile
    }
}

# Function to view existing log files
function View-LogFiles {
    Write-Host "`n📂 CSV Log Files in Current Directory:" -ForegroundColor Cyan
    Write-Host "────────────────────────────────────────────────" -ForegroundColor DarkGray
    
    $csvFiles = Get-ChildItem -Path "." -Filter "*.csv" | Where-Object {$_.Name -like "*battery*"} | Sort-Object LastWriteTime -Descending
    
    if ($csvFiles.Count -eq 0) {
        Write-Host "No battery log files found." -ForegroundColor Yellow
    }
    else {
        foreach ($file in $csvFiles) {
            $size = [Math]::Round($file.Length / 1KB, 2)
            Write-Host ("  📄 {0,-40} {1,8} KB  {2}" -f $file.Name, $size, $file.LastWriteTime) -ForegroundColor Gray
        }
    }
    Write-Host ""
}

# Main execution
Clear-Host
Write-Host "════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "       BATTERY CHECK SCRIPT v2.0" -ForegroundColor White
Write-Host "════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
Write-Host "Choose an option:" -ForegroundColor White
Write-Host ""
Write-Host "  [1] Check battery once" -ForegroundColor Gray
Write-Host "  [2] Check battery once and log to CSV" -ForegroundColor Gray
Write-Host "  [3] Monitor battery continuously (30-second refresh)" -ForegroundColor Gray
Write-Host "  [4] Monitor battery continuously with CSV logging" -ForegroundColor Gray
Write-Host "  [5] Monitor battery continuously (custom settings)" -ForegroundColor Gray
Write-Host "  [6] View existing log files" -ForegroundColor Gray
Write-Host "  [7] Exit" -ForegroundColor Gray
Write-Host ""

$choice = Read-Host "Enter your choice (1-7)"

switch ($choice) {
    "1" {
        Invoke-SingleCheck -LogToCSV $false
    }
    "2" {
        Invoke-SingleCheck -LogToCSV $true
    }
    "3" {
        Start-BatteryMonitor -RefreshSeconds 30 -LogToCSV $false
    }
    "4" {
        $logFile = Read-Host "Enter log file name (press Enter for auto-generated name)"
        if ([string]::IsNullOrEmpty($logFile)) {
            Start-BatteryMonitor -RefreshSeconds 30 -LogToCSV $true
        }
        else {
            if (-not $logFile.EndsWith(".csv")) {
                $logFile += ".csv"
            }
            Start-BatteryMonitor -RefreshSeconds 30 -LogToCSV $true -LogFile $logFile
        }
    }
    "5" {
        $seconds = Read-Host "Enter refresh rate in seconds (default: 30)"
        if (-not ($seconds -match '^\d+$') -or [int]$seconds -le 0) {
            $seconds = 30
            Write-Host "Using default 30 seconds." -ForegroundColor Yellow
        }
        
        $logChoice = Read-Host "Enable CSV logging? (y/n)"
        $enableLog = $logChoice -eq "y" -or $logChoice -eq "Y"
        
        if ($enableLog) {
            $logFile = Read-Host "Enter log file name (press Enter for auto-generated name)"
            if ([string]::IsNullOrEmpty($logFile)) {
                Start-BatteryMonitor -RefreshSeconds ([int]$seconds) -LogToCSV $true
            }
            else {
                if (-not $logFile.EndsWith(".csv")) {
                    $logFile += ".csv"
                }
                Start-BatteryMonitor -RefreshSeconds ([int]$seconds) -LogToCSV $true -LogFile $logFile
            }
        }
        else {
            Start-BatteryMonitor -RefreshSeconds ([int]$seconds) -LogToCSV $false
        }
    }
    "6" {
        View-LogFiles
        Read-Host "Press Enter to continue"
        & $PSCommandPath  # Re-run the script
    }
    "7" {
        Write-Host "`n👋 Goodbye!" -ForegroundColor Cyan
        exit
    }
    default {
        Write-Host "Invalid choice. Running single check..." -ForegroundColor Yellow
        Invoke-SingleCheck -LogToCSV $false
    }
}