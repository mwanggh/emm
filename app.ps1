# Define the list of programs to monitor
$programs = @("chrome", "Code")

# Log file path (create in the script's directory)
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "efficiency_monitor.log"

# Overwrite log file at the start
Set-Content -Path $logFile -Value ""

# Indicate script start
Write-Host "Process monitoring script started." -ForegroundColor Green
Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Script started."

# Infinite loop to monitor processes
while ($true) {
    foreach ($program in $programs) {
        # Get all processes for the current program
        $processes = Get-Process -Name $program -ErrorAction SilentlyContinue

        foreach ($process in $processes) {
            # Check if the process has not exited and its priority is Idle
            if (-not $process.HasExited -and $process.PriorityClass -eq [System.Diagnostics.ProcessPriorityClass]::Idle) {
                
                # Change the process priority to Normal
                $process.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::Normal

                # Log process information to the log file
                $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Changed priority: $program (PID: $($process.Id)) $($process.MainWindowTitle)"
                Add-Content -Path $logFile -Value $logMessage

                # Output to console
                Write-Host $logMessage -ForegroundColor Cyan
            }
        }
    }

    # Sleep for 0.5 seconds before the next check
    Start-Sleep -Milliseconds 200
}