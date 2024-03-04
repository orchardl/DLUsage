 $workingDirectory = "C:\path\to\working\directory\DLusage"

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO",

        [Parameter(Mandatory=$false)]
        [string]$LogFilePath = $workingDirectory + "\app.log"
    )

    Begin {
        # Check if log file directory exists, if not, create it
        $logFileDirectory = Split-Path -Path $LogFilePath -Parent
        if (-not (Test-Path -Path $logFileDirectory)) {
            New-Item -ItemType Directory -Path $logFileDirectory | Out-Null
        }
    }

    Process {
        # Format the log entry
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp [$Level] $Message"

        # Write the log entry to the log file
        Add-Content -Path $LogFilePath -Value $logEntry
    }

    End {
        if ($Level -eq "ERROR") {
            Write-Host "An error has been logged to $LogFilePath" -ForegroundColor Red
        } else {
            Write-Host "Log entry added to $LogFilePath" -ForegroundColor Green
        }
    }
}

$date = Get-Date

# Start the main script as a job
$job = Start-Job -FilePath "C:\path\to\working\directory\DLusage\DLUsageWeekly.ps1"
$jobID = $job.Id
Write-Log -Level DEBUG -Message "Starting job, $jobID at $date"

# Wait for up to 1 hour (3600 seconds)
Start-Sleep -Seconds 3600

# Check if the job is still running
if (Get-Job -Id $job.Id -State Running) {
    # Stop the job because it exceeded the time limit
    Stop-Job -Id $job.Id
    Remove-Job -Id $job.Id
    Write-Log -Level ERROR -Message "The script has been terminated because it exceeded the 1-hour time limit."
    Send-MailMessage -From "Script Run Failure <Error-Email-Admin@domain.com>" -To my-user@domain.com -Subject "Terminal Failure: DL Usage Report" -Body "The script has been terminated because it exceeded the 1-hour time limit." -SmtpServer "smtp.domain.com"
} else {
    # Collect the job's results if it finished within the time limit
    $result = Receive-Job -Id $job.Id
    Write-Log -Level INFO -Message "Job, $jobID started at $date was completed within time limit."
    Remove-Job -Id $job.Id
}
 
