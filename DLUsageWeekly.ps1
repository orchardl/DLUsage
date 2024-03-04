 $thumbprint = "thubprint"
$appID = "app_id"
$org = "org.onmicrosoft.com"
$workingDirectory = "C:\path\to\working\folder\DLusage"

$SPSite = "https://myaapc.sharepoint.com/sites/ComplianceandSecurity"
$SPFolder = "Shared Documents\DL Usage Reports" #change to your own SP Folder

# Get the current date in yyyymmdd format
$dateSuffix = Get-Date -Format "yyyyMMdd"

# Construct the file name with the date suffix
$resultsFile = $workingDirectory + "\" + $dateSuffix + "auditResults.csv"

function Email-Me {
    param (
        [string]$Subject = "",
        [string]$Body = "",
        [string]$Attachment = ""
    )

    $mailParams = @{
        From       = "no-reply@domain.com"
        To         = "your-email@domain.com"
        Subject    = $Subject
        Body       = "DL Usage Report: " + $Body
        SmtpServer = "smtp.domain.com"
    }

    if ($Attachment) {
        $mailParams.Add("Attachment", $Attachment)
    }

    Send-MailMessage @mailParams
}

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

# Step 1: Connect to Exchange Online
# Check if the ExchangeOnlineManagement module is installed and import it
$moduleName = "ExchangeOnlineManagement"
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Write-Log -Level WARN -Message "Attempting to install module: $moduleName"
    Install-Module $moduleName -Force -Scope CurrentUser
}
Import-Module $moduleName

# Attempt to connect to Exchange Online, handle any connection errors
try {
    Connect-ExchangeOnline -CertificateThumbPrint $thumbprint -AppID $appID -Organization $org -ShowBanner:$false -ErrorAction Stop
    Write-Log -Level DEBUG -Message "Connected to Exchange Online"
} catch {
    Write-Log -Level ERROR -Message "Script Error: Connection Failed with error: $_.Exception.Message"
    Email-Me -Subject "Script Error: Connection Failed" -Body $_.Exception.Message
    exit 1
    throw
}

# Get the list of distribution groups
try {
    $DistributionGroups = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop
    Write-Log -Level INFO -Message "Pulled DLs."
} catch {
    Write-Log -Level ERROR -Message "Error in pulling DLs: $_.Exception.Message"
    Email-Me -Subject "Script Error: DL pull failed" -Body $_.Exception.Message
    exit 1
}

# Get the last 7 days with a flattened date
Get-Date | Out-File -FilePath $($workingDirectory + "\tempTimeNow.txt")
$EndDate = Get-Content -Path $($workingDirectory + "\tempTimeNow.txt") -Raw
(Get-Date).AddDays(-7) | Out-File -FilePath $($workingDirectory + "\tempTimeOld.txt")
$StartDate = Get-Content -Path $($workingDirectory + "\tempTimeOld.txt") -Raw
Remove-Item -Path $($workingDirectory + "\tempTimeNow.txt")
Remove-Item -Path $($workingDirectory + "\tempTimeOld.txt")

# Create an array to store results
$Results = @()

# Loop through each distribution group and count emails
try {
    Write-Log -Level DEBUG -Message "Begging message traces..."
    foreach ($DL in $DistributionGroups) {
        try {
            $EmailCount = Get-MessageTrace -RecipientAddress $DL.PrimarySmtpAddress -StartDate $StartDate -EndDate $EndDate -PageSize 5000 | Measure-Object | Select-Object -ExpandProperty Count

            # Create an object with the result
            $ResultObject = [PSCustomObject]@{
                'DistributionGroup' = $DL.DisplayName
                'EmailCount' = $EmailCount
            }

            # Add the object to the results array
            $Results += $ResultObject
        } catch {
            Write-Log -Level ERROR -Message "Error in getting message trace for $DL : $_.Exception.Message"
        }
    }
} catch {
    Write-Log -Level ERROR -Message "Error in looping through DLs collected: $_.Exception.Message"
}
# Export the results to CSV
try {
    $Results | Export-Csv -Path $resultsFile -NoTypeInformation
    Write-Log -Level INFO -Message "Exported message trace results to CSV"
} catch {
    Write-Log -Level ERROR -Message "Error in exporting CSV: $_.Exception.Message"
}

# Check if there are exactly 5 files with the "auditResults.csv" pattern
try {
    $auditFiles = Get-ChildItem -Path $workingDirectory -Filter "*auditResults.csv"
} catch {
    Write-Log -Level ERROR -Message "Error in finding the *auditResults.csv files in $workingDirectory"
}

if ($auditFiles.Count -ge 5) {
    Write-Log -Level DEBUG -Message "5 weeks of audits found. Attempting to aggregate..."
    try {
        
        # Initialize a hashtable to store the aggregated counts
        $aggregatedCounts = @{}

        # Iterate through each CSV file
        foreach ($file in $auditFiles) {
            # Import CSV data from the current file
            $csvData = Import-Csv $file.FullName

            # Iterate through each row in the CSV data
            foreach ($row in $csvData) {
                # Extract DistributionGroup and EmailCount from the row
                $distributionGroup = $row.DistributionGroup
                $emailCount = [int]$row.EmailCount

                # If the DistributionGroup already exists in the hashtable, add the count; otherwise, create a new entry
                if ($aggregatedCounts.ContainsKey($distributionGroup)) {
                    $aggregatedCounts[$distributionGroup] += $emailCount
                } else {
                    $aggregatedCounts[$distributionGroup] = $emailCount
                }
            }
        }

        Write-Log -Level DEBUG -Message "Audits aggregated. Dumping in file and deleting old audits."

        # Create a new CSV file with aggregated counts
        $shortDateSuffix = Get-Date -Format "yyyy-MM"
        $finalResults = $workingDirectory + "\" + $shortDateSuffix + "auditResults.csv"
        $aggregatedCounts.GetEnumerator() | Select-Object Name, Value | Export-Csv -Path $finalResults -NoTypeInformation

        # Archiving current and old reports
        $archivedReports = $workingDirectory + "\ArchivedReports"
        try {
            foreach ($file in $auditFiles) {
                Move-Item -Path $file.FullName -Destination $archivedReports
            }
            Write-Log -Level DEBUG -Message "Moved old reports to archive"
        } catch {
            Write-Log -Level ERROR -Message "Error in moving resulting files to archive: $_"
        }

        try {
            # Check the number of CSV files in the archivedReports folder
            $csvFilesCount = (Get-ChildItem -Path $archivedReports -Filter *.csv).Count

            # Zip files if there are more than 100 CSV files
            if ($csvFilesCount -gt 100) {
                Write-Log -Level INFO -Message "Zipping up old archived reports"
                $zipFileName = $archivedReports + "\ArchivedReports" + $shortDateSuffix + ".zip"

                # Get the list of all CSV files to be zipped
                $filesToZip = Get-ChildItem -Path $archivedReports -Filter *.csv

                # Create the zip file
                Compress-Archive -Path $filesToZip.FullName -DestinationPath $zipFileName -Force

                # Remove the original CSV files
                Remove-Item -Path $filesToZip.FullName -Force
            } else {
                Write-Log -Level DEBUG -Message "Less than 100 csv files in archive. Not compressing yet."
            }
        } catch {
            Write-Log -Level ERROR -Message "Error in archiving log files: $_"
        }

    } catch {
        Write-Log -Level ERROR -Message "Error in aggregating last 5 weeks of audits with error: $_.Exception.Message"
    }

    # Email the resulting file
    try {
        Write-Log -Level DEBUG -Message "Emailing $finalResults"
        Email-Me -Subject "Aggregated 5 weeks of DL usage" -Body "see attachment" -Attachment $finalResults
    } catch {
        Write-Log -Level ERROR -Message "Error in emailing $finalResults with error: $_.Exception.Message"
    }

    # Drop results in SharePoint Site
    try {
        # Check if the ExchangeOnlineManagement module is installed and import it
        $PnPmoduleName = "SharePointPnPPowerShellOnline"
        if (-not (Get-Module -ListAvailable -Name $PnPmoduleName)) {
            Write-Log -Level WARN -Message "Attempting to install module: $moduleName"
            Install-Module $PnPmoduleName -Force -Scope CurrentUser
        }
        Import-Module $PnPmoduleName
    } catch {
        Write-Log -Level ERROR -Message "Error in importing/installing PnP Module: $_.Exception.Message"
    }

    try {
        # Connect to PnPOnline
        Connect-PnPOnline -Url $SPSite -ClientId $appID -Tenant $org -Thumbprint $thumbprint
        Write-Log -Level DEBUG -Message "Connected to $SPSite"
    } catch {
        Write-Log -Level ERROR -Message "Error in connecting to PnPOnline: $_.Exception.Message"
    }
    try {
        # Upload to SP site folder
        Add-PnPFile -Path $finalResults -Folder $SPFolder
        Write-Log -Level INFO -Message "Uploaded $resultsFile to $SPFolder"
    } catch {
        Write-Log -Level ERROR -Message "Error in uploading to PnPOnline: $_.Exception.Message"
    }
    try {
        Start-Sleep 30
        Disconnect-PnPOnline
    } catch {
        Write-Log -Level WARN -Message "Error in disconnecting PnPOnline: $_.Exception.Message"
    }
    Move-Item -Path $finalResults -Destination $archivedReports

} else {
    Write-Log -Level DEBUG -Message "There are not greater than 5 auditResults.csv files in the directory."
}

# Disconnect from Exchange Online session
Start-Sleep 30
try {
    Disconnect-ExchangeOnline -Confirm:$false
} catch {
    Write-Log -Level WARN -Message "Error in disconnecting Exchange Online: $_.Exception.Message"
}

# compress old logs; delete the ones over a year old
# Define the log file path
$theLogFilePath = $workingDirectory + "\app.log"

$today = Get-Date -Format "yyyyMMdd"

# Temporary file paths
$tempFilePath = $workingDirectory + "\temp.log"
$finalFilePath = $workingDirectory + "\ArchivedLogs\app_log" + $today + ".zip"

# Get the total line count of the log file
$totalLines = (Get-Content $theLogFilePath).Count

# Calculate lines to skip (total lines - 1000)
$linesToSkip = $totalLines - 1000

$applog = $workingDirectory + "\app.log"

# Check if the file has more than 1000 lines
if ((Get-Item $applog).length -gt 100000000) {
    # Extract all but the last 1000 lines and save to a temporary file
    Get-Content $theLogFilePath | Select-Object -First $linesToSkip | Set-Content $tempFilePath

    # Compress the temporary file
    Compress-Archive -Path $tempFilePath -DestinationPath "$finalFilePath" -Force

    # Extract the last 1000 lines and overwrite the original log file
    Get-Content $theLogFilePath | Select-Object -Last 1000 | Set-Content $theLogFilePath

    Write-Log -Level INFO -Message "Log file size reduced; old logs compressed to $finalFilePath"
} else {
    Write-Log -Message "The log file has 1000 lines or fewer. No need to split and compress." -Level DEBUG
}

# Clean up the temporary file if it exists
if (Test-Path $tempFilePath) {
    Remove-Item $tempFilePath -Force
}

# Set the directory where your zip files are stored
$targetDirectory = $workingDirectory + "\ArchivedLogs"


# Calculate the date one year ago from today
$oneYearAgo = (Get-Date).AddYears(-1)

# Get a list of all zip files in the target directory
$zipFiles = Get-ChildItem -Path $targetDirectory -Filter "app_log*.zip"

foreach ($file in $zipFiles) {
    # Extract the date part of the file name (assuming format is "app_logyyyyMMdd.zip")
    $dateString = $file.BaseName -replace "app_log", ""

    # Parse the date string into a DateTime object
    try {
        $fileDate = [DateTime]::ParseExact($dateString, "yyyyMMdd", $null)

        # Check if the file date is older than one year ago
        if ($fileDate -lt $oneYearAgo) {
            # Delete the file
            Remove-Item $file.FullName -Force
            Write-Log -Message "Old Archival files found: Deleted file: $($file.FullName)" -Level INFO
        }
    }
    catch {
        Write-Log -Level ERROR -Message "Error in Archival Delete: Could not parse date for file: $($file.Name)"
    }
}  
