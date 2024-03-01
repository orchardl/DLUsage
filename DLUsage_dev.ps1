 # PowerShell Script for Managing Distribution Lists and Historical Searches
$thumbprint = "your thumbprint"
$appID = "<appID>"
$org = "conosco.onmicrosoft.com"
$workingDirectory = "C:\path\to\you\directory\DLusage"

$SPSite = "https://myaapc.sharepoint.com/sites/ComplianceandSecurity"
$SPFolder = "Shared Documents\DL Usage Reports" #change to your own SP Folder

$archiveFolder = $workingDirectory + "\DLUsageArchive"
$DLsFile = $workingDirectory + "\DLs.txt"
$DLIDfile = $workingDirectory + "\DLID.csv"
$resultsFile = $workingDirectory + "\auditResults.csv"

function Email-Me {
    param (
        [Parameter(Mandatory=$false)]
        [string]$Body,

        [Parameter(Mandatory=$false)]
        [string]$Subject
    )

    Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject $Subject -Body $Body -SmtpServer "smtp.example.com"

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

# Step 0: Check and Archive Existing Files
$filesToCheck = @($DLsFile, $DLIDfile, $resultsFile) + (Get-ChildItem $($workingDirectory + "\DLgroup*.txt"))
Write-Log -Level DEBUG -Message "Archiving old reports and files."
if (-not (Test-Path -Path $archiveFolder)) {
    New-Item -ItemType Directory -Path $archiveFolder
}
foreach ($file in $filesToCheck) {
    if (Test-Path -Path $file) {
        Move-Item -Path $file -Destination $archiveFolder
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
    throw
}

# Step 2: Get a list of DLs and save them to a file
# Redirect any errors in getting the Distribution Group to an email notification
try {
    Write-Log -Level DEBUG -Message "Collecting DLs"
    (Get-DistributionGroup -ErrorAction Stop).PrimarySmtpAddress > $DLsFile
} catch {
    Write-Log -Level ERROR -Message "Script Error: Fetching DLs Failed with error $_.Exception.Message"
    Email-Me -Subject "Script Error: Fetching DLs Failed" -Body $_.Exception.Message
    throw
}

# Step 3: Divide the DLs into files containing 250 entries
$numfiles = 0
try {
    Write-Log -Level DEBUG -Message "Dividing DLs"
    $DLs = Get-Content -Path $DLsFile
    $numfiles = [math]::Ceiling($DLs.length / 250)
    for ($i = 1; $i -le $numfiles; $i++) {
        $a = ($i - 1) * 250
        $b = [math]::Min($i * 250 - 1, $DLs.length - 1)
        $DLs[$a..$b] > ~\DLgroup$i.txt
    }
} catch {
    Write-Log -Level ERROR -Message "Script Error: Dividing DLs Failed with error $_.Exception.Message"
    Email-Me -Subject "Script Error: Dividing DLs Failed" -Body $_.Exception.Message
    throw
}

# Step 4: Start Historical Search every day until done
Email-Me -Subject "Starting DL Audit" -Body "This will probably take a few days. Please do NOT restart until it completes."
Write-Log -Level INFO -Message "Starting Historical Searches..."
New-Item $DLIDfile -Force
for ($j = 1; $j -le $numfiles; $j++) {
    Write-Log -Level INFO -Message "Starting Historical Searches for day $j"
    try {
        Get-Content -Path $($workingDirectory + "\DLgroup$j.txt") | %{
            Start-HistoricalSearch -ReportTitle "Day $j" -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) -ReportType MessageTrace -RecipientAddress $_ -ErrorAction Stop
            Start-Sleep -Milliseconds 500
        } | Export-CSV -Path $DLIDfile -Append
    } catch {
        Email-Me -Subject "Script Error: Historical Search Day $j Failed" -Body $_.Exception.Message
        Write-Log -Level ERRO -Message "Script Error: Historical Search Day $j Failed."
        throw
    }
    Write-Log -Level INFO -Message "Searchs for day $j finished."
    if ($j -lt $numfiles) {
        Start-Sleep -Seconds 86401
    }
}
Write-Log -Level INFO -Message "Finshed Searches, now lets get the searches"

# Wait for 1 hour to ensure all searches are completed
Start-Sleep -Seconds 3600

# Step 5: Get Historical Search results
$days = $numfiles + 1
try {
    Get-HistoricalSearch | 
        Where-Object {$_.SubmitDate -gt (Get-Date).AddDays(-$days)} | 
        ForEach-Object {
            $jobDetails = Get-HistoricalSearch -JobID $_.JobID
            [PSCustomObject]@{
                'DLname' = $jobDetails.RecipientAddress
                'ReportStatus' = $jobDetails.ReportStatusDescription
                'EmailsSent' = $jobDetails.Rows
            }
        } | Export-CSV $resultsFile -Append -NoTypeInformation
    Write-Log -Level INFO -Message "Pulled Historical Searches and dropped them in $resultsFile"
} catch {
    Write-Log -Level ERROR -Message "Script Error: Exporting Results Failed $_.Exception.Message"
    Email-Me -Subject "Script Error: Exporting Results Failed" -Body $_.Exception.Message
    throw
}

# Step 5.6 Drop results in SharePoint Site
# Check if the ExchangeOnlineManagement module is installed and import it
$PnPmoduleName = "SharePointPnPPowerShellOnline"
if (-not (Get-Module -ListAvailable -Name $PnPmoduleName)) {
    Write-Log -Level WARN -Message "Attempting to install module: $moduleName"
    Install-Module $PnPmoduleName -Force -Scope CurrentUser
}
Import-Module $PnPmoduleName

try {
    # Connect to PnPOnline
    Connect-PnPOnline -Url $SPSite -ClientId $appID -Tenant $org -Thumbprint $thumbprint
    Write-Log -Level DEBUG -Message "Connected to $SPSite"
    try {
        # Upload to SP site folder
        Add-PnPFile -Path $resultsFile -Folder $SPFolder
        Write-Log -Level INFO -Message "Uploaded $resultsFile to $SPFolder"
    } catch {
        Write-Log -Level ERROR -Message "Error in uploading to PnPOnline: $_Exception.Message"
    }
} catch {
    Write-Log -Level ERROR -Message "Error in connecting to PnPOnline: $_.Exception.Message"
}


# Step 6: Identify and Email Unused DLs
# Process the auditResults.csv to find DLs not used in the last 90 days
try {
    $unusedDLs = Import-Csv $resultsFile | Where-Object { $_.EmailsSent -eq 0 }
    if ($unusedDLs) {
        $body = $unusedDLs | Format-Table | Out-String
        Write-Log -Level INFO -Message "Sending Email"
        Email-Me -Subject "Unused DLs in Last 90 Days" -Body $body
    }
} catch {
    Write-Log -Level ERROR -Message "Script Error: Identifying Unused DLs Failed $_.Exception.Message"
    Email-Me -Subject "Script Error: Identifying Unused DLs Failed" -Body $_.Exception.Message
    throw
}

# Sending a success email notification
Email-Me -Subject "Script Completed Successfully" -Body "The script has completed its execution successfully."
Write-Log -Leven INFO -Message "DL Audit Complete."

# Disconnect from Exchange Online session
Start-Sleep 30
try {
    Disconnect-ExchangeOnline -Confirm:$false
} catch {
    Write-Log -Level WARN -Message "Error in disconnecting Exchange Online: $_"
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
