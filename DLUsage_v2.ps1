# PowerShell Script for Managing Distribution Lists and Historical Searches

# Step 0: Check and Archive Existing Files
$archiveFolder = "~\DLUsageArchive"
$filesToCheck = @("~\DLs.txt", "~\DLID.csv", "~\auditResults.csv") + (Get-ChildItem "~\DLgroup*.txt")
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
    Install-Module $moduleName -Force
}
Import-Module $moduleName

# Attempt to connect to Exchange Online, handle any connection errors
try {
    Connect-ExchangeOnline -ErrorAction Stop
} catch {
    Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Script Error: Connection Failed" -Body $_.Exception.Message -SmtpServer "smtp.example.com"
    throw
}

# Step 2: Get a list of DLs and save them to a file
# Redirect any errors in getting the Distribution Group to an email notification
try {
    (Get-DistributionGroup -ErrorAction Stop).PrimarySmtpAddress > ~\DLs.txt
} catch {
    Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Script Error: Fetching DLs Failed" -Body $_.Exception.Message -SmtpServer "smtp.example.com"
    throw
}

# Step 3: Divide the DLs into files containing 250 entries
$numfiles = 0
try {
    $DLs = Get-Content -Path ~\DLs.txt
    $numfiles = [math]::Ceiling($DLs.length / 250)
    for ($i = 1; $i -le $numfiles; $i++) {
        $a = ($i - 1) * 250
        $b = [math]::Min($i * 250 - 1, $DLs.length - 1)
        $DLs[$a..$b] > ~\DLgroup$i.txt
    }
} catch {
    Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Script Error: Dividing DLs Failed" -Body $_.Exception.Message -SmtpServer "smtp.example.com"
    throw
}

# Step 4: Start Historical Search every day until done
New-Item ~\DLID.csv -Force
for ($j = 1; $j -le $numfiles; $j++) {
    try {
        Get-Content -Path ~\DLgroup$j.txt | %{
            Start-HistoricalSearch -ReportTitle "Day $j" -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) -ReportType MessageTrace -RecipientAddress $_ -ErrorAction Stop
            Start-Sleep -Milliseconds 500
        } | Export-CSV -Path ~\DLID.csv -Append
    } catch {
        Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Script Error: Historical Search Day $j Failed" -Body $_.Exception.Message -SmtpServer "smtp.example.com"
        throw
    }
    if ($j -lt $numfiles) {
        Start-Sleep -Seconds 86401
    }
}

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
        } | Export-CSV ~\auditResults.csv -Append -NoTypeInformation
} catch {
    Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Script Error: Exporting Results Failed" -Body $_.Exception.Message -SmtpServer "smtp.example.com"
    throw
}

# Step 6: Identify and Email Unused DLs
# Process the auditResults.csv to find DLs not used in the last 90 days
try {
    $unusedDLs = Import-Csv ~\auditResults.csv | Where-Object { $_.EmailsSent -eq 0 }
    if ($unusedDLs) {
        $body = $unusedDLs | Format-Table | Out-String
        Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Unused DLs in Last 90 Days" -Body $body -SmtpServer "smtp.example.com"
    }
} catch {
    Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Script Error: Identifying Unused DLs Failed" -Body $_.Exception.Message -SmtpServer "smtp.example.com"
    throw
}

# Sending a success email notification
Send-MailMessage -From "sender@example.com" -To "your-email@example.com" -Subject "Script Completed Successfully" -Body "The script has completed its execution successfully." -SmtpServer "smtp.example.com"
