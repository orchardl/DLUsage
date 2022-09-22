# I'm a simple guy who like simple things to get past this 250 issue.

# step 1: Connect exchange online
install-module ExchangeOnlineManagement 
import-module ExchangeOnlineManagement
Connect-ExchangeOnline

# step 2: Get a list of DLs and drop them into a file
(Get-DistributionGroup).PrimarySmtpAddress > ~\DLs.txt

# step 3: Divide the DLs into files containing 250 entries
$numfiles = 0
for ($i = 0; $i -lt (Get-Content -Path ~\DLs.txt).length/250; $i++) {
	$a = $i*250
	$b = ($i+1)*250
	(Get-Content -Path ~\DLs.txt)[$a..$b] > ~\DLgroup$i.txt
	$numfiles = $i
}

# step 4: Start Historical Search every day until done
New-Item ~\DLID.csv
for ($j=0; $j -lt $numfiles; $j++) {
	Get-Content -Path ~\DLgroup$j.txt | %{
		Start-HistoricalSearch -ReportTitle "Day $j" -StartDate (Get-Date).AddDays(-90) -EndDate (Get-Date) -ReportType MessageTrace -RecipientAddress $_
		Start-Sleep -Milliseconds 500
	} | Export-CSV -Path ~\DLID.csv -Append
	Start-Sleep -Seconds 86401
}

#step 5: Get Historical Search on everything we just ran
Get-HistoricalSearch | 
	ForEach-Object -Process {
		if ($_.SubmitDate -gt (Get-Date).AddDays(-4)) {
			New-Object psobject -Property @{
				'DLname'=(Get-HistoricalSearch -JobID $_).RecipientAddress
				'ReportStatus'=(Get-HistoricalSearch -JobID $_).ReportStatusDescription
				'EmailsSent'=(Get-HistoricalSearch -JobID $_).Rows
			}
		}
	} | 
	Export-CSV ~\auditResults.csv -Append -NoTypeInformation
