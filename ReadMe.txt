O365 doesn't provide a simple way to audit all DLs to determine which are not being used. They will let us do a mail trace on a DL for the last 90 days to see if it was used. In essance, that's what this script does.

One limitation of O365 is you can only run 250 Historical Searches via PowerShell per day. This script accounts for that by dropping 250 DLs into separate files, then running one file per day.

The day after the last file is run, the historical search results will be compiled and placed into a nice CSV file located at ~\auditResults.csv. 0 EmailsSent or "Complete - No results found" means no one has sent anything to the DL over the last 90 days.

This is for use with Exchange Online Management. User will be prompted to login with MS365 credentials. The account used to log in must be assigned the proper role to send remote PowerShell commands.

WARNINGS:
This script will only run successfully if it is uninterrupted during the entire time it runs. It will process 250 DLs per day, then it will take an extra few hours on top of that to compile the results. In other words, be patient with this baby, and don't restart your computer for a few days;)

Also, if you've run any other historical searches in the last 24 hours, I'd wait 24 hours before running this script, otherwise you may miss some, and the script may pull up results from non-DLs.


<h1>UPDATED ReadMe:</h1>

DLUsageWeekly.ps1
This is made to be on a schedule and totally automated from the background. Schedule it to run once per week. You'll need to create the app/api, certificate, and assign all necessary permissions beforehand.

DLUsage_v2.ps1
This does basically the same thing as DLusage.ps1, but it's updated with better error handling.

DLusage.ps1
This is the most simple version of the script.

Weekly_DL_Report_Wrapper.ps1
This is just a wrapper for DLUsageWeekly.ps1 that I created to assist with the automation in case something bigger breaks.

^sorry I wrote those descriptions when I didn't have much sleep. It's good enough for now.
