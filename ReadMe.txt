O365 doesn't provide a simple way to audit all DLs to determine which are not being used. They will let us do a mail trace on a DL for the last 90 days to see if it was used. In essance, that's what this script does.

One limitation of O365 is you can only run 250 Historical Searches via PowerShell per day. This script accounts for that by dropping 250 DLs into separate files, then running one file per day.

The day after the last file is run, the historical search results will be compiled and placed into a nice CSV file located at ~\auditResults.csv. 0 EmailsSent or "Complete - No results found" means no one has sent anything to the DL over the last 90 days.

This is for use with Exchange Online Management. User will be prompted to login with MS365 credentials. The account used to log in must be assigned the proper role to send remote PowerShell commands.

WARNINGS:
This script will only run successfully if it is uninterrupted during the entire time it runs. It will process 250 DLs per day, then it will take 1 extra day to compile the results. In other words, be patient with this baby, and don't restart your computer for a few days;)

Also, if you've run any other historical searches in the last 24 hours, I'd wait 24 hours before running this script, otherwise you may miss some, and the script may pull up results from non-DLs.
