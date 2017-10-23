# Journal Mailbox Archiving PowerShell Script for Exchange 2013
# Written by wandersick (https://wandersick.blogspot.com/2016/07/powershell-automation-of-journal.html)
# Version: 1.1.20160712

# ----------------------------------------
#       1. DEFINING MAIN VARIABLES
# ----------------------------------------

# Get last day of last month, start day of this month (not converted to string yet)

# [Output Examples]
# endDate = Monday, February 1, 2016 12:00:00 AM
# startDate = Friday, Jaunuary 1, 2016 12:00:00 AM

$endDate = Get-Date -Day 1 "00:00:00"
$startDate = $endDate.AddMonths(-1)

# Get last day of last month in "yyyy_MM_dd" string and set file path from it

# [Output Examples]
# archiveMonth = 2016_01_31
# archiveFile = archive 2016_01_31.pst
# archiveFileDir = \\localhost\e$\Archive_PST\
# archiveFilePath = \\localhost\e$\Archive_PST\archive 2016_01_31.pst

$archiveMonth = (Get-Date).AddDays(- (Get-Date).Day).ToString('yyyy_MM_dd')
$archiveFile = "archive " + $archiveMonth + ".pst"
$archiveFileDir = "\\localhost\e$\Archive_PST\"
$archiveFilePath = $archiveFileDir + $archivefile

# Make a secondary copy at the end

# [Output Examples]
# archiveEnable2ndCopy = $true
# archiveFileDir2ndCopy = "\\172.16.123.124\Archive_PST\Secondary_Backup\"
# archiveFilePath2ndCopy = "\\172.16.123.124\Archive_PST\Secondary_Backup\archive 2016_01_31.pst"

$archiveEnable2ndCopy = $true
$archiveFileDir2ndCopy = "\\172.16.123.124\Archive_PST\Secondary_Backup\"
$archiveFilePath2ndCopy = $archiveFileDir2ndCopy + $archivefile

# Define the archive job string using the archive month
# Give the archive job a unique name (each time the script runs) in order for Get-MailboxExportRequest to identify the correct job as Complete.

# [Output Examples]
# archiveJobName = messagejournal 2016_01_31

$archiveJobName = "messagejournal " + $archiveMonth

# Convert dates to short date format

# [Output Examples]
# endDateShort = 02/01/2016
# startDateShort = 01/01/2016

$endDateShort = $endDate.ToShortDateString()
$startDateShort = $startDate.ToShortDateString()

# Define search date string (of last month) for email message deletion after archiving

# [Output Examples]
# searchStartDate = 01/01/2016
# searchEndDate = 01/31/2016
# searchDateRange = Received:01/01/2016..31/01/2016

$searchStartDate = $startDate.ToString('MM/dd/yyyy')
$searchEndDate = (Get-Date).AddDays(- (Get-Date).Day).ToString('MM/dd/yyyy')
$searchDateRange = "Received:" + $searchStartDate + ".." + $searchEndDate

# ----------------------------------------
#        2. DEFINING LOG VARIABLES
# ----------------------------------------

# Log folder (Set log folder as current script folder)
$scriptDir = $PSScriptRoot

# Log date & time (Set current date and time for use as log file name)
$dateTime = Get-Date -format "yyyyMMdd_hhmmsstt"

# Log file (stdout, stderr)
$logFilePath = "C:\scripts\Log\Log_Email_Archive_$($archiveMonth)_$($dateTime).log"

# ----------------------------------------
#    3. DEFINING SEARCH LOG VARIABLES
# ----------------------------------------

# TargetFolder and TargetMailbox to save a log file of mail deleted under a specified folder within a specified mailbox
# Refer to example one of this Technet article: https://technet.microsoft.com/en-us/library/dd298173(v=exchg.160).aspx

$saveSearchLogFolder = "exchMailFolderName"
$saveSearchLogMailbox = "exchMailboxName"

# ----------------------------------------
#      4. DEFINING EMAIL VARIABLES
# ----------------------------------------

$emailSender = "Sender <sender@wandersick.com>"
$emailRecipient = "Recipient A <recipienta@wandersick.com>", "Recipient B <recipientb@wandersick.com>"
$emailCc = "Recipient C <recipientc@wandersick.com>"
$emailBcc = "Recipient D <recipientd@wandersick.com>"
$emailSubject = "Monthly Email Archive COMPLETE - " + $startDate.ToString('yyyy/MM') + " - Log Attached"
$emailBody = "This is an automated message after email archiving scheduled task has completed. Please check attached log file as well as $saveSearchLogFolder (folder) within $saveSearchLogMailbox (mailbox) for any problems."
$emailSubjectFailure = "Monthly Email Archive FAILED - " + $startDate.ToString('yyyy/MM') + " - Log Attached"
$emailBodyFailure = "This is an automated message after email archiving scheduled task has failed. Please check attached log file as well as $saveSearchLogFolder (folder) within $saveSearchLogMailbox (mailbox) for any problems."
$emailServer = "172.16.123.123"
$emailAttachment = $logFilePath

# ----------------------------------------
#     RECORDING VARILABLES TO LOG FILE
# ----------------------------------------

Write-Output "" | tee $logFilePath -Append
Write-Output "[Variables]" | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$dateTime = ' $dateTime | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$endDate = ' $endDate | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$startDate = ' $startDate | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveMonth = ' $archiveMonth | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFile = ' $archiveFile | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFileDir = ' $archiveFileDir | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFilePath = ' $archiveFilePath | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveEnable2ndCopy = ' $archiveEnable2ndCopy | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFileDir2ndCopy = ' $archiveFileDir2ndCopy | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveFilePath2ndCopy = ' $archiveFilePath2ndCopy | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$archiveJobName = ' $archiveJobName | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$endDateShort = ' $endDateShort | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$startDateShort = ' $startDateShort | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$searchStartDate = ' $searchStartDate | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$searchEndDate = ' $searchEndDate | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$searchDateRange = ' $searchDateRange | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailSender = ' $emailSender | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailRecipient = ' $emailRecipient | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailCc = ' $emailCc | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailSubject = ' $emailSubject | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailBody = ' $emailBody | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailSubjectFailure = ' $emailSubjectFailure | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailBodyFailure = ' $emailBodyFailure | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailServer = ' $emailServer | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$emailAttachment = ' $emailAttachment | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$scriptDir = ' $scriptDir | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append
Write-Output '$logFilePath = ' $logFilePath | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append

# ----------------------------------------
#        EXECUTING MAIN PROCEDURE
# ----------------------------------------

# Create an archiving job (an export request of journal mailbox within last month, with a job name that is uniquely identified, to specified location)

Write-Output "" | tee $logFilePath -Append
Write-Output "Exporting items between $startDate and $endDate" | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append

# Ensure no job of the same name exists.
Get-MailboxExportRequest -Name $archiveJobName | Remove-MailboxExportRequest -confirm:$false
New-MailboxExportRequest -ContentFilter "(Received -ge '$startDate') -and (Received -lt '$endDate')" -Mailbox "messagejournal" -FilePath $archiveFilePath -Name $archiveJobName | tee $logFilePath -Append

# Wait for the archiving job to complete

while (!(Get-MailboxExportRequest -Name $archiveJobName -Status Completed))
{
	Write-Output "" | tee $logFilePath -Append
	Write-Output "Still exporting... Waiting 30 minutes for it to complete..." | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	Start-Sleep -s 1800
}
# To be safe before deletion, double-check the Complete status of email archiving
$emailArchivingStatus = Get-MailboxExportRequest -Name $archiveJobName -Status Completed | select -expand status
if ($emailArchivingStatus -eq "Completed")
{
	Write-Output "" | tee $logFilePath -Append
	Write-Output "Export completed." | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	# Getting more details for manual verification/troubleshooting in log file
	Get-MailboxExportRequest -Name $archiveJobName | fl | tee $logFilePath -Append
	Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl | tee $logFilePath -Append
	# Optional: Clean up (but Get-MailboxExportRequest will no longer return information of the previous job for troubleshooting; instead, check log file for troubleshooting)
	# Get-MailboxExportRequest -Name $archiveJobName | Remove-MailboxExportRequest -confirm:$false
}
else
{
	# Report failure via email and exit script
	Write-Output "" | tee $logFilePath -Append
	Write-Output "Export failed." | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	# Getting more details for manual verification/troubleshooting in log file
	Get-MailboxExportRequest -Name $archiveJobName | fl | tee $logFilePath -Append
	Get-MailboxExportRequest -Name $archiveJobName | Get-MailboxExportRequestStatistics | fl | tee $logFilePath -Append
	# Optional: Clean up (but Get-MailboxExportRequest will no longer return information of the previous job for troubleshooting; instead, check log file for troubleshooting)
	# Get-MailboxExportRequest -Name $archiveJobName | Remove-MailboxExportRequest -confirm:$false
	Send-MailMessage -From $emailSender -To $emailRecipient -Cc $emailCc -Subject $emailSubjectFailure -Body $emailBodyFailure -SmtpServer $emailServer -Attachments $emailAttachment
	Exit
}

# Creating a second copy to specified location

if ($archiveEnable2ndCopy) { Copy-Item -Path $archiveFilePath -Destination $archiveFilePath2ndCopy -Force -Confirm:$false | tee $logFilePath -Append }

# In case there is any error during copying, send email and exit script so that no deletion will occur.

$fileHashSrc = Get-FileHash $archiveFilePath | select -expand hash
$fileHashDst = Get-FileHash $archiveFilePath2ndCopy | select -expand hash

if ($fileHashSrc -eq $fileHashDst)
{
	Write-Output "" | tee $logFilePath -Append
	Write-Output "Second file copy completed and verified successfully." | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
}
else
{
	# Report failure via email and exit script
	Write-Output "" | tee $logFilePath -Append
	Write-Output "Second file copy verification failed." | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	Send-MailMessage -From $emailSender -To $emailRecipient -Cc $emailCc -Subject $emailSubjectFailure -Body $emailBodyFailure -SmtpServer $emailServer -Attachments $emailAttachment
	Exit
}

# Creating a log file of email messages to be deleted, then delete email (after PST export task is complete)
# This has to be run multiple times in order to get rid of all (>10000) email messages as 10000 is the limit of a single Search-Mailbox query.
# Except the first run, any sebsequent run of the mail deletion command only happens when remaining item count is 10000.

do
{
	Write-Output "" | tee $logFilePath -Append
	Write-Output "Logging a list of items in journal mailbox to be deleted into $saveSearchLogFolder (Folder) of $saveSearchLogMailbox (Mailbox)" | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	# Creating a log file of email messages to be deleted
	$remainingItemCount = Get-Mailbox -Identity "messagejournal" | Search-Mailbox -SearchQuery $searchDateRange -LogOnly -LogLevel Full -TargetFolder $saveSearchLogFolder -TargetMailbox $saveSearchLogMailbox | select -expand ResultItemsCount
	Write-Output '$remainingItemCount = ' $remainingItemCount | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	Write-Output "Deleting archived messages from journal mailbox with specified date range: $searchDateRange" | tee $logFilePath -Append
	Write-Output "" | tee $logFilePath -Append
	# Delete email
	Get-Mailbox -Identity "messagejournal" | Search-Mailbox -SearchQuery $searchDateRange -DeleteContent -force | tee $logFilePath -Append
	# Sleep for 5 minutes only when there are more email messages to delete (may not be required)
	if ($remainingItemCount -eq 10000)
	{
		Write-Output "" | tee $logFilePath -Append
		Write-Output "Sleep for 5 minutes" | tee $logFilePath -Append
		Write-Output "" | tee $logFilePath -Append
		start-sleep -s 300
	}
}
while ($remainingItemCount -eq 10000)

Write-Output "" | tee $logFilePath -Append
Write-Output "Sending an email with log to $emailRecipient cc $emailCc bcc $emailBcc " | tee $logFilePath -Append
Write-Output "" | tee $logFilePath -Append

# Send an email using local SMTP server with log when done.
Send-MailMessage -From $emailSender -To $emailRecipient -Cc $emailCc -Bcc $emailBcc -Subject $emailSubject -Body $emailBody -SmtpServer $emailServer -Attachments $emailAttachment