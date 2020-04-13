# Exchange 2013/16 Journal Mailbox Archive-to-PST PowerShell Script with Reporting

Recent Exchange versions have built-in support of journaling for recording all inbound and outbound email messages for backup or compliance reasons. Overtime, the journal mailbox grows so large and needs to be trimmed or pruned.

This article documents a PowerShell maintenance script written for reporting and automating the monthly archive-to-PST process of the Exchange 2013 journaling mailbox. (Based on reader feedback, it also suits Exchange 2016 Standard.)

# Archiving Concept

This script uses the PowerShell cmdlet _New-MailboxExportRequest -Mailbox &lt;journal mailbox&gt;_ to export Exchange Journaling Mailbox of previous month (e.g. 2016-01-01 to 2016-01-31) as a standard PST file (e.g. archive 2016\_01\_31.pst) to specified locations (up to two locations) and then uses _Search-Mailbox -DeleteContent_ to delete email messages within the date range if successful. It is designed to be run at the beginning of each month (e.g. 2/Feb/16) using Windows Task Scheduler.

## Email Alerting, Reporting and Logging

![Reporting Email](https://farm5.staticflickr.com/4651/25509672137_c4b4acf7ac_o.png)

A log file is created under a specified directory during each execution. An email is sent to specified email addresses with the log file when done. The email message subject indicates COMPLETE/FAILURE.

In addition, a list of email messages returned from search and deleted by Search-Mailbox cmdlet will be available as zip(csv) attachments in another specified mailbox and folder as mentioned in the message body of above email.

## Safety Measure to Prevent Accidental Email Deletion

In case mail archiving has failed, script will send an alert mail and exit so that no mail deletion will occur. In technical terms, when status of _Get-MailboxExportRequest_ is not &quot;Completed&quot;, script keeps on waiting (looping) until it is complete. If the loop is broken or script receives any status other than &quot;Completed&quot;, execution will be terminated and failure will be reported by email.

# Assumptions and Requirements

- PowerShell 3.0 or above
- Execute this script on Exchange server where (preferably):
  - System Locale is English (United States)
  - Short date format (for en-US) is MM/dd/yyyy
- Sufficient disk space for storing PST files in the specified location(s)

Note: This script was designed and tested for email messages of previous month, and has not been tested in other configurations but could still work with some tweaking. (In fact, it does not only work for journal mailbox.)

# Getting Started

## Script Usage and Instructions

1. Edit variable parameters as required ( **Refer to the script comment sections 1-4** )
2. Save the script in a local folder e.g. C:\scripts\ws\_email\_archiving\_script.ps1
3. To export email of last month, the script should be scheduled to run any day within the current month.

For example, to export Jan 2016 email, running the script any day within Feb 2016 would do; the beginning of month is recommended (e.g. 2/Feb/16) .

## Setting up Windows Task Scheduler

Program/script:

- C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

Add arguments:

- -PSConsoleFile &quot;C:\Program Files\Microsoft\Exchange Server\V15\Bin\exshell.psc1&quot; -command &quot;. &#39;C:\Program Files\Microsoft\Exchange Server\V15\Bin\Exchange.ps1&#39;; &amp;&#39;C:\scripts\ws\_email\_archiving\_script.ps1&#39;&quot;

Start in:

- C:\Scripts

Alternatively, start the .ps1 script using a .bat script:

- &quot;C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe&quot; -PSConsoleFile &quot;C:\Program Files\Microsoft\Exchange Server\V15\Bin\exshell.psc1&quot; -command &quot;. &#39;C:\Program Files\Microsoft\Exchange Server\V15\Bin\Exchange.ps1&#39;; &amp;&#39;C:\scripts\ws\_email\_archiving\_script.ps1&#39;&quot;

## Release History

| Ver | Date | Changes |
| --- | --- | --- |
| 1.2 | 20171022 | - Turn journal mailbox name into a variable<br>- Further ensure New/Get-MailboxExportRequest job name uniqueness by naming it with current date time (in yyyyMMdd_hhmmsstt)<br>- Remove unused variables, scriptDir, startDateShort and endDateShort<br>- Improve script comments |
| 1.1 | 20160712 | First public release |
| 1.0 | 20160314 | First internal release |
