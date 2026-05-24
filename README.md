# Add-M365-Aliase
PowerShell script for adding secondary SMTP aliases to Exchange Online mailboxes in Microsoft 365.

********************************************************************************

Features:
Single mailbox mode
Bulk mode using:
CSV
TSV
Excel
Auto-generates aliases from existing mailbox names
Reads accepted domains from Exchange Online
Auto-detects email columns
Supports -WhatIf
Verifies aliases after update

********************************************************************************

Requirements:
PowerShell
Exchange Online permissions
ExchangeOnlineManagement module

Install module:
Install-Module ExchangeOnlineManagement -Scope CurrentUser

********************************************************************************

Parameter:
-UserEmail	Single mailbox
-TargetDomain	Alias domain
-InputFilePath	CSV/Excel/TSV file
-WorksheetName	Excel worksheet
-EmailColumn	Email column
-WhatIf	Simulation mode

********************************************************************************

Single Mailbox:
powershell -ExecutionPolicy Bypass -File ".\Add-M365MailboxAlias.ps1" `
-UserEmail "test@example.com"

Single Mailbox + Target Domain:
powershell -ExecutionPolicy Bypass -File ".\Add-M365MailboxAlias.ps1" `
-UserEmail "test@example.com" `
-TargetDomain "alias.com"

Bulk CSV:
powershell -ExecutionPolicy Bypass -File ".\Add-M365MailboxAlias.ps1" `
-InputFilePath "C:\Temp\mailboxes.csv"

Bulk Excel:
powershell -ExecutionPolicy Bypass -File ".\Add-M365MailboxAlias.ps1" `
-InputFilePath "C:\Temp\mailboxes.xlsx" `
-WorksheetName "Sheet1" `
-EmailColumn "Mailbox" `
-TargetDomain "alias.com"

WhatIf Mode:
powershell -ExecutionPolicy Bypass -File ".\Add-M365MailboxAlias.ps1" `
-InputFilePath "C:\Temp\mailboxes.csv" `
-TargetDomain "aliase.com" `
-WhatIf


********************************************************************************
Result Statuses:
Status:	        Meaning
Added:	        Alias added successfully
AlreadyExists:	Alias already exists
NotVerified:	  Alias not confirmed after update
WhatIf:	        Simulation only
Error:	        Processing failed


Verify Mailbox:
Get-Mailbox user@domain.com | Format-List PrimarySmtpAddress,EmailAddresses

********************************************************************************

Notes:
Does NOT change primary SMTP address
Does NOT remove aliases
Target domain must already exist in Exchange Online


