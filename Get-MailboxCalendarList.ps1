Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

$MailboxAlias = Get-Mailbox | Select-Object alias 

foreach ($Mailbox in $MailboxAlias) { 
    Get-MailboxFolderStatistics $Mailbox.alias | Where-Object{$_.FolderType -eq "Calendar"} | Select Identity 
} 