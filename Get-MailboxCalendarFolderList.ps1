$MailboxAlias = Get-Mailbox | Select-Object alias 

foreach ($Mailbox in $MailboxAlias) { 
    #Get-MailboxFolderStatistics -Identity $Mailbox.alias | select-object Identity, ItemsInFolder, FolderSize, FolderAndSubfolderSize 
    Get-MailboxFolderStatistics $Mailbox.alias | Where-Object{$_.FolderType -eq "Calendar"} | Select Identity 
} 