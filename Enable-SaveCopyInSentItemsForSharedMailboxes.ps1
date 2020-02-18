Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

$sharedmailboxes = Get-Mailbox -Filter {recipienttypedetails -eq "SharedMailbox"} 

foreach($mailbox in $sharedmailboxes){
    Write-Host "Enabling Save Sent Item in Shared Mailbox folder for :" $mailbox
    Set-Mailbox $mailbox -MessageCopyForSentAsEnabled $True
}
