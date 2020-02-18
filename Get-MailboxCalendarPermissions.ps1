Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

$mailboxes = get-mailbox -Resultsize Unlimited

$permreport = @()

foreach ($mailbox in $mailboxes) {
    $calendarfolders = Get-MailboxFolderStatistics $mailbox.alias | Where-Object{$_.FolderType -eq "Calendar"}
    foreach ( $calendar in $calendarfolders ) {

        $identity = $calendar.Identity -Replace ("\\",":\") # Won't work if the calendar is not in the root folder.
        $permissions = get-mailboxfolderpermission $identity

        foreach ( $perm in $permissions ) {


            $permreport += [PSCustomObject]@{
                Mailbox = $mailbox
                GrantedUser = $perm.User
                Identity = $perm.Identity
                Access = [string]::join(', ', $perm.AccessRights)
                Valid = $perm.IsValid
            } 
        }
    }
}

$permreport | Export-Csv "C:\Temp\CalendarReport.csv"