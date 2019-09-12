$FailedRequests = Get-MailboxImportRequest -Status Failed

for($FailedRequest in $FailedRequests){
    Get-MailboxImportRequestStatistics $FailedRequest -IncludeReport
}
