#=============================================================================================================
# AUTHOR: Riaz Ansary
#=============================================================================================================
Param ([Parameter(Mandatory = $true)]$NewDomain ,[Parameter(Mandatory = $true)]$OU)
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
Set-ADServerSettings -ViewEntireForest $true
$DLs = Get-Mailbox -OrganizationalUnit $OU
If (([bool](Get-AcceptedDomain -ErrorAction SilentlyContinue | Where-Object {$_.DomainName -eq $NewDomain})) -eq $false)
{
Write-Warning "Cannot find an Accepted Domain for: $NewDomain Adding it now ... Please Wait"
New-AcceptedDomain -DomainName $NewDomain -DomainType Authoritative -Name $NewDomain
Write-Warning "$newdomain is added as an accepted domain in our Exchange Infrastracture."
}
    Write-Host "Changing Primary SMTP Address of Users in: $OU" -ForegroundColor Green
    ForEach($i in $DLs)
    {
    $oldSmtp = $i.PrimarySmtpAddress
    $Alias = $oldSmtp.Local
    $OldDomain = $oldSmtp.Domain
    $NewSMTP = "$Alias@$NewDomain"
    Set-Mailbox -Identity "$i" -PrimarySmtpAddress "$NewSMTP" -EmailAddressPolicyEnabled $false
    }
    Write-Host "Completed. $NewDomain is now setup as PrimarySMTPAddress of Users in $OU" -ForegroundColor Green