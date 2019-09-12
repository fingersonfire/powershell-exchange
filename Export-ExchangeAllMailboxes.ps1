#requires -version 4
<#
.SYNOPSIS
  Export all mailboxes from a Microsoft Exchange server

.DESCRIPTION
  Retreive all mailboxes names and run an Export Mailbox request for each of them

.PARAMETER ExportPath
  The writable location where the PST files will be stored.

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:        1.0
  Author:         FingersOnFire
  Creation Date:  2019-05-02
  Purpose/Change: Initial script development

.EXAMPLE
  Simple Export
  
  Export-ExchangeAllMailboxes.ps1 -ExportPath \\server\sharedfolder
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param
(
	[parameter(Mandatory=$true)]
	[ValidateNotNull()]
	$ExportPath
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Any Global Declarations go here

#-----------------------------------------------------------[Functions]------------------------------------------------------------



#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Script Execution goes here

foreach ($Mailbox in (Get-Mailbox)) 
{ 
	New-MailboxExportRequest -Mailbox $Mailbox -FilePath "$ExportPath\$($Mailbox.Alias).pst"
}