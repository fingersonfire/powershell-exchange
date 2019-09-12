#requires -version 4
<#
.SYNOPSIS
  Remove duplicated Exchange Server Rules

.DESCRIPTION
  WARNING : This script deletes rules from a mailbox without confirmation. Make a backup if you're not sure what you are doing.
  It checks for duplicated rules and keeps only the first occurence of each of them.
  Rule are considered duplicates if they have the same name.

.PARAMETER Mailbox
  The email address you want to delete duplicated rules from.

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
  Main Usage
  
  Remove-DuplicatedInboxRules.ps1 -Mailbox name@company.com
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
		[Parameter(Position=0,mandatory=$true)]
        [string] $Mailbox
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# List of all the rules for the user
$rules = @()

# List of all unique rules names
$uniquerulenames = @()

#-----------------------------------------------------------[Functions]------------------------------------------------------------



#-----------------------------------------------------------[Execution]------------------------------------------------------------


write-host "Retreiving Rules"
$rules = Get-InboxRule -Mailbox $Mailbox -WarningAction $ErrorActionPreference

foreach($rule in $rules)
{
	if($uniquerulenames.Contains($rule.Name)){
        Remove-InboxRule -Mailbox $Mailbox -Identity $rule.RuleIdentity -Confirm:$false
		write-host -ForegroundColor Red "DELETED:" $rule.Name
	}
	else{
		$uniquerulenames += $rule.Name
		write-host -ForegroundColor Green "KEPT:" $rule.Name
	}	
}