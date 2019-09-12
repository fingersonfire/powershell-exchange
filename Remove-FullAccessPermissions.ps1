#requires -version 4
<#
.SYNOPSIS
  Remove full access permissions for a user on different mailboxes

.DESCRIPTION
  WARNING : This script remove full access permissions for a bunch of users without prompting for confirmation
  
  This script make sure all permissions are revoked for the user, even if the permission has been inherited.
  The trick is to add the permission again and then removing it. 

.PARAMETER None

.INPUTS
  None

.OUTPUTS
  None

.NOTES
  Version:        1.0
  Author:         FingersOnFire
  Creation Date:  2019-06-20
  Purpose/Change: Initial script development

.EXAMPLE
  Main Usage
  
  Remove-FullAccessPermissions.ps1
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------


#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# List of all the users you want to get permisssions removed from
$identities = @('user1', 'user2')

# The user that should not have access to the mailboxes listed in identities array anymore
$user = 'username'

#-----------------------------------------------------------[Functions]------------------------------------------------------------



#-----------------------------------------------------------[Execution]------------------------------------------------------------

foreach($identity in $identities) {
    write-host
    write-host Checking existance of $identity    
    if(Get-Mailbox $identity){
        
		write-host Re-adding permission on $identity
        $addresult = Add-MailboxPermission -Identity $identity -User $user -AccessRights FullAccess -InheritanceType All -Automapping $false
        
		write-host Removing permission on $identity
        Remove-MailboxPermission -Identity $identity -User $user -AccessRights FullAccess -InheritanceType All -Confirm:$false
    }
}
