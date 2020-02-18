Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 


$MySearchBase = "*"

# For our first step – we load up a variable with the groups we want (filtered by type):
$MyGroupList = get-adgroup -Filter 'GroupCategory -eq "Distribution" -and GroupScope -eq "Universal"' -SearchBase "$MySearchBase" -Properties Mail

# If you want to validate you got the correct groups in the variable, list out the names of your objects in the variable:
#$MyGroupList | Select Name, Mail

# Now, for every group in the list, we flip the type to Universal:
foreach($Group in $MyGroupList){
    Write-Host "Disabling Email Address Policy for group:" $Group.Name
    Set-DistributionGroup -Identity $Group.Name -EmailAddressPolicyEnabled $false
}
