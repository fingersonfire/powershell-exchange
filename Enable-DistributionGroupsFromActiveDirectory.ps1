Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn; 

$MySearchBase = "*" # Set the OU you want to search in

# Loading Distribution groups
$MyGroupList = get-adgroup -Filter 'GroupCategory -eq "Distribution" -and GroupScope -eq "Universal"' -SearchBase "$MySearchBase" -Properties Mail


foreach($Group in $MyGroupList){
    Write-Host "Enabling group:" $Group.Name
    Enable-DistributionGroup $Group.Name
}