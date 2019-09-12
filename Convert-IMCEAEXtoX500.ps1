# Source : https://gallery.technet.microsoft.com/office/IMCEAEX-to-X500-Converter-4632149a

$IMCEAEX= Read-Host "Please enter the IMCEAEX that needs to be converted"
$Clean = $IMCEAEX.replace("+20", " ").replace("+28", "(").Replace("+29", ")").replace("IMCEAEX-", "X500:").replace("_", "/").replace("+2E", ".").split("@")[0]

# -replace "+28", "(" -replace "+29", ")" -replace "IMCEAEX", "X500:" -replace "@"*,  "" -replace "_", "/"

Write-Host "The converted X500 is: $Clean" -ForegroundColor Yellow