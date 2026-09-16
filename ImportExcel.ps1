# 1. Télécharger via le navigateur (qui passe le proxy) :
#    https://www.powershellgallery.com/packages/ImportExcel
#    → cliquez "Manual Download" → sauvegardez le .nupkg

# 2. Renommer et extraire :
$nupkg = "$env:USERPROFILE\Downloads\ImportExcel.7.8.10.nupkg"  # adaptez le nom
$dest  = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\ImportExcel"
New-Item -ItemType Directory -Path $dest -Force
Rename-Item $nupkg "$nupkg.zip"
Expand-Archive "$nupkg.zip" -DestinationPath $dest -Force

# 3. Test :
Import-Module ImportExcel -ErrorAction Stop
Get-Command Export-Excel