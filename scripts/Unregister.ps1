# Removes the VisioSizeTools COM add-in registration for the current user.
$clsid  = '{4A11EACB-6B73-4915-AD54-D92ECF41F340}'
$progId = 'VisioSizeTools.Connect'

foreach ($k in @(
    "HKCU:\Software\Microsoft\Visio\Addins\$progId",
    "HKCU:\Software\Classes\$progId",
    "HKCU:\Software\Classes\CLSID\$clsid")) {
    if (Test-Path $k) { Remove-Item $k -Recurse -Force; Write-Host "Removed $k" }
}
