Import-Module  -Name ".\OneDrive.psm1" -Force
$entryFolder = "$($env:OneDrive)"
$ResultsFolder = "C:\Workspace\_Privat\PowerShell_Snippets\OneDriveState\Results"
Get-UnsyncedFiles -entryFolder $entryFolder -ResultsFolder $ResultsFolder
