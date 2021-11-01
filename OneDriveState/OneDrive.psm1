<#
.SYNOPSIS
    Verify the synchronization state of files that should be synced to OneDrive

.DESCRIPTION
    
    Made use of these helpful modules
        https://rodneyviana.com/powershell-cmdlet-to-check-onedrive-for-business-or-onedrive-personal-status/
        https://github.com/dfinke/ImportExcel


.PARAMETER EntryFolder
    The entry folder from where the verification process should be started 
.PARAMETER ResultsFolder
    The folder in which the results.xlsx will be stored
.PARAMETER OneDriveLibFolder
    The folder in which the OneDriveLib resides. Needs only to be set, if the script is not executed from this folder.
.PARAMETER BackupFolderPath
    An optional path to a local backup storage.

.EXAMPLE
    cd C:\Workspace\_Privat\PowerShell_Snippets\OneDriveState
    Import-Module .\OneDrive.psm1 -force
    $EntryFolder = "$($env:OneDrive)\Test Folder"
    $ResultsFolder = "C:\Workspace\_Privat\PowerShell_Snippets\OneDriveState\Results"
    Get-UnsyncedFiles -EntryFolder $EntryFolder -ResultsFolder $ResultsFolder

.NOTES
General notes
#>
function Get-UnsyncedFiles {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$EntryFolder,
        
        [Parameter(Mandatory = $true)]
        [string]$ResultsFolder,
        
        [Parameter(Mandatory = $false)]
        [string]$OneDriveLibFolder = ".",

        [Parameter(Mandatory = $false)]
        [string]$BackupFolderPath = "f:\" 
    )    
    begin {        
        
        Assert-ImportOneDriveLib -OneDriveLibFolder $OneDriveLibFolder
        Assert-ImportExcelModule   
        if ((Test-Path $ResultsFolder) -ne $true ) {
            throw "The result folder '$ResultsFolder' does not exist or is not accessible."
        }
        $labels =@{
            "Header_RealtivePath"="Relative path"
            "Header_Status"="Status"
            "Header_FolderLink"="Folder link"
            "Label_FolderLink"="Folder"
            "Header_FileLink"="File link"
            "Label_FileLink"="File"
            "Header_LocalBackupFolder"="Backup link"            
            "Label_LocalBackupFolder"="Backup"
            "Label_LocalBackupFolderParam"="Local Backup root folder"
            
        }
        if ([System.Threading.Thread]::CurrentThread.CurrentUICulture.TwoLetterISOLanguageName -eq 'de') {
            $labels.Header_RealtivePath ="Realtiver Pfad"
            $labels.Header_Status ="Status"
            $labels.Header_FolderLink ="Ordner Link"
            $labels.Label_FolderLink="Ordner"
            $labels.Header_FileLink="Datei Link"
            $labels.Label_FileLink="Datei"
            $labels.Header_LocalBackupFolder="Sicherung Link"
            $labels.Label_LocalBackupFolder="Sicherung"
            $labels.Label_LocalBackupFolderParam="Lokales Backupverzeichnis"
            
        }
        $resultFileName = "$ResultsFolder\result_$((Get-Date).ToString("yyyy-MM-dd-HHmmss")).xlsx"
    }    
    process {
        # Force returns hidden as well as non hidden files.
        Write-host "Fetching files, this can take a while"
        
        $excel = "" | Export-Excel $resultFileName  -Append:$false -ClearSheet -FreezeTopRow -FreezeFirstColumn -AutoFilter -PassThru
        $params = $excel.Workbook.Worksheets.Add("Params")
        Set-ExcelRange -Range "A1" -Worksheet $params -Value $labels.Label_LocalBackupFolderParam 
        Set-ExcelRange -Range "B1" -Worksheet $params -Value $BackupFolderPath 
        
        $ws = $excel.Workbook.Worksheets["Sheet1"]
        Set-ExcelRange -Range "A1"  -Worksheet $ws -Value  $labels.Header_RealtivePath 
        Set-ExcelRange -Range "B1"  -Worksheet $ws -Value $labels.Header_Status 
        Set-ExcelRange -Range "C1"  -Worksheet $ws -Value $labels.Header_FolderLink 
        Set-ExcelRange -Range "D1"  -Worksheet $ws -Value $labels.Header_FileLink 
        Set-ExcelRange -Range "E1"  -Worksheet $ws -Value $labels.Header_LocalBackupFolder        
        
        $files = Get-ChildItem -Path $EntryFolder -force -File -Recurse
        $rowCounter = 1
        Write-Host "Going to process $($files.Count)"
        for($i = 0; $i -lt $files.Length ; $i++){
            $file = $files[$i]
            if ($i % 250 -eq 0) {
                $message = "This is updated every 250 files. Last file $i '$($file.Fullname)'"
                Write-Host $message
                Write-Progress -Activity "OneDrive State verification" -status "This is updated every 250 files. Last file $i/$($files.Count) '$($file.Fullname)'" -percentComplete ($i / $files.Length*100)
            }
            $state = Get-ODStatus –ByPath  $file.FullName
            if ($state -in @([OdSyncService.ServiceStatus]::ReadOnly,[OdSyncService.ServiceStatus]::Shared,[OdSyncService.ServiceStatus]::UpToDate))
            {
                continue
            }
            $rowCounter ++;
            $relativePath = $file.FullName.Substring($EntryFolder.Length)
            $relativeFolderPath = $file.DirectoryName.Substring($EntryFolder.Length)
            Set-ExcelRange -Range "A$rowCounter"  -Worksheet $ws -Value $relativePath 
            Set-ExcelRange -Range "B$rowCounter"  -Worksheet $ws -Value $state
            Set-ExcelRange -Range "C$rowCounter"  -Worksheet $ws -Formula "=hyperlink(`"$($file.DirectoryName)`",`"$($labels.Label_FolderLink)`")"
            Set-ExcelRange -Range "D$rowCounter"  -Worksheet $ws -Formula "=hyperlink(`"$($file.FullName)`",`"$($labels.Label_FileLink)`")"
            Set-ExcelRange -Range "E$rowCounter"  -Worksheet $ws -Formula "=hyperlink(Params!B1&`"$($relativeFolderPath)`",`"$($labels.Label_LocalBackupFolder)`")"
        }
        Set-ExcelRange -Range "A:E" -Worksheet $ws -AutoSize
        
        Close-ExcelPackage $excel -Show 
    }    
    end {
    }
}

function Assert-ImportExcelModule {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Write-Host "Module exists"
    } 
    else {
        Install-Module -Name ImportExcel  -scope CurrentUser -Force
    }
    Import-Module -Name ImportExcel -Force
}

function Assert-ImportOneDriveLib ($OneDriveLibFolder){
    $IsAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match “S-1-5-32-544”)
    $IsInteractive = [Environment]::UserInteractive

    if($IsAdmin -or (!$IsInteractive))
    {
        throw “You need to execute this in an interactive mode without activated admin mode. Running as admin '$($IsAdmin)', running interactive '$($IsInteractive)'”
    } 
    $onedriveLibPath =  Resolve-path "$OneDriveLibFolder\OneDriveLib.dll"
    if (!(Test-Path -Path "$OneDriveLibFolder\OneDriveLib.dll")) {
        throw "The OneDriveLib.dll was not found in folder '$OneDriveLibFolder'" 
    }
    Unblock-File $onedriveLibPath 
    Import-Module $onedriveLibPath -Force
            
    if (("OdSyncService.ServiceStatus" -as [type]) -eq $null) {
        [void][System.Reflection.Assembly]::LoadFile($onedriveLibPath);
    }
}
