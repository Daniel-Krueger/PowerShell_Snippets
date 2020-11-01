
function Assert-ImportExcelModule {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Write-Host "Module exists"
    } 
    else {
        Install-Module -Name ImportExcel  -scope CurrentUser -Force
    }
    Import-Module -Name ImportExcel -Force
}

<#
.SYNOPSIS
Merges the first worksheet of each file and saves it into a new one.
#

.DESCRIPTION
Takes the content of the first worksheet of each file and merges these into a new file.


.PARAMETER filePaths
An array of file path objets returned by Get-ChildItem. They should contain a FullName and Name property

.EXAMPLE
An example
$filePaths = get-ChildItem ".\Excel\MergeFiles\ExampleFiles" -filter "*.xlsx"
$outputFilePath = ".\Excel\MergeFiles\ResultFile\output.xlsx" 

Merge-Files -filePaths $filePaths -outputFilePath $outPutFilePath

.NOTES
General notes
#>
function Merge-Files {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [array]$filePaths,
        
        [Parameter(Mandatory = $true)]
        [string]
        $outputFilePath,

        [Parameter(Mandatory = $false)]
        [bool]
        $overwrite = $true
    )    
    begin {
        Assert-ImportExcelModule   
        if ((Test-Path $outputFilePath) -eq $true -and !$overwrite) {
            throw "The output file '$outputFilePath' already exist and should not be overwritten."
        }
        
    }
    
    process {
        for ( $i = 0; $i -lt $filePaths.Count; $i++) {
            $filePath = $filePaths[$i]
            Write-Host "Opening file '$($filePath.FullName)'"
            $existingWorksheet = Import-Excel -Path $filePath.FullName -NoHeader -Verbose
            Write-Host "Adding source file column '$($filePath.Name)'"
            foreach ($row in $existingWorksheet) {
                Add-Member -InputObject $row -Name "originial filename" -Value $filePath.Name  -MemberType NoteProperty -Force
            }

            if ($i -gt 0){
                Write-Host "Appending conent of file '$($filePath.FullName)' to '$outputFilePath'."
            
                $existingWorksheet | Export-Excel -Path $outputFilePath  -WorkSheetname "Merged" -Append
            }
            else {                
                Write-Host "Creating new output file '$outputFilePath' for file '$($filePath.FullName)'."
                $existingWorksheet | Export-Excel -Path $outputFilePath  -WorkSheetname "Merged" 
            }
        }
    }
    
    end {
        
    }
}