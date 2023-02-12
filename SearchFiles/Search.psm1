$ErrorActionPreference = "Inquire"
<#
.SYNOPSIS
	Searches files for the specified string and copies the results to clipboard
.DESCRIPTION
    The function will get all directories which should be searched.
    If a directory matches the name of an excluded folder, this folder and any sub folders will be excluded.
    The result is saved to $global:foldersToSearch. 
    If you execute the search for a second time in the same session without changing the parameters $basefolder and $ExcludedFolders, this folder list will be reused on subsequent calls.

    Afterwards the files are identified, which should be searched. The result will be saved to $global:filesToSearch.
    If none of the following parameters changed, than these files will be searched again on a subsequent call.

    Upon completion the search, the result is copied to the clipboard. In addition it's stored to $global:hits so that you can access it.
    You can also call 'Save-HitsToDesktop' to save the hits to a file which will be placed on the the desktop.   


.NOTES
    Author:   Daniel Krüger
	Date:     2023-02-10
.PARAMETER BaseFolder
    The folder from which the search should start
.PARAMETER SearchString
    The string which should be searched
.PARAMETER ExcludedFolders
    The folder names which should be excluded in form of an array.
    The default value is: @(".devops",".git",".vs",".vscode","packages","node_modules","debug","bin","build,"dist","Ruby","Ruby27-x64") 
.PARAMETER IncludedFileTypes
    The file types which should be searched in form of an array.
    The default value is all file types / an empty: @() 
.PARAMETER ExcludedFileTypes
    The file types which should be excluded from the search in form of an array.
    The default value is: @(".dll") 
.PARAMETER MaxFileSizeInKb
    The max size the files in kb. Larger files will be ignored.
    The default value is: 100 
.PARAMETER target
    Determines whether the results will either intialy be stored in a text file on the desktop or copied to the clipboard
    The default value is: Desktop
.EXAMPLE     
    Import-Module C:\Workspace\_Privat\private_github\PowerShell_Snippets\SearchFiles\search.psm1 -force
    $baseFolder = "C:\Workspace\_Privat\private_github" 
    $searchString ="prototype" 
    
    $ExcludedFolders = @(".devops",".git",".vs",".vscode","packages","node_modules","debug","bin","obj","dist","build","Ruby","Ruby27-x64")    
    Search-FilesWithExcludedFolders -baseFolder $baseFolder -searchString $searchString 
.EXAMPLE     
    Import-Module C:\Workspace\_Privat\private_github\PowerShell_Snippets\SearchFiles\search.psm1 -force
    $baseFolder = "C:\Workspace\_Privat\private_github\PowerShell_Snippets\SearchFiles\Test" 
    $searchString ="prototype"     
    Search-FilesWithExcludedFolders -baseFolder $baseFolder -searchString $searchString -IncludedFileTypes @() -MaxFileSizeInKb 10 -ExcludedFileTypes @()
#> 


function Search-FilesWithExcludedFolders {

  [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [string]$BaseFolder,
        [Parameter(Mandatory = $true, ValueFromPipeline)]
        [string]$SearchString,
        [Parameter(Mandatory = $false)]
        [array]$ExcludedFolders = @(".devops",".git",".vs",".vscode","packages","node_modules","debug","bin","dist","Ruby","Ruby27-x64"),      
        [Parameter(Mandatory = $false)]
        [array]$IncludedFileTypes = @(),
        [Parameter(Mandatory = $false)]
        [array]$ExcludedFileTypes = @(".dll",".png" ),
        [Parameter(Mandatory = $false)]
        [int]$MaxFileSizeInKb= 10,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Clipboard","Desktop")]    
        $Target = "Desktop"

    )
    begin{
        $updateFolders =  $global:lastFolder -ne $baseFolder -and $global:lastExcludedFolders -ne $ExcludedFolders
        
        if ($updateFolders -eq $false) {
            $updateFiles = ($global:lastIncludedFileTypes -ne $IncludedFileTypes) -or ($global:lastMaxFileSizeInKb -ne $MaxFileSizeInKb) -or ($global:lastExcludedFileTypes -ne $ExcludedFileTypes)
        } else{
            $updateFiles = $true
        }
        if ($updateFolders) {
            Write-Host "Folders to search need to be refreshed because the conditions changed."
        } else {
            Write-Host "Reusing folders, conditions did not change."
        }
        if ($updateFiles) {
            Write-Host "Files to search need to be refreshed because either the folders changed or the conditions."        
        } else {
            Write-Host "Reusing files, conditions did not change."
        }
    }
    process{
        Write-Progress -Activity "Searching folder '$baseFolder' for string '$searchString'" -Status "Getting all folders without '$ExcludedFolders'" -CurrentOperation "this may take a while"
        if ($updateFolders) {
            $result = $null
            $global:foldersToSearch  = New-Object System.Collections.ArrayList
            [void]$global:foldersToSearch.Add($baseFolder)        
            $result = [array](Get-FoldersWithoutExcluded -baseFolder $baseFolder -ExcludedFolders $ExcludedFolders)
            if ($result.Count -gt 0) { 
                $global:foldersToSearch.AddRange($result)
            }
        } 

        if ($updateFiles) {
            $total = $global:foldersToSearch.Count
            $progress = 0;
            Write-Progress -Activity "Searching folder '$baseFolder' for string '$searchString'" -Status "Getting all files in the relevant folders" -CurrentOperation "Processed folders" -PercentComplete 0
            $global:filesToSearch = new-object System.Collections.ArrayList;      
        
            foreach ($folder in $global:foldersToSearch) {
                <# For debugging purposes 
                $folder = $global:foldersToSearch[0]
                #> 
                $progress++;
                if ($progress % 50 -eq 0) {
                    Write-Progress -Activity "Searching folder '$baseFolder' for string '$searchString'" -Status "Getting all files in the relevant folders" -CurrentOperation "Current folder $($folder)" -PercentComplete ($progress*100/$total)
                }

                # see https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_splatting?view=powershell-7.3
               
                $allFilesInFolder = [array](Get-ChildItem -Path $folder -File)

                if ($allFilesInFolder.Count -gt 0) {
                
                    $filesToSearch =  [array]($allFilesInFolder | where {
                                
                                (($_.Extension -in $ExcludedFileTypes) -eq $false) -and 
                                (($IncludedFileTypes.Count -eq 0) -or ($_.Extension -in $IncludedFileTypes)) -and
                                ($_.Length -lt $MaxFileSizeInKb*1024 )
                                
                    })
                    #$filesToSearch =  $allFilesInFolder | where { ($_.Length -lt $MaxFileSizeInKb*1024 )}                
                    if ($filesToSearch.Count -gt 0) {
                        $global:filesToSearch.AddRange($filesToSearch)
                    }
                }
            }    
        }  
        
        
        if ($global:filesToSearch.Count -eq 0) {
            Write-Host "No files where found, witch matched the conditions." 
            return
        }

        # Splitting $filesToSearch into chunks, so that the progress bar can be used.
        $chunkSize = 100
        if ($chunkSize -gt $global:filesToSearch.Count) {
            $fileChunks = New-Object System.Collections.ArrayList  
            [void]$fileChunks.Add($global:filesToSearch)
        }
        else {
            $fileChunks = $global:filesToSearch | Select-Chunk -ReadCount $chunkSize
        }

        $total = $global:filesToSearch.Count
        $progress = 0;       
        $Global:hits = New-Object System.Collections.ArrayList          
        foreach ($chunk in $fileChunks) {
            <# For debugging purposes 
            $chunk = $fileChunks[0]
            #> 
            Write-Progress -Activity "Searching folder '$baseFolder' for string '$searchString'" -Status "Searching files" -CurrentOperation " searched files $progress/$total" -PercentComplete ($progress*100/$total)           
                       
            $found = [array]($chunk | Select-String $searchString)
            if ($found.Count -gt 0){
                $Global:hits.AddRange($found)
            }
            $progress += $chunkSize
        }
        
        Write-Host "$searchString was found '$($global:Hits.Count)' times in '$($global:filesToSearch.Count)' files"
        Write-host 'You can use ' -nonewline; Write-Host '$global:filesToSearch' -ForegroundColor Cyan -NoNewline; Write-Host ' to get all files which have been searched'
        Write-Host '$global:hits | clip' -ForegroundColor Cyan -NoNewline; Write-Host ' will copy the result to the clipboard'
        Write-Host 'Save-HitsToDesktop' -ForegroundColor Cyan -NoNewline; Write-Host ' will create a file on the desktop'
        
        
        if ($target -eq "Clipboard") {        
            $Global:hits |clip      
            Write-Host "Results have been copied to clipboard"       
        } else{
            Save-HitsToDesktop  
        }
    }
    end{
        $global:lastFolder = $baseFolder
        $global:lastExcludedFolders = $ExcludedFolders
        $global:lastIncludedFileTypes  = $IncludedFileTypes
        $global:lastMaxFileSizeInKb =  $MaxFileSizeInKb
        $global:lastExcludedFileTypes = $ExcludedFileTypes
    }
}

<#
.SYNOPSIS
    Will save the values of $global:hits to a desktop file and open it.    
#>
function Save-HitsToDesktop (){
    $targetFilename = "$($env:USERPROFILE)\desktop\searchresult_$((Get-date).ToString("yyyyMMdd_HHmmss")).txt"
    #$Global:hits > $targetFilename
    $Global:hits | Out-File -Encoding UTF8 -FilePath $targetFilename -Width 1000
    explorer.exe $targetFilename
}



function Get-FoldersWithoutExcluded {
    [OutputType([System.Collections.ArrayList])]
    param(
        [string]$baseFolder,
        [array]$ExcludedFolders
    )
 
    $result = New-Object System.Collections.ArrayList
    $allFolders = [array](get-childitem -Path $baseFolder -Directory )
    foreach ($folder in $allFolders) {
        # $folder = $allFolders[0]
        if ($ExcludedFolders -contains $folder.Name) {
            continue
        }
        [void]$result.Add($folder.FullName) 
        $subFolders = [array](Get-FoldersWithoutExcluded -baseFolder $folder.Fullname -ExcludedFolders $ExcludedFolders)
        if ($subFolders.Count -gt 0) {
            [void]$result.AddRange($subFolders)
        }
    }
    return $result
} 

function Select-Chunk {

  <#
  Source/Credits:
  https://stackoverflow.com/questions/59259674/a-better-way-to-slice-an-array-or-a-list-in-powershell


  .SYNOPSIS
  Chunks pipeline input.
  
  .DESCRIPTION
  Chunks (partitions) pipeline input into arrays of a given size.
  
  By design, each such array is output as a *single* object to the pipeline,
  so that the next command in the pipeline can process it as a whole.
  
  That is, for the next command in the pipeline $_ contains an *array* of
  (at most) as many elements as specified via -ReadCount.
  
  .PARAMETER InputObject
  The pipeline input objects binds to this parameter one by one.
  Do not use it directly.
  
  .PARAMETER ReadCount
  The desired size of the chunks, i.e., how many input objects to collect
  in an array before sending that array to the pipeline.
  
  0 effectively means: collect *all* inputs and output a single array overall.
  
  .EXAMPLE
  1..7 | Select-Chunk 3 | ForEach-Object { "$_" }
  
  1 2 3
  4 5 6
  7
  
  The above shows that the ForEach-Object script block receive the following
  three arrays: (1, 2, 3), (4, 5, 6), and (, 7)
  #>
  
  [CmdletBinding(PositionalBinding = $false)]
  [OutputType([object[]])]
  param (
    [Parameter(ValueFromPipeline)] 
    $InputObject
    ,
    [Parameter(Mandatory, Position = 0)]
    [ValidateRange(0, [int]::MaxValue)]
    [int] $ReadCount
  )
      
  begin {
    $q = [System.Collections.Generic.Queue[object]]::new($ReadCount)
  }
      
  process {
    $q.Enqueue($InputObject)
    if ($q.Count -eq $ReadCount) {
      , $q.ToArray()
      $q.Clear()
    }
  }
      
  end {
    if ($q.Count) {
      , $q.ToArray()
    }
  }

}