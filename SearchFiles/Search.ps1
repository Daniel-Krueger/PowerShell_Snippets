param($baseFolder,$searchString) 
Import-Module .\search.psm1 

$defaultParameters = @{
    #"ExcludedFolders"= @(".devops",".git",".vs",".vscode","packages","node_modules","debug","bin","dist","Ruby","Ruby27-x64")
    #"IncludedFileTypes" = @()
    #"ExcludedFileTypes" = @(".dll")
    #"MaxFileSizeInKb"= 100
    "Target" = "Desktop"
}
Search-FilesWithExcludedFolders -baseFolder $baseFolder -searchString $searchString   @defaultParameters



