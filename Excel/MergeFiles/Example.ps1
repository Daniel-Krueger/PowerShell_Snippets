Import-Module  -Name ".\Merge.psm1" -Force
$filePaths = get-ChildItem ".\ExampleFiles" -filter "*.xlsx"
$outputFilePath = ".\ResultFile\output.xlsx" 

Merge-Files -filePaths $filePaths -outputFilePath $outPutFilePath