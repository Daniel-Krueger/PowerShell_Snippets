Import-Module .\NAVHelper.psm1 -force
$file = "C:\Temp\German Developer BC V14_W1.flf"
$instancesToExclude = @(BC180)
Import-License  -file $file -instancesToExclude $instancesToExclude