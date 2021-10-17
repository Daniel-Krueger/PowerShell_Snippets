function Test-IsAdmin {
    return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

if ((Test-IsAdmin) -eq $false){
    throw "This moduel must be executed as an administrator"
}

function Get-NAVServiceFolder (){
    $services = Get-Service | where {$_.Name.toLower().StartsWith("MicrosoftDynamicsNavServer".ToLower())}
    $serviceExecutionPath = (gwmi win32_service|?{$_.name -eq $services[0].Name}).pathname
    $serviceFolder = $serviceExecutionPath.Substring(1,$serviceExecutionPath.IndexOf("Microsoft.Dynamics.Nav.Server.exe")-2)
    return $serviceFolder
    
}

<#
.# Removing and DELETING the app data
    
    $appName = "BootCamp "
    $serverinstance = "BC180"
    Get-NAVAppInfo -Name "$appName"  -ServerInstance $serverinstance
    Uninstall-NavApp -ServerInstance $serverinstance -Name $appName -Version "1.0.0.1"  -DoNotSaveData -Force
    Sync-NAVApp -ServerInstance $ServerInstance -Name $appName  -Version "1.0.0.1" -mode Clean -Force
    Unpublish-NavApp -ServerInstance $serverinstance -Name $appName -Version "1.0.0.1"

    Uninstall-NavApp -ServerInstance $serverinstance -Name $appName -Version "1.0.0.0"  -DoNotSaveData -Force
    Sync-NAVApp -ServerInstance $ServerInstance -Name $appName -Version "1.0.0.0" -mode Clean  -Force
    Unpublish-NavApp -ServerInstance $serverinstance -Name $appName -Version "1.0.0.0"

#>
<#
.SYNOPSIS
	Ensures that the provided app is published.
.DESCRIPTION
    If the app does not exist, it will be installed otherwise the existing one gets updated.
.NOTES
    Author:   Daniel Krüger
	Date:     2021-10-17
.PARAMETER file
    The path to the application. The default naming convention should be used Publisher_AppName_AppVersion
.PARAMETER instancesToExclude
    If there are multiple instances on the server an array of names can be provides which instances should be ignored.
.PARAMETER skipVerification
    Defines whether the Publish-NAVApp is called with SkipVerification
.EXAMPLE     
    #Ensure the app on all serverinstances
    Import-Module C:\Workspace\_Privat\PowerShell_Snippets\BC\NAVHelper.psm1 -force
    $file = "C:\Workspace\_Privat\PowerShell_Snippets\BC\COSMO CONSULT_BootCamp _1.0.0.1.app"
    $instancesToExclude = @()
    Ensure-NavApp  -file $file -instancesToExclude $instancesToExclude -skipVerification $true
.EXAMPLE     
    #Ensure the app on all serverinstances except the one, which name is NAV
    Import-Module C:\Workspace\_Privat\PowerShell_Snippets\BC\NAVHelper.psm1 -force
    $file = "C:\Workspace\_Privat\PowerShell_Snippets\BC\\COSMO CONSULT_BootCamp _1.0.0.0.app"
    $instancesToExclude = @("NAV")
    Ensure-NavApp  -file $file -instancesToExclude $instancesToExclude -skipVerification $true
.EXAMPLE     
    #Docker example with my folder
    Import-Module C:\Workspace\_Privat\PowerShell_Snippets\BC\NAVHelper.psm1 -force
    $file = "c:\run\my\COSMO CONSULT_BootCamp _1.0.0.0.app"
    $instancesToExclude = @()
    Ensure-NavApp  -file $file -instancesToExclude $instancesToExclude -skipVerification $true
#> 
function Ensure-NavApp {
     param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path -Path $_ -pathType Leaf})]
        [string]$file, 
        [array]$instancesToExclude,
        [bool]$skipVerification
     )
     begin{        
        $fileItem = Get-ChildItem $file
        $appName = $fileItem.Name.split("_")[1]
        $appVersion = ($fileItem.Name.split("_")[2]).Replace(".app","")      
        $serviceFolder  = Get-NAVServiceFolder
        import-module "$serviceFolder\NavAdminTool.ps1" -Force
        Start-Transcript "$($fileItem.Directory.FullName)\PublishApp_$($appName)_$($appVersion)_$((Get-Date).ToString("yyyy-MM-dd_HHmmss")).log"        
     }
     process{
        
        $services = Get-NAVServerInstance 
        foreach ($service in $services)
        {
            #$service = $services[0]
            $installApp = $true;
            foreach ($instanceToExclude in $instancesToExclude){
                if ($service.Serverinstance.endswith('$'+$instanceToExclude)){
                    Write-Host "App will not be imported to $($service.Serverinstance)" -ForegroundColor Yellow
                    $installApp = $false;
                    break;
                }
            }
            if ($installApp) {
                # Deploy to NAV
                $appExists = Get-NAVAppInfo -Name "$appName"  -ServerInstance $service.serverInstance
                $appVersionExists = Get-NAVAppInfo -Name "$appName" -Version $appVersion -ServerInstance $service.serverInstance
                if ($appExists -ne $null -and $appVersionExists -ne $null){
                    Write-Host "App '$appName' in version '$appVersion' is already deployed to server instance $($service.serverInstance)" -ForegroundColor DarkYellow
                    continue
                }

                Write-Host "Publishing App '$appName' in version '$appVersion' to server instance $($service.serverInstance)" -ForegroundColor Cyan
                $publishAppParam = @{}
                if ($skipVerification){
                    $publishAppParam["SkipVerification"] = $true
                }
                Publish-NAVApp -ServerInstance $service.serverInstance -Path $file @publishAppParam

                Write-Host "Syncing '$appName' in version '$appVersion' on server instance $($service.serverInstance)" -ForegroundColor Cyan
                # Sync to all 
                Sync-NavApp  -ServerInstance $service.serverInstance -Name $appName -Version $appVersion 

                # Install App
                if ($appExists -eq $null -and $appVersionExists -eq $null){
                    Write-Host "Installing '$appName' with version '$appVersion' on server instance $($service.serverInstance)" -ForegroundColor Cyan
                    Install-NAVApp  -ServerInstance $service.serverInstance -Name $appName -Version $appVersion
                }
                else{
                    Write-Host "Upgrading'$appName' to version '$appVersion' on server instance $($service.serverInstance)" -ForegroundColor Cyan
                    Start-NAVAppDataUpgrade -ServerInstance $service.serverInstance -Name $appName -Version $appVersion
                }
                Write-Host "App '$appName' in version '$appVersion' was deployed to server instance $($service.serverInstance)" -ForegroundColor Green
                       
            }
        }
    }
    end{     
        Stop-Transcript
    }
}


<#
.SYNOPSIS
	Ensures that the provided app is published.
.DESCRIPTION
    Imports a license and restarts the server instances.
.NOTES
    Author:   Daniel Krüger
	Date:     2021-10-17
.PARAMETER file
    The path to the license. 
.PARAMETER instancesToExclude
    If there are multiple instances on the server an array of names can be provides which instances should be ignored.
.EXAMPLE     
    #Installs the license to all serverinstances
    Import-Module C:\Workspace\_Privat\PowerShell_Snippets\BC\NAVHelper.psm1 -force
    $file = "C:\Workspace\_Privat\PowerShell_Snippets\BC\license.flf"
    $instancesToExclude = @()
    Import-License   -file $file -instancesToExclude $instancesToExclude
.EXAMPLE     
    #Docker example with my folder 
    Import-Module C:\Workspace\_Privat\PowerShell_Snippets\BC\NAVHelper.psm1 -force
    $file = "c:\run\my\license.flf"
    $instancesToExclude = @()
    Import-License  -file $file -instancesToExclude $instancesToExclude
#> 
function Import-License {
     param(
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path -Path $_ -pathType Leaf})]
        [string]$file, 
        [array]$instancesToExclude
     )
     begin{
        $fileItem = Get-ChildItem $file
        $serviceFolder  = Get-NAVServiceFolder        
        import-module "$serviceFolder\NavAdminTool.ps1"
        Start-Transcript "$($fileItem.Directory.FullName)\importLicense_$((Get-Date).ToString("yyyy-MM-dd_HHmmss")).log"        
     }
     process{
        
        $services = Get-NAVServerInstance 
        foreach ($service in $services)
        {
            $importLicense = $true;
            #$service = $services[0]
            foreach ($instanceToExclude in $instancesToExclude){
                if ($service.Serverinstance.endswith('$'+$instanceToExclude)){
                    Write-Host "License will not be imported to $($service.Serverinstance)"  -ForegroundColor Yellow
                    $importLicense = $false;
                    break;
                }
            }
            if ($importLicense) {
                Import-NAVServerLicense -ServerInstance $service.Serverinstance -LicenseFile $file -verbose
                Restart-NAVServerInstance -ServerInstance $service.serverInstance -verbose
            }
            Write-Host "License '$file' has been deployed and server instance $($service.serverInstance) has been restarted." -ForegroundColor Green         
        }
    }
    end{     
        Stop-Transcript
    }
}
