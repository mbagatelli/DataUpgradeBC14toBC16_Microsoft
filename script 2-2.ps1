###################### Parameters ######################
$DatabaseServer       = "ARQ-MBG"
#$DatabaseInstance     = "BC160"
$DatabaseName         = "GTOGHERMIGRATEDB_14to16_WORKV3" 
$ServiceName          = "BC160"
$DeveloperLicenseFile = "C:\Temp\MSDynLicenses\5190281 - D365BC 160 DEV.flf"
$BC15Version          = "15.8.43801.0"
$BaseAppPath          = "C:\Temp\Dynamics.365.BC.15953.W1.DVD_BC16\applications\BaseApp\Source\Microsoft_Base Application.app"
$SystemAppPath        = "C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\160\AL Development Environment\System.app"
$MicrosoftSysPath     = "C:\Temp\Dynamics.365.BC.15953.W1.DVD_BC16\applications\system application\source\Microsoft_System Application.app"
$MicrosoftApplicPath  = "C:\Temp\Dynamics.365.BC.15953.W1.DVD_BC16\applications\Application\Source\Microsoft_Application.app"
$CustomAppPath        = "C:\Users\Marco Bagatelli\Documents\AL\BC14toBC15\Growing Together_BC14toBC15_$CustomAppVersion.app"
$CustomAppName        = "BC14toBC15"
$CustomAppVersion     = "3.0.0.13"
$ServerInstance       = "BC160"

##################################################################


Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\160\Service\NavAdminTool.ps1'

Write-Host "Run the Invoke-NAVApplicationDatabaseConversion cmdlet to start the conversion"
Invoke-NAVApplicationDatabaseConversion -DatabaseServer $DatabaseServer -DatabaseName $DatabaseName

Write-Host "Set the server instance to connect to the application database."
Set-NAVServerConfiguration -ServerInstance $ServiceName -KeyName DatabaseName -KeyValue $DatabaseName

Write-Host "Configure the DestinationAppsForMigration setting of the server instance to table migration extension."
Set-NAVServerConfiguration -ServerInstance $ServiceName -KeyName "DestinationAppsForMigration" -KeyValue '[{"appId":"335d5fa2-777c-4340-a45e-8b92cf21860e", "name":"MicrosoftBC14to16", "publisher": "Marco Bagatelli"}]'

Write-Host "Disable task scheduler on the server instance for purposes of upgrade."
Set-NavServerConfiguration -ServerInstance $ServiceName -KeyName "EnableTaskScheduler" -KeyValue false

Write-Host "Restart the server instance."
Restart-NAVServerInstance -ServerInstance $ServiceName

Write-Host "Import the version 16 partner license."
Import-NAVServerLicense -ServerInstance $ServiceName -LicenseFile $DeveloperLicenseFile

Write-Host "Restart the server instance."
Restart-NAVServerInstance -ServerInstance $ServiceName

Write-Host "Change the application version"
Set-NAVApplication -ServerInstance $ServiceName -ApplicationVersion 16.5.15941 -Force

Write-Host "Publish version 16 system symbols extension."
Publish-NAVApp -ServerInstance $ServiceName -Path $SystemAppPath -PackageType SymbolsOnly

Write-Host "Publish the first version of the table migration extension, which is the version that contains the table objects."
Publish-NAVApp -ServerInstance $ServiceName -Path "C:\Users\Marco Bagatelli\Documents\AL\Vamos\MicrosoftBC14to16\Marco Bagatelli_MicrosoftBC14toBC16_1.0.0.0.app" -SkipVerification

Write-Host "Publish the empty versions of the following extensions"
Publish-NAVApp -ServerInstance $ServiceName -Path "C:\Users\Marco Bagatelli\Documents\AL\Vamos\SysApplication\Microsoft_System Application_14.0.0.0.app" -SkipVerification
Publish-NAVApp -ServerInstance $ServiceName -Path "C:\Users\Marco Bagatelli\Documents\AL\Vamos\BaseApp\Microsoft_Base Application_14.0.0.0.app" -SkipVerification
Publish-NAVApp -ServerInstance $ServiceName -Path "C:\Users\Marco Bagatelli\Documents\AL\Vamos\BC14toBC15\Growing Together_BC14toBC15_1.0.0.0.app" -SkipVerification

Write-Host "Restart the Service"
Restart-NAVServerInstance -ServerInstance $ServiceName

Write-Host "Synchronize the tenant with the application database."
Sync-NAVTenant -ServerInstance $ServiceName -Mode Sync

Write-Host "Synchronize the tenant with the table migration extension."
Sync-NAVApp -ServerInstance $ServiceName -Name "MicrosoftBC14toBC16" -Version 1.0.0.0
Sync-NAVApp -ServerInstance $ServiceName -Name "System Application" -Version 14.0.0.0
Sync-NAVApp -ServerInstance $ServiceName -Name "Base Application" -Version 14.0.0.0
Sync-NAVApp -ServerInstance $ServiceName -Name "BC14toBC15" -Version 1.0.0.0

Write-Host "Run the Data Upgrade"
Start-NAVDataUpgrade -ServerInstance $ServiceName -FunctionExecutionMode Serial -SkipCompanyInitialization #-SkipAppVersionCheck

## It will stop here, need to wait 3-5m for the data process to finish
#Write-Host "Pausing for 4 minutes"
#Start-Sleep -Seconds 240

Write-Host "Install the empty versions of the system, base, and custom extensions that you published."
Install-NAVApp -ServerInstance $ServiceName -Name "System Application" -Version 14.0.0.0
Install-NAVApp -ServerInstance $ServiceName -Name "Base Application" -Version 14.0.0.0
Install-NAVApp -ServerInstance $ServiceName -Name "BC14toBC15" -Version 1.0.0.0

Write-Host "Publish extensions using the Publish-NAVApp cmdlet like you did in previous steps."
Publish-NAVApp -ServerInstance $ServiceName -Path "C:\Users\Marco Bagatelli\Documents\AL\Vamos\MicrosoftBC14to16\Marco Bagatelli_MicrosoftBC14toBC16_1.0.0.1.app" -SkipVerification
Publish-NAVApp -ServerInstance $ServiceName -Path $MicrosoftSysPath
Publish-NAVApp -ServerInstance $ServiceName -Path $BaseAppPath
Publish-NAVApp -ServerInstance $ServiceName -Path "C:\Users\Marco Bagatelli\Documents\AL\Vamos\BC14toBC15\Growing Together_BC14toBC15_1.0.0.1.app" -SkipVerification

Write-Host "Synchronize the newly published extensions using the Sync-NAVApp cmdlet like you did in previous steps."
Sync-NAVApp -ServerInstance $ServiceName -Name "System Application" -Version 16.5.15897.15953
Sync-NAVApp -ServerInstance $ServiceName -Name "Base Application" -Version 16.5.15897.15953
Sync-NAVApp -ServerInstance $ServiceName -Name "BC14toBC15" -Version 1.0.0.1
Sync-NAVApp -ServerInstance $ServiceName -Name "MicrosoftBC14toBC16" -Version 1.0.0.1
