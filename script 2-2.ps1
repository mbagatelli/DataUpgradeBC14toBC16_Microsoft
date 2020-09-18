###################### Change Parameters ######################

$ServiceName                = "BC160"
$DatabaseName               = "GTOGHERMIGRATEDB_14to16_WORKV6" 
$DatabaseServer             = "ARQ-MBG"

$DvdBC16                    = "C:\Temp\Dynamics.365.BC.15953.W1.DVD_BC16"
$BC16Version                = "16.5.15941"

$CustomAppName              = "BC14toBC15"
$CustomAppPath              = "C:\Users\Marco Bagatelli\Documents\AL\Vamos\BC14toBC15\Growing Together_BC14toBC15_1.0.0.0.app"
$CustomAppPath2             = "C:\Users\Marco Bagatelli\Documents\AL\Vamos\BC14toBC15\Growing Together_BC14toBC15_1.0.0.1.app"
$DeveloperLicenseFile       = "C:\Temp\MSDynLicenses\5190281 - D365BC 160 DEV.flf"

$CustomSysPathApp           = "C:\Users\Marco Bagatelli\Documents\AL\Vamos\SysApplication\Microsoft_System Application_14.0.0.0.app"
$CustomBasePathApp          = "C:\Users\Marco Bagatelli\Documents\AL\Vamos\BaseApp\Microsoft_Base Application_14.0.0.0.app"
 
#$TableMigrationExtID        = 335d5fa2-777c-4340-a45e-8b92cf21860e
$TableMigrationExtPath      = "C:\Users\Marco Bagatelli\Documents\AL\Vamos\MicrosoftBC14to16\Marco Bagatelli_MicrosoftBC14toBC16_1.0.0.0.app"
$TableMigrationExtPath2     = "C:\Users\Marco Bagatelli\Documents\AL\Vamos\MicrosoftBC14to16\Marco Bagatelli_MicrosoftBC14toBC16_1.0.0.1.app"
$TableMigrationExtName      = "MicrosoftBC14toBC16"
$TableMigrationExtPublisher = "Marco Bagatelli"

$Minutes                    = 4 #How long to wait for the Data Upgrade

###################### No Need to Change ########################

$UpgradeBreak               = $Minutes * 60
$BaseAppPath                = "$DvdBC16\applications\BaseApp\Source\Microsoft_Base Application.app"
$MicrosoftSysPath           = "$DvdBC16\applications\system application\source\Microsoft_System Application.app"
$MicrosoftApplicPath        = "$DvdBC16\applications\Application\Source\Microsoft_Application.app"
$SystemAppPath              = "C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\160\AL Development Environment\System.app"
$ServicesAddinsFolder       = "C:\Program Files\Microsoft Dynamics 365 Business Central\160\Service\Add-ins"

##################################################################

Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\160\Service\NavAdminTool.ps1'

Write-Host "1. Run the Invoke-NAVApplicationDatabaseConversion cmdlet to start the conversion"
Invoke-NAVApplicationDatabaseConversion -DatabaseServer $DatabaseServer -DatabaseName $DatabaseName -Force

Write-Host "2. Set the server instance to connect to the application database."
Set-NAVServerConfiguration -ServerInstance $ServiceName -KeyName DatabaseName -KeyValue $DatabaseName

Write-Host "3. Configure the DestinationAppsForMigration setting of the server instance to table migration extension."
Set-NAVServerConfiguration -ServerInstance $ServiceName -KeyName "DestinationAppsForMigration" -KeyValue '[{"appId":"335d5fa2-777c-4340-a45e-8b92cf21860e", "name":"MicrosoftBC14toBC16", "publisher": "Marco Bagatelli" }]'
#Set-NAVServerConfiguration -ServerInstance $ServiceName -KeyName "DestinationAppsForMigration" -KeyValue '[{"appId":$TableMigrationExtID, "name":$TableMigrationExtName, "publisher": $TableMigrationExtPublisher }]'

Write-Host "4. Disable task scheduler on the server instance for purposes of upgrade."
Set-NavServerConfiguration -ServerInstance $ServiceName -KeyName "EnableTaskScheduler" -KeyValue false

Write-Host "5. Restart the server instance."
Restart-NAVServerInstance -ServerInstance $ServiceName

Write-Host "6. Import the version 16 partner license."
Import-NAVServerLicense -ServerInstance $ServiceName -LicenseFile $DeveloperLicenseFile

Write-Host "7. Restart the server instance."
Restart-NAVServerInstance -ServerInstance $ServiceName

Write-Host "8. Change the application version"
Set-NAVApplication -ServerInstance $ServiceName -ApplicationVersion $BC16Version -Force

Write-Host "9. Publish version 16 system symbols extension."
Publish-NAVApp -ServerInstance $ServiceName -Path $SystemAppPath -PackageType SymbolsOnly

Write-Host "10. Publish the first version of the table migration extension, which is the version that contains the table objects."
Publish-NAVApp -ServerInstance $ServiceName -Path $TableMigrationExtPath -SkipVerification

Write-Host "11. Publish the empty versions of the following extensions"
Publish-NAVApp -ServerInstance $ServiceName -Path $CustomSysPathApp -SkipVerification
Publish-NAVApp -ServerInstance $ServiceName -Path $CustomBasePathApp -SkipVerification
Publish-NAVApp -ServerInstance $ServiceName -Path $CustomAppPath -SkipVerification

Write-Host "12. Restart the Service"
Restart-NAVServerInstance -ServerInstance $ServiceName

Write-Host "13. Synchronize the tenant with the application database."
Sync-NAVTenant -ServerInstance $ServiceName -Mode Sync -Force

#Write-Host "14. Pausing for 1 minute"
#Start-Sleep -Seconds 60

Write-Host "15. Synchronize the tenant with the table migration extension."
Sync-NAVApp -ServerInstance $ServiceName -Name $TableMigrationExtName -Version 1.0.0.0
Sync-NAVApp -ServerInstance $ServiceName -Name "System Application" -Version 14.0.0.0
Sync-NAVApp -ServerInstance $ServiceName -Name "Base Application" -Version 14.0.0.0
Sync-NAVApp -ServerInstance $ServiceName -Name $CustomAppName -Version 1.0.0.0

Write-Host "16. Run the Data Upgrade"
Start-NAVDataUpgrade -ServerInstance $ServiceName -FunctionExecutionMode Serial -SkipCompanyInitialization -Force #-SkipAppVersionCheck 

## It will stop here, need to wait 3-5m for the data process to finish
Write-Host "17. Pausing for $Minutes minutes"
Start-Sleep -Seconds $UpgradeBreak

Write-Host "18. Install the empty versions of the system, base, and custom extensions that you published."
Install-NAVApp -ServerInstance $ServiceName -Name "System Application" -Version 14.0.0.0
Install-NAVApp -ServerInstance $ServiceName -Name "Base Application" -Version 14.0.0.0
Install-NAVApp -ServerInstance $ServiceName -Name $CustomAppName -Version 1.0.0.0

Write-Host "19. Publish the extensions."
Publish-NAVApp -ServerInstance $ServiceName -Path $TableMigrationExtPath2 -SkipVerification
Publish-NAVApp -ServerInstance $ServiceName -Path $MicrosoftSysPath
Publish-NAVApp -ServerInstance $ServiceName -Path $BaseAppPath
Publish-NAVApp -ServerInstance $ServiceName -Path $CustomAppPath2 -SkipVerification

Write-Host "20. Synchronize the newly published extensions."
Sync-NAVApp -ServerInstance $ServiceName -Name "System Application" -Version 16.5.15897.15953
Sync-NAVApp -ServerInstance $ServiceName -Name "Base Application" -Version 16.5.15897.15953
Sync-NAVApp -ServerInstance $ServiceName -Name $CustomAppName -Version 1.0.0.1
Sync-NAVApp -ServerInstance $ServiceName -Name $TableMigrationExtName -Version 1.0.0.1

Write-Host "21. Run data upgrade on the table migration extension"
Start-NAVAppDataUpgrade -ServerInstance $ServiceName -Name $TableMigrationExtName -version 1.0.0.1

Write-Host "22. Upgrade final extensions"
Start-NAVAppDataUpgrade -ServerInstance $ServiceName -Name "System Application" -version 16.5.15897.15953
Start-NAVAppDataUpgrade -ServerInstance $ServiceName -Name "Base Application" -Version 16.5.15897.15953
Start-NAVAppDataUpgrade -ServerInstance $ServiceName -Name $CustomAppName -Version 1.0.0.1

Write-Host "23. Upgrade control add-ins"
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.BusinessChart' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'BusinessChart\Microsoft.Dynamics.Nav.Client.BusinessChart.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.FlowIntegration' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'FlowIntegration\Microsoft.Dynamics.Nav.Client.FlowIntegration.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.OAuthIntegration' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'OAuthIntegration\Microsoft.Dynamics.Nav.Client.OAuthIntegration.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.PageReady' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'PageReady\Microsoft.Dynamics.Nav.Client.PageReady.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.PowerBIManagement' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'PowerBIManagement\Microsoft.Dynamics.Nav.Client.PowerBIManagement.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.RoleCenterSelector' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'RoleCenterSelector\Microsoft.Dynamics.Nav.Client.RoleCenterSelector.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.SatisfactionSurvey' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'SatisfactionSurvey\Microsoft.Dynamics.Nav.Client.SatisfactionSurvey.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.SocialListening' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'SocialListening\Microsoft.Dynamics.Nav.Client.SocialListening.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.VideoPlayer' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'VideoPlayer\Microsoft.Dynamics.Nav.Client.VideoPlayer.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.WebPageViewer' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'WebPageViewer\Microsoft.Dynamics.Nav.Client.WebPageViewer.zip')
Set-NAVAddIn -ServerInstance $ServiceName -AddinName 'Microsoft.Dynamics.Nav.Client.WelcomeWizard' -PublicKeyToken 31bf3856ad364e35 -ResourceFile ($AppName = Join-Path $ServicesAddinsFolder 'WelcomeWizard\Microsoft.Dynamics.Nav.Client.WelcomeWizard.zip')

Write-Host "24. Enable task scheduler on the server instance."
Set-NavServerConfiguration -ServerInstance $ServiceName -KeyName "EnableTaskScheduler" -KeyValue true

Write-Host "25. Restart the Service"
Restart-NAVServerInstance -ServerInstance $ServiceName

Write-Host "26 Unpublishing apps"
Unpublish-NAVApp -ServerInstance $ServiceName -Name "Base Application" -Version 14.0.0.0
Unpublish-NAVApp -ServerInstance $ServiceName -Name "System Application" -Version 14.0.0.0
Unpublish-NAVApp -ServerInstance $ServiceName -Name $CustomAppName -Version 1.0.0.0
Unpublish-NAVApp -ServerInstance $ServiceName -Name $TableMigrationExtName -Version 1.0.0.0

Write-Host "27 Publish & Install System App"
Publish-NAVApp -ServerInstance $ServiceName -Path $MicrosoftApplicPath -SkipVerification
Sync-NAVApp -ServerInstance $ServiceName -Name "Application" -Version 16.5.15897.15953 
Install-NAVApp -ServerInstance $ServiceName -Name "Application" -Version 16.5.15897.15953



#    If you want to use data encryption as before, enable it.
#
#    Grant users permission to the Open in Excel and Edit in Excel actions.
#    Version 16 introduces a system permission that protects these two actions. The permission is granted by the system object 6110 Allow Action Export To Excel. 
#    Because of this change, users who had permission to these actions before upgrading, will lose permission. To grant permission again, do one of the following steps:
#    Assign the EXCEL EXPORT ACTION permission set to appropriate users.
#    Add the system object 6110 Allow Action Export To Excel permission directly to appropriate permission sets.