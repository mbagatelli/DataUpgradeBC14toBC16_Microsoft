# Before running delete all old tables and UPG tables

###################### Parameters ######################
$ServiceName          = "BC140"
##################################################################


Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\140\Service\NavAdminTool.ps1'
Import-Module 'C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\140\RoleTailored Client\NavModelTools.ps1'
Import-Module 'C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\140\RoleTailored Client\Microsoft.Dynamics.Nav.Ide.psm1'

Get-NAVAppInfo -ServerInstance $ServiceName | % { Uninstall-NAVApp -ServerInstance $ServiceName -Name $_.Name -Version $_.Version -Force}

Get-NAVAppInfo -ServerInstance $ServiceName | % { Unpublish-NAVApp -ServerInstance $ServiceName -Name $_.Name -Version $_.Version }

Get-NAVAppInfo -ServerInstance $ServiceName -SymbolsOnly | % { Unpublish-NAVApp -ServerInstance $ServiceName -Name $_.Name -Version $_.Version }

Stop-NAVServerInstance -ServerInstance $ServiceName