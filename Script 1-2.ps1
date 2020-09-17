Import-Module 'C:\Program Files\Microsoft Dynamics 365 Business Central\140\Service\NavAdminTool.ps1'
Import-Module 'C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\140\RoleTailored Client\NavModelTools.ps1'
Import-Module 'C:\Program Files (x86)\Microsoft Dynamics 365 Business Central\140\RoleTailored Client\Microsoft.Dynamics.Nav.Ide.psm1'

Get-NAVAppInfo -ServerInstance bc140 | % { Uninstall-NAVApp -ServerInstance bc140 -Name $_.Name -Version $_.Version -Force}

Get-NAVAppInfo -ServerInstance bc140 | % { Unpublish-NAVApp -ServerInstance bc140 -Name $_.Name -Version $_.Version }

Get-NAVAppInfo -ServerInstance bc140 -SymbolsOnly | % { Unpublish-NAVApp -ServerInstance bc140 -Name $_.Name -Version $_.Version }

Stop-NAVServerInstance -ServerInstance bc140