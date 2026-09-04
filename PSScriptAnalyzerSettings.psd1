@{
    # ImpactIQ targets Windows PowerShell 5.1 first (the one-click launcher runs there) and must also run on
    # PowerShell 7 (Azure DevOps hosted agents, local testing). The compatibility rules below flag syntax or
    # commands that only exist in one of the two.
    Severity     = @('Error', 'Warning')
    ExcludeRules = @(
        'PSAvoidUsingWriteHost',              # coloured progress output is intentional for the interactive launcher
        'PSUseBOMForUnicodeEncodedFile',
        'PSAvoidUsingConvertToSecureStringWithPlainText',  # env-var credentials are documented and opt-in
        'PSUseShouldProcessForStateChangingFunctions',
        'PSAvoidGlobalVars'
    )
    Rules        = @{
        PSUseCompatibleSyntax   = @{
            Enable         = $true
            TargetVersions = @('5.1', '7.0')
        }
        PSUseCompatibleCommands = @{
            Enable         = $true
            TargetProfiles = @(
                'win-8_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework',
                'win-8_x64_10.0.17763.0_7.0.0_x64_3.1.2_core'
            )
            IgnoreCommands = @(
                'Export-Excel', 'Import-Excel',                                   # ImportExcel module
                'Connect-PowerBIServiceAccount', 'Get-PowerBIAccessToken', 'Disconnect-PowerBIServiceAccount', 'Invoke-PowerBIRestMethod',
                'Connect-AzAccount', 'Get-AzAccessToken', 'Get-AzContext', 'Update-AzConfig', 'Enable-AzContextAutosave',
                'Invoke-ScriptAnalyzer', 'Invoke-Pester'
            )
        }
        PSAvoidUsingCmdletAliases = @{ Whitelist = @() }
    }
}
