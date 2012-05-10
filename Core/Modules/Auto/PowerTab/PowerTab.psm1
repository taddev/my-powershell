param()

## Load forms library when not loaded 
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {$_.ManifestModule -like "System.Windows.Forms*"})) {
    [Void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
}

## Load shares library
if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object {$_.ManifestModule -like "Shares.*"})) {
    [Void][System.Reflection.Assembly]::LoadFile((Join-Path $PSScriptRoot "Shares.dll"))
}

#########################
## Cleanup
#########################

$OldTabExpansion = Get-Content Function:TabExpansion
$Module = $MyInvocation.MyCommand.ScriptBlock.Module 
$Module.OnRemove = {
    Set-Content Function:\TabExpansion -Value $OldTabExpansion
}


#########################
## Private properties
#########################

$dsTabExpansionDatabase = New-Object System.Data.DataSet

$dsTabExpansionConfig = New-Object System.Data.DataSet

$TabExpansionCommandRegistry = @{}

$TabExpansionParameterRegistry = @{}

$TabExpansionCommandInfoRegistry = @{}

$TabExpansionParameterNameRegistry = @{}

$ConfigFileName = "PowerTabConfig.xml"


#########################
## Public properties
#########################

$PowerTabConfig = New-Object System.Management.Automation.PSObject

$PowerTabError = New-Object System.Collections.ArrayList	


#########################
## Functions
#########################

Import-Module (Join-Path $PSScriptRoot "Lerch.PowerShell.dll")
. (Join-Path $PSScriptRoot "TabExpansionResources.ps1")
Import-LocalizedData -BindingVariable "Resources" -FileName "Resources" -ErrorAction SilentlyContinue
. (Join-Path $PSScriptRoot "TabExpansionCore.ps1")
. (Join-Path $PSScriptRoot "TabExpansionLib.ps1")
. (Join-Path $PSScriptRoot "TabExpansionUtil.ps1")
. (Join-Path $PSScriptRoot "TabExpansionHandlers.ps1")
. (Join-Path $PSScriptRoot "ConsoleLib.ps1")
. (Join-Path $PSScriptRoot "Handlers\PSClientManager.ps1")
. (Join-Path $PSScriptRoot "Handlers\Robocopy.ps1")
. (Join-Path $PSScriptRoot "Handlers\Utilities.ps1")


#########################
## Initialization code
#########################

$ConfigurationPathParam = ""

. {
	[CmdletBinding(SupportsShouldProcess = $false,
		SupportsTransactions = $false,
		ConfirmImpact = "None",
		DefaultParameterSetName = "")]
    param(
		[Parameter(Position = 0)]
        [String]
        $ConfigurationPath = ""
    )

    if ($ConfigurationPath) {
        $script:ConfigurationPathParam = $ConfigurationPath
    } elseif ($PrivateData = (Parse-Manifest).PrivateData) {
        $script:ConfigurationPathParam = $PrivateData
    } elseif (Test-Path (Join-Path (Split-Path $Profile) $ConfigFileName)) {
        $script:ConfigurationPathParam = (Join-Path (Split-Path $Profile) $ConfigFileName)
    } elseif (Test-Path (Join-Path $PSScriptRoot $ConfigFileName)) {
        $script:ConfigurationPathParam = (Join-Path $PSScriptRoot $ConfigFileName)
    }
} @args

if ($ConfigurationPathParam) {
    if ((Test-Path $ConfigurationPathParam) -or (
            ($ConfigurationPathParam -eq "IsolatedStorage") -and (Test-IsolatedStoragePath "PowerTab\$ConfigFileName"))) {
        Initialize-PowerTab $ConfigurationPathParam
    } else {
        ## Config specified, but does not exist
        Write-Warning "Configuration File does not exist: '$ConfigurationPathParam'"  ## TODO: localize

        ## Create config and database
        New-TabExpansionConfig $ConfigurationPathParam
        CreatePowerTabConfig
        New-TabExpansionDatabase

        ## Update database
        Update-TabExpansionDataBase -Confirm

        ## Export changes
        Export-TabExpansionConfig
        Export-TabExpansionDatabase
    }
} else {
    $Yes = New-Object System.Management.Automation.Host.ChoiceDescription $Resources.global_choice_yes
    $No = New-Object System.Management.Automation.Host.ChoiceDescription $Resources.global_choice_no
    $YesNoChoices = [System.Management.Automation.Host.ChoiceDescription[]]($No,$Yes)

    ## Launch setup wizard?
    $Answer = $Host.UI.PromptForChoice($Resources.setup_wizard_caption, $Resources.setup_wizard_message, $YesNoChoices, 1)
    if ($Answer) {
        ## Ask for location to place config and database
        $ProfileDir = New-Object System.Management.Automation.Host.ChoiceDescription $Resources.setup_wizard_choice_profile_directory
        $InstallDir = New-Object System.Management.Automation.Host.ChoiceDescription $Resources.setup_wizard_choice_install_directory
        $AppDataDir = New-Object System.Management.Automation.Host.ChoiceDescription $Resources.setup_wizard_choice_appdata_directory
        $IsoStorageDir = New-Object System.Management.Automation.Host.ChoiceDescription $Resources.setup_wizard_choice_isostorage_directory
        $OtherDir = New-Object System.Management.Automation.Host.ChoiceDescription $Resources.setup_wizard_choice_other_directory
        $LocationChoices = [System.Management.Automation.Host.ChoiceDescription[]]($ProfileDir,$InstallDir,$AppDataDir,$IsoStorageDir,$OtherDir)
        $Answer = $Host.UI.PromptForChoice($Resources.setup_wizard_config_location_caption, $Resources.setup_wizard_config_location_message, $LocationChoices, 0)
        $SetupConfigurationPath = switch ($Answer) {
            0 {Split-Path $Profile}
            1 {$PSScriptRoot}
            2 {Join-Path ([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ApplicationData)) "PowerTab"}
            3 {"IsolatedStorage"}
            4 {
                $Path = Read-Host $Resources.setup_wizard_other_directory_prompt
                while ((-not $Path) -or -not (Test-Path -IsValid $Path)) {
                    ## TODO: Maybe write-error instead?
                    Write-Host $Resources.setup_wizard_err_path_not_valid -ForegroundColor $Host.PrivateData.ErrorForegroundColor `
                        -BackgroundColor $Host.PrivateData.ErrorBackgroundColor
                    $Path = Read-Host $Resources.setup_wizard_other_directory_prompt
                }
                $Path
            }
        }

        ## Create config in chosen location
        if ($SetupConfigurationPath -eq "IsolatedStorage") {
            New-TabExpansionConfig $SetupConfigurationPath
        } else {
            New-TabExpansionConfig (Join-Path $SetupConfigurationPath $ConfigFileName)
        }
        CreatePowerTabConfig

        ## Profile text
        ## TODO: Ask to update profile
        $ProfileText = @"

<############### Start of PowerTab Initialization Code ########################
    Added to profile by PowerTab setup for loading of custom tab expansion.
    Import other modules after this, they may contain PowerTab integration.
#>

Import-Module "PowerTab" -ArgumentList "$(Join-Path $SetupConfigurationPath $ConfigFileName)"
################ End of PowerTab Initialization Code ##########################

"@
        Write-Host ""
        Write-Host $Resources.setup_wizard_add_to_profile
        Write-Host $ProfileText

        ## Create new database or load existing database
        if ($SetupConfigurationPath -eq "IsolatedStorage") {
            $SetupDatabasePath = $SetupConfigurationPath
            if (Test-IsolatedStoragePath "PowerTab\TabExpansion.xml") {
                $Answer = $Host.UI.PromptForChoice($Resources.setup_wizard_upgrade_existing_database_caption, $Resources.setup_wizard_upgrade_existing_database_message, $YesNoChoices, 1)
            } else {
                $Answer = 0
            }
        } else {
            $SetupDatabasePath = Join-Path $SetupConfigurationPath "TabExpansion.xml"
            if (Test-Path $SetupDatabasePath) {
                $Answer = $Host.UI.PromptForChoice($Resources.setup_wizard_upgrade_existing_database_caption, $Resources.setup_wizard_upgrade_existing_database_message, $YesNoChoices, 1)
            } else {
                $Answer = 0
            }
        }
        if ($Answer) {
            Import-TabExpansionDataBase $SetupDatabasePath
        } else {
            New-TabExpansionDatabase
        }

        ## Update database
        Update-TabExpansionDataBase -Confirm

        ## Export changes
        Export-TabExpansionConfig
        Export-TabExpansionDatabase
        Write-Host ""
    } else {
        New-TabExpansionConfig
        CreatePowerTabConfig
        New-TabExpansionDatabase
    }
}

if ($PowerTabConfig.Enabled) {
    . "$PSScriptRoot\TabExpansion.ps1"
}

if ($PowerTabConfig.ShowBanner) {
    $CurVersion = (Parse-Manifest).ModuleVersion
    Write-Host -ForegroundColor 'Yellow' "PowerTab version ${CurVersion} PowerShell TabExpansion Library"
    Write-Host -ForegroundColor 'Yellow' "Host: $($Host.Name)"
    Write-Host -ForegroundColor 'Yellow' "PowerTab Enabled: $($PowerTabConfig.Enabled)"
}


## Exported functions, variables, etc.
$ExcludedFuctions = @("Initialize-TabExpansion")
$Functions = Get-Command "*-TabExpansion*","New-TabItem" | Where-Object {$ExcludedFuctions -notcontains $_.Name}
#$Functions = Get-Command "*-*" | Where-Object {$ExcludedFuctions -notcontains $_.Name}
Export-ModuleMember -Function $Functions -Variable PowerTabConfig, PowerTabError -Alias *

<#
TODOs
- Support variables in path:  $test = "C:"; $test\<TAB>
~ Expand items in a list:  Get-Command -CommandType Cm<TAB>,Fun<TAB>
- Assignment to strongly type variables:  $ErrorActionPreference = <TAB>
- Alias and Variable replace:  ls^A  or  $test^A

Just ideas:
- DateTime formats:  ^D<TAB>  or  2008/01/20^D<TAB>
- Paste clipboard:  ^V<TAB>
- Cut line:  Get-Foo -Bar something^X<TAB>  -->  
- Cut word:  Get-Foo -Bar something^Z<TAB>  -->  Get-Foo -Bar

- handle group start tokens ('{', '(', etc.)
~ Not detecting possitional parameters bound from pipeline
#>

# SIG # Begin signature block
# MIIY1AYJKoZIhvcNAQcCoIIYxTCCGMECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUB7vUjdGv+cPLsKybFqZnot/W
# TPegghSGMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
# AQUFADBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNpZ24sIEluYy4xKzAp
# BgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EwHhcNMDcw
# NjE1MDAwMDAwWhcNMTIwNjE0MjM1OTU5WjBcMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMOVmVyaVNpZ24sIEluYy4xNDAyBgNVBAMTK1ZlcmlTaWduIFRpbWUgU3RhbXBp
# bmcgU2VydmljZXMgU2lnbmVyIC0gRzIwgZ8wDQYJKoZIhvcNAQEBBQADgY0AMIGJ
# AoGBAMS18lIVvIiGYCkWSlsvS5Frh5HzNVRYNerRNl5iTVJRNHHCe2YdicjdKsRq
# CvY32Zh0kfaSrrC1dpbxqUpjRUcuawuSTksrjO5YSovUB+QaLPiCqljZzULzLcB1
# 3o2rx44dmmxMCJUe3tvvZ+FywknCnmA84eK+FqNjeGkUe60tAgMBAAGjgcQwgcEw
# NAYIKwYBBQUHAQEEKDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC52ZXJpc2ln
# bi5jb20wDAYDVR0TAQH/BAIwADAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3Js
# LnZlcmlzaWduLmNvbS90c3MtY2EuY3JsMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MA4GA1UdDwEB/wQEAwIGwDAeBgNVHREEFzAVpBMwETEPMA0GA1UEAxMGVFNBMS0y
# MA0GCSqGSIb3DQEBBQUAA4IBAQBQxUvIJIDf5A0kwt4asaECoaaCLQyDFYE3CoIO
# LLBaF2G12AX+iNvxkZGzVhpApuuSvjg5sHU2dDqYT+Q3upmJypVCHbC5x6CNV+D6
# 1WQEQjVOAdEzohfITaonx/LhhkwCOE2DeMb8U+Dr4AaH3aSWnl4MmOKlvr+ChcNg
# 4d+tKNjHpUtk2scbW72sOQjVOCKhM4sviprrvAchP0RBCQe1ZRwkvEjTRIDroc/J
# ArQUz1THFqOAXPl5Pl1yfYgXnixDospTzn099io6uE+UAKVtCoNd+V5T9BizVw9w
# w/v1rZWgDhfexBaAYMkPK26GBPHr9Hgn0QXF7jRbXrlJMvIzMIIDxDCCAy2gAwIB
# AgIQR78Zld+NUkZD99ttSA0xpDANBgkqhkiG9w0BAQUFADCBizELMAkGA1UEBhMC
# WkExFTATBgNVBAgTDFdlc3Rlcm4gQ2FwZTEUMBIGA1UEBxMLRHVyYmFudmlsbGUx
# DzANBgNVBAoTBlRoYXd0ZTEdMBsGA1UECxMUVGhhd3RlIENlcnRpZmljYXRpb24x
# HzAdBgNVBAMTFlRoYXd0ZSBUaW1lc3RhbXBpbmcgQ0EwHhcNMDMxMjA0MDAwMDAw
# WhcNMTMxMjAzMjM1OTU5WjBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNp
# Z24sIEluYy4xKzApBgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcgU2Vydmlj
# ZXMgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCpyrKkzM0grwp9
# iayHdfC0TvHfwQ+/Z2G9o2Qc2rv5yjOrhDCJWH6M22vdNp4Pv9HsePJ3pn5vPL+T
# rw26aPRslMq9Ui2rSD31ttVdXxsCn/ovax6k96OaphrIAuF/TFLjDmDsQBx+uQ3e
# P8e034e9X3pqMS4DmYETqEcgzjFzDVctzXg0M5USmRK53mgvqubjwoqMKsOLIYdm
# vYNYV291vzyqJoddyhAVPJ+E6lTBCm7E/sVK3bkHEZcifNs+J9EeeOyfMcnx5iIZ
# 28SzR0OaGl+gHpDkXvXufPF9q2IBj/VNC97QIlaolc2uiHau7roN8+RN2aD7aKCu
# FDuzh8G7AgMBAAGjgdswgdgwNAYIKwYBBQUHAQEEKDAmMCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC52ZXJpc2lnbi5jb20wEgYDVR0TAQH/BAgwBgEB/wIBADBBBgNV
# HR8EOjA4MDagNKAyhjBodHRwOi8vY3JsLnZlcmlzaWduLmNvbS9UaGF3dGVUaW1l
# c3RhbXBpbmdDQS5jcmwwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQD
# AgEGMCQGA1UdEQQdMBukGTAXMRUwEwYDVQQDEwxUU0EyMDQ4LTEtNTMwDQYJKoZI
# hvcNAQEFBQADgYEASmv56ljCRBwxiXmZK5a/gqwB1hxMzbCKWG7fCCmjXsjKkxPn
# BFIN70cnLwA4sOTJk06a1CJiFfc/NyFPcDGA8Ys4h7Po6JcA/s9Vlk4k0qknTnqu
# t2FB8yrO58nZXt27K4U+tZ212eFX/760xX71zwye8Jf+K9M7UhsbOCf3P0owggZw
# MIIEWKADAgECAgEkMA0GCSqGSIb3DQEBBQUAMH0xCzAJBgNVBAYTAklMMRYwFAYD
# VQQKEw1TdGFydENvbSBMdGQuMSswKQYDVQQLEyJTZWN1cmUgRGlnaXRhbCBDZXJ0
# aWZpY2F0ZSBTaWduaW5nMSkwJwYDVQQDEyBTdGFydENvbSBDZXJ0aWZpY2F0aW9u
# IEF1dGhvcml0eTAeFw0wNzEwMjQyMjAxNDZaFw0xNzEwMjQyMjAxNDZaMIGMMQsw
# CQYDVQQGEwJJTDEWMBQGA1UEChMNU3RhcnRDb20gTHRkLjErMCkGA1UECxMiU2Vj
# dXJlIERpZ2l0YWwgQ2VydGlmaWNhdGUgU2lnbmluZzE4MDYGA1UEAxMvU3RhcnRD
# b20gQ2xhc3MgMiBQcmltYXJ5IEludGVybWVkaWF0ZSBPYmplY3QgQ0EwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDKI4siNR6aoBs8nUnQPwyXOBYpuvh9
# iVtFWO+EcO1+EU3pFDGrQ+NNDFGBbPAVA0okJ1Tl+0qgzk3hhKMh3pk1q9xJrr8x
# xWeEMBCb7wfcdagPTfQ1U7FuOAP8iHcdpXf/P3Xn2ee/LFARyRFl+kkHYp+Tpoep
# bcmdK9F75dVlK58NUJ7++3EZITAoJo2uwtz2luhShggLejLNahRNnrn5zQfilpHx
# zx4r+YL3XiYGjo3R1DnXb9uRJ1p5j1hpCka1b+H9b8WRtBFPewKm20tWUiOeS5ji
# v37O+qFOg+PFx8NgR/5cPxUaQCqV7wBryFD4zWoZ1CMDJ7w7NtW5Q7DvAgMBAAGj
# ggHpMIIB5TAPBgNVHRMBAf8EBTADAQH/MA4GA1UdDwEB/wQEAwIBBjAdBgNVHQ4E
# FgQU0E4PQJlsuEsZbzsouODjiAc0qrcwHwYDVR0jBBgwFoAUTgvvGqRAW6UXaYcw
# yjRoQ9BBrvIwPQYIKwYBBQUHAQEEMTAvMC0GCCsGAQUFBzAChiFodHRwOi8vd3d3
# LnN0YXJ0c3NsLmNvbS9zZnNjYS5jcnQwWwYDVR0fBFQwUjAnoCWgI4YhaHR0cDov
# L3d3dy5zdGFydHNzbC5jb20vc2ZzY2EuY3JsMCegJaAjhiFodHRwOi8vY3JsLnN0
# YXJ0c3NsLmNvbS9zZnNjYS5jcmwwgYAGA1UdIAR5MHcwdQYLKwYBBAGBtTcBAgEw
# ZjAuBggrBgEFBQcCARYiaHR0cDovL3d3dy5zdGFydHNzbC5jb20vcG9saWN5LnBk
# ZjA0BggrBgEFBQcCARYoaHR0cDovL3d3dy5zdGFydHNzbC5jb20vaW50ZXJtZWRp
# YXRlLnBkZjARBglghkgBhvhCAQEEBAMCAAEwUAYJYIZIAYb4QgENBEMWQVN0YXJ0
# Q29tIENsYXNzIDIgUHJpbWFyeSBJbnRlcm1lZGlhdGUgT2JqZWN0IFNpZ25pbmcg
# Q2VydGlmaWNhdGVzMA0GCSqGSIb3DQEBBQUAA4ICAQBycwsDdVo3g4gT2XhBPk4S
# 1nLk8HIGK3egeKpCmBURCjsMdGyNcPkf8jJOK+kyKRpp5HEi/3ltpF3iGhRwzAOP
# gkiMLdYD0Wg0VXfVIyWMRlrrobxFAQJ0xJK5+B8Ni7VdD5xQrGEPcS0sYZwUaOMw
# vsRC/YiiXvjWsSzJxfAhdyvLF6IxtTZM+Ltfd6VvBAxzgkWUngHL0WEHO5kHUNXa
# w3aKsZVsLcb/X5LZ2g8OMvUJoSXBFr9PSqSra+8/FSCvICgKmlQUpWLDnKgZgL7P
# UZp6xZaI/V4UoAvTAjsiBK8vNTfLVWnu+xhrE5UGpm15sVNZEe1eMKwWutAGeC3R
# 3fdBtBEjmbCDMSntcn3G7l3pFVYzhM9FSx34MNmkEeb2azO+L2BUVvZkbupFFcJK
# rKzj6780sE9teL+b+VTTRw4NBOUL967COT0dC1GtdD/OqwElLpQn54sbDWo5+P4d
# UGX9lCl+guTsihaVFC9EvWzuiKsRqo9lQhZj+Cter2vqMMoCnctl0pCk86eeiC2q
# VTh/v+QuMQmGutz3yas5aZUwr8G4VEB9DmgNQydWYLMDMsyMp8ZxVb+Ix7DjXJ+G
# ApvCl/ObcsGvVm/6kQGByBbqidEtICfdcczR423P4CTEfqtF/oHaZiEsQQYtqkfx
# HUAwCjgFtUU5lHmRdwwLCjCCBsgwggWwoAMCAQICAgIHMA0GCSqGSIb3DQEBBQUA
# MIGMMQswCQYDVQQGEwJJTDEWMBQGA1UEChMNU3RhcnRDb20gTHRkLjErMCkGA1UE
# CxMiU2VjdXJlIERpZ2l0YWwgQ2VydGlmaWNhdGUgU2lnbmluZzE4MDYGA1UEAxMv
# U3RhcnRDb20gQ2xhc3MgMiBQcmltYXJ5IEludGVybWVkaWF0ZSBPYmplY3QgQ0Ew
# HhcNMTAxMDIzMDAyMjU5WhcNMTIxMDI0MDcyNzEzWjCBxTEgMB4GA1UEDRMXMjgw
# NjI4LVA3TFV5Q0ZyUWs1dEgyV3kxCzAJBgNVBAYTAlVTMRUwEwYDVQQIEwxTb3V0
# aCBEYWtvdGExEzARBgNVBAcTClJhcGlkIENpdHkxLTArBgNVBAsTJFN0YXJ0Q29t
# IFZlcmlmaWVkIENlcnRpZmljYXRlIE1lbWJlcjEUMBIGA1UEAxMLVGFkIERlVnJp
# ZXMxIzAhBgkqhkiG9w0BCQEWFHRhZGRldnJpZXNAZ21haWwuY29tMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6MJoHOBh3YnusPxodtJLsUFujkRUYojA
# v4zbVPIV5Tc/lHFYl4dHIFMyR6ML16ZcmKp1byhKbecj+nD6vtj8pOMdXSZW/YQs
# w92nyfaJbkV0DSqvQrecYDFBw010A3WvOIelgtCRQjxd9l/FGCpm5nIrV0AKJkKH
# DtIgR5LQ792XEkmPQXi85TPYRKlgdbSpA8+DTKKFH2PAcAYEUQSlDV0DyD5TL2Gs
# iK8FIBJSj0ZRfaBSMqZbiVAujtY7rgHU5DMleSRAT4zvIeE6F2+1x1S5MJnI23AD
# UQmHPPRnLgNKDXDvm/hcxnAnvY42OQDHnisrkO2wwUXdDHeavM6WFQIDAQABo4IC
# 9zCCAvMwCQYDVR0TBAIwADAOBgNVHQ8BAf8EBAMCB4AwOgYDVR0lAQH/BDAwLgYI
# KwYBBQUHAwMGCisGAQQBgjcCARUGCisGAQQBgjcCARYGCisGAQQBgjcKAw0wHQYD
# VR0OBBYEFEldJ6UCXfE7XFuqGd7n0QOVKUsNMB8GA1UdIwQYMBaAFNBOD0CZbLhL
# GW87KLjg44gHNKq3MIIBQgYDVR0gBIIBOTCCATUwggExBgsrBgEEAYG1NwECAjCC
# ASAwLgYIKwYBBQUHAgEWImh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL3BvbGljeS5w
# ZGYwNAYIKwYBBQUHAgEWKGh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL2ludGVybWVk
# aWF0ZS5wZGYwgbcGCCsGAQUFBwICMIGqMBQWDVN0YXJ0Q29tIEx0ZC4wAwIBARqB
# kUxpbWl0ZWQgTGlhYmlsaXR5LCBzZWUgc2VjdGlvbiAqTGVnYWwgTGltaXRhdGlv
# bnMqIG9mIHRoZSBTdGFydENvbSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBQb2xp
# Y3kgYXZhaWxhYmxlIGF0IGh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL3BvbGljeS5w
# ZGYwYwYDVR0fBFwwWjAroCmgJ4YlaHR0cDovL3d3dy5zdGFydHNzbC5jb20vY3J0
# YzItY3JsLmNybDAroCmgJ4YlaHR0cDovL2NybC5zdGFydHNzbC5jb20vY3J0YzIt
# Y3JsLmNybDCBiQYIKwYBBQUHAQEEfTB7MDcGCCsGAQUFBzABhitodHRwOi8vb2Nz
# cC5zdGFydHNzbC5jb20vc3ViL2NsYXNzMi9jb2RlL2NhMEAGCCsGAQUFBzAChjRo
# dHRwOi8vd3d3LnN0YXJ0c3NsLmNvbS9jZXJ0cy9zdWIuY2xhc3MyLmNvZGUuY2Eu
# Y3J0MCMGA1UdEgQcMBqGGGh0dHA6Ly93d3cuc3RhcnRzc2wuY29tLzANBgkqhkiG
# 9w0BAQUFAAOCAQEApcNTZ0O+n3FTUi2Z/iEKkhrX39yuZlpXoOHcdAKfdO8ZRDnu
# YipvJi9loevLS8EQGjtrroG8c6zmC/HWGDXvbNFCmvvduKVOpUUX1TkHcUfc3CHV
# jHsnui69XQYyMdbsu0tOEMY9m8DBno3/0GAXq9FrK8HLAlLkdT7eKoof7YcNLxfq
# q5J8/mwAaQUwyTfiZ4wd1AlW6c6HuuDauuX7yt7prIjH740/bQ7Y02m6c7kZRs59
# kH/+CEdcLTr5cvzi1+hQ6GQPeVOgHJYQCWtfMwevDppyzBIYl5ZHvC7CilCD1O8V
# aba4E8Y5h/fThhgF1XlHoGyt6g6QaslBiPg0QTGCA7gwggO0AgEBMIGTMIGMMQsw
# CQYDVQQGEwJJTDEWMBQGA1UEChMNU3RhcnRDb20gTHRkLjErMCkGA1UECxMiU2Vj
# dXJlIERpZ2l0YWwgQ2VydGlmaWNhdGUgU2lnbmluZzE4MDYGA1UEAxMvU3RhcnRD
# b20gQ2xhc3MgMiBQcmltYXJ5IEludGVybWVkaWF0ZSBPYmplY3QgQ0ECAgIHMAkG
# BSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJ
# AzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEWMCMG
# CSqGSIb3DQEJBDEWBBSmEvp5Izd1F9RWuuCaBZVh2SxXJDANBgkqhkiG9w0BAQEF
# AASCAQCXYEALTSlHyejyJk5fe3XFChbvnzcu8TS0Xqw6n+qpIm/rFwhHKI2CVZ15
# pXsQMU0tCkPDxy3VcPPD1OWlgXf0TTrojg4fLjbiUh2+qKLf7lORjOqVLhfSRGhx
# C+uIWHk4r96PPjBnAyg5ayAFvFawiFIz2N0y38H53nK1Zgkl5ZC5sx4uhvcA+usz
# ViZeNavBssaBkFKtPs3/PE7vHbBTcTHPhcNQ6yLGYKPHzWWtTQg6xSLSiHqMaW69
# LY9GSdAtOggoqk74LnxTOr6xcYlf19gsjgMs95b9MH9fpUVM5Bl56O+aa+U+Dx01
# bg6Ld/eHBHUXpsCioSDZ3YmxUaagoYIBfzCCAXsGCSqGSIb3DQEJBjGCAWwwggFo
# AgEBMGcwUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBAhA4Jdf6
# +GGvnvSQ5ya11lrVMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0xMTEwMjgyMjAxMTdaMCMGCSqGSIb3DQEJBDEW
# BBQ6vRMH8zOlW83zFCwYFvmbJaHe7zANBgkqhkiG9w0BAQEFAASBgCQwQd6WWdQ8
# TJUiWwEacePIAdi8/x4qjIN7bI95VQSHjWJ5sZYKUk1WEqrqyAGv1y+64GWRvXqZ
# zjyO2WKzUPlq4Ur6lurMR4E9yNNqi44+H373hvHGE7iAkJZ+DUaRPCfVaGfUGiPK
# 7wPweDrl4Q7Zwlq9jo5NMgp7t2KsP2Z7
# SIG # End signature block
