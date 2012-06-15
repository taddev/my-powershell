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
# MIIY+QYJKoZIhvcNAQcCoIIY6jCCGOYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUB7vUjdGv+cPLsKybFqZnot/W
# TPegghSrMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
# AQUFADBTMQswCQYDVQQGEwJVUzEXMBUGA1UEChMOVmVyaVNpZ24sIEluYy4xKzAp
# BgNVBAMTIlZlcmlTaWduIFRpbWUgU3RhbXBpbmcgU2VydmljZXMgQ0EwHhcNMTIw
# NTAxMDAwMDAwWhcNMTIxMjMxMjM1OTU5WjBiMQswCQYDVQQGEwJVUzEdMBsGA1UE
# ChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xNDAyBgNVBAMTK1N5bWFudGVjIFRpbWUg
# U3RhbXBpbmcgU2VydmljZXMgU2lnbmVyIC0gRzMwgZ8wDQYJKoZIhvcNAQEBBQAD
# gY0AMIGJAoGBAKlZZnTaPYp9etj89YBEe/5HahRVTlBHC+zT7c72OPdPabmx8LZ4
# ggqMdhZn4gKttw2livYD/GbT/AgtzLVzWXuJ3DNuZlpeUje0YtGSWTUUi0WsWbJN
# JKKYlGhCcp86aOJri54iLfSYTprGr7PkoKs8KL8j4ddypPIQU2eud69RAgMBAAGj
# geMwgeAwDAYDVR0TAQH/BAIwADAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3Js
# LnZlcmlzaWduLmNvbS90c3MtY2EuY3JsMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AudmVyaXNp
# Z24uY29tMA4GA1UdDwEB/wQEAwIHgDAeBgNVHREEFzAVpBMwETEPMA0GA1UEAxMG
# VFNBMS0zMB0GA1UdDgQWBBS0t/GJSSZg52Xqc67c0zjNv1eSbzANBgkqhkiG9w0B
# AQUFAAOCAQEAHpiqJ7d4tQi1yXJtt9/ADpimNcSIydL2bfFLGvvV+S2ZAJ7R55uL
# 4T+9OYAMZs0HvFyYVKaUuhDRTour9W9lzGcJooB8UugOA9ZresYFGOzIrEJ8Byyn
# PQhm3ADt/ZQdc/JymJOxEdaP747qrPSWUQzQjd8xUk9er32nSnXmTs4rnykr589d
# nwN+bid7I61iKWavkugszr2cf9zNFzxDwgk/dUXHnuTXYH+XxuSqx2n1/M10rCyw
# SMFQTnBWHrU1046+se2svf4M7IV91buFZkQZXZ+T64K6Y57TfGH/yBvZI1h/MKNm
# oTkmXpLDPMs3Mvr1o43c1bCj6SU2VdeB+jCCA8QwggMtoAMCAQICEEe/GZXfjVJG
# Q/fbbUgNMaQwDQYJKoZIhvcNAQEFBQAwgYsxCzAJBgNVBAYTAlpBMRUwEwYDVQQI
# EwxXZXN0ZXJuIENhcGUxFDASBgNVBAcTC0R1cmJhbnZpbGxlMQ8wDQYDVQQKEwZU
# aGF3dGUxHTAbBgNVBAsTFFRoYXd0ZSBDZXJ0aWZpY2F0aW9uMR8wHQYDVQQDExZU
# aGF3dGUgVGltZXN0YW1waW5nIENBMB4XDTAzMTIwNDAwMDAwMFoXDTEzMTIwMzIz
# NTk1OVowUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBMIIBIjAN
# BgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAqcqypMzNIK8KfYmsh3XwtE7x38EP
# v2dhvaNkHNq7+cozq4QwiVh+jNtr3TaeD7/R7Hjyd6Z+bzy/k68Numj0bJTKvVIt
# q0g99bbVXV8bAp/6L2sepPejmqYayALhf0xS4w5g7EAcfrkN3j/HtN+HvV96ajEu
# A5mBE6hHIM4xcw1XLc14NDOVEpkSud5oL6rm48KKjCrDiyGHZr2DWFdvdb88qiaH
# XcoQFTyfhOpUwQpuxP7FSt25BxGXInzbPifRHnjsnzHJ8eYiGdvEs0dDmhpfoB6Q
# 5F717nzxfatiAY/1TQve0CJWqJXNroh2ru66DfPkTdmg+2igrhQ7s4fBuwIDAQAB
# o4HbMIHYMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29jc3Au
# dmVyaXNpZ24uY29tMBIGA1UdEwEB/wQIMAYBAf8CAQAwQQYDVR0fBDowODA2oDSg
# MoYwaHR0cDovL2NybC52ZXJpc2lnbi5jb20vVGhhd3RlVGltZXN0YW1waW5nQ0Eu
# Y3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIBBjAkBgNVHREE
# HTAbpBkwFzEVMBMGA1UEAxMMVFNBMjA0OC0xLTUzMA0GCSqGSIb3DQEBBQUAA4GB
# AEpr+epYwkQcMYl5mSuWv4KsAdYcTM2wilhu3wgpo17IypMT5wRSDe9HJy8AOLDk
# yZNOmtQiYhX3PzchT3AxgPGLOIez6OiXAP7PVZZOJNKpJ056rrdhQfMqzufJ2V7d
# uyuFPrWdtdnhV/++tMV+9c8MnvCX/ivTO1IbGzgn9z9KMIIGcDCCBFigAwIBAgIB
# JDANBgkqhkiG9w0BAQUFADB9MQswCQYDVQQGEwJJTDEWMBQGA1UEChMNU3RhcnRD
# b20gTHRkLjErMCkGA1UECxMiU2VjdXJlIERpZ2l0YWwgQ2VydGlmaWNhdGUgU2ln
# bmluZzEpMCcGA1UEAxMgU3RhcnRDb20gQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkw
# HhcNMDcxMDI0MjIwMTQ2WhcNMTcxMDI0MjIwMTQ2WjCBjDELMAkGA1UEBhMCSUwx
# FjAUBgNVBAoTDVN0YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBEaWdpdGFs
# IENlcnRpZmljYXRlIFNpZ25pbmcxODA2BgNVBAMTL1N0YXJ0Q29tIENsYXNzIDIg
# UHJpbWFyeSBJbnRlcm1lZGlhdGUgT2JqZWN0IENBMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEAyiOLIjUemqAbPJ1J0D8MlzgWKbr4fYlbRVjvhHDtfhFN
# 6RQxq0PjTQxRgWzwFQNKJCdU5ftKoM5N4YSjId6ZNavcSa6/McVnhDAQm+8H3HWo
# D030NVOxbjgD/Ih3HaV3/z9159nnvyxQEckRZfpJB2Kfk6aHqW3JnSvRe+XVZSuf
# DVCe/vtxGSEwKCaNrsLc9pboUoYIC3oyzWoUTZ65+c0H4paR8c8eK/mC914mBo6N
# 0dQ512/bkSdaeY9YaQpGtW/h/W/FkbQRT3sCpttLVlIjnkuY4r9+zvqhToPjxcfD
# YEf+XD8VGkAqle8Aa8hQ+M1qGdQjAye8OzbVuUOw7wIDAQABo4IB6TCCAeUwDwYD
# VR0TAQH/BAUwAwEB/zAOBgNVHQ8BAf8EBAMCAQYwHQYDVR0OBBYEFNBOD0CZbLhL
# GW87KLjg44gHNKq3MB8GA1UdIwQYMBaAFE4L7xqkQFulF2mHMMo0aEPQQa7yMD0G
# CCsGAQUFBwEBBDEwLzAtBggrBgEFBQcwAoYhaHR0cDovL3d3dy5zdGFydHNzbC5j
# b20vc2ZzY2EuY3J0MFsGA1UdHwRUMFIwJ6AloCOGIWh0dHA6Ly93d3cuc3RhcnRz
# c2wuY29tL3Nmc2NhLmNybDAnoCWgI4YhaHR0cDovL2NybC5zdGFydHNzbC5jb20v
# c2ZzY2EuY3JsMIGABgNVHSAEeTB3MHUGCysGAQQBgbU3AQIBMGYwLgYIKwYBBQUH
# AgEWImh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL3BvbGljeS5wZGYwNAYIKwYBBQUH
# AgEWKGh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL2ludGVybWVkaWF0ZS5wZGYwEQYJ
# YIZIAYb4QgEBBAQDAgABMFAGCWCGSAGG+EIBDQRDFkFTdGFydENvbSBDbGFzcyAy
# IFByaW1hcnkgSW50ZXJtZWRpYXRlIE9iamVjdCBTaWduaW5nIENlcnRpZmljYXRl
# czANBgkqhkiG9w0BAQUFAAOCAgEAcnMLA3VaN4OIE9l4QT5OEtZy5PByBit3oHiq
# QpgVEQo7DHRsjXD5H/IyTivpMikaaeRxIv95baRd4hoUcMwDj4JIjC3WA9FoNFV3
# 1SMljEZa66G8RQECdMSSufgfDYu1XQ+cUKxhD3EtLGGcFGjjML7EQv2Iol741rEs
# ycXwIXcryxeiMbU2TPi7X3elbwQMc4JFlJ4By9FhBzuZB1DV2sN2irGVbC3G/1+S
# 2doPDjL1CaElwRa/T0qkq2vvPxUgryAoCppUFKViw5yoGYC+z1GaesWWiP1eFKAL
# 0wI7IgSvLzU3y1Vp7vsYaxOVBqZtebFTWRHtXjCsFrrQBngt0d33QbQRI5mwgzEp
# 7XJ9xu5d6RVWM4TPRUsd+DDZpBHm9mszvi9gVFb2ZG7qRRXCSqys4+u/NLBPbXi/
# m/lU00cODQTlC/euwjk9HQtRrXQ/zqsBJS6UJ+eLGw1qOfj+HVBl/ZQpfoLk7IoW
# lRQvRL1s7oirEaqPZUIWY/grXq9r6jDKAp3LZdKQpPOnnogtqlU4f7/kLjEJhrrc
# 98mrOWmVMK/BuFRAfQ5oDUMnVmCzAzLMjKfGcVW/iMew41yfhgKbwpfzm3LBr1Zv
# +pEBgcgW6onRLSAn3XHM0eNtz+AkxH6rRf6B2mYhLEEGLapH8R1AMAo4BbVFOZR5
# kXcMCwowggbIMIIFsKADAgECAgICBzANBgkqhkiG9w0BAQUFADCBjDELMAkGA1UE
# BhMCSUwxFjAUBgNVBAoTDVN0YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBE
# aWdpdGFsIENlcnRpZmljYXRlIFNpZ25pbmcxODA2BgNVBAMTL1N0YXJ0Q29tIENs
# YXNzIDIgUHJpbWFyeSBJbnRlcm1lZGlhdGUgT2JqZWN0IENBMB4XDTEwMTAyMzAw
# MjI1OVoXDTEyMTAyNDA3MjcxM1owgcUxIDAeBgNVBA0TFzI4MDYyOC1QN0xVeUNG
# clFrNXRIMld5MQswCQYDVQQGEwJVUzEVMBMGA1UECBMMU291dGggRGFrb3RhMRMw
# EQYDVQQHEwpSYXBpZCBDaXR5MS0wKwYDVQQLEyRTdGFydENvbSBWZXJpZmllZCBD
# ZXJ0aWZpY2F0ZSBNZW1iZXIxFDASBgNVBAMTC1RhZCBEZVZyaWVzMSMwIQYJKoZI
# hvcNAQkBFhR0YWRkZXZyaWVzQGdtYWlsLmNvbTCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAOjCaBzgYd2J7rD8aHbSS7FBbo5EVGKIwL+M21TyFeU3P5Rx
# WJeHRyBTMkejC9emXJiqdW8oSm3nI/pw+r7Y/KTjHV0mVv2ELMPdp8n2iW5FdA0q
# r0K3nGAxQcNNdAN1rziHpYLQkUI8XfZfxRgqZuZyK1dACiZChw7SIEeS0O/dlxJJ
# j0F4vOUz2ESpYHW0qQPPg0yihR9jwHAGBFEEpQ1dA8g+Uy9hrIivBSASUo9GUX2g
# UjKmW4lQLo7WO64B1OQzJXkkQE+M7yHhOhdvtcdUuTCZyNtwA1EJhzz0Zy4DSg1w
# 75v4XMZwJ72ONjkAx54rK5DtsMFF3Qx3mrzOlhUCAwEAAaOCAvcwggLzMAkGA1Ud
# EwQCMAAwDgYDVR0PAQH/BAQDAgeAMDoGA1UdJQEB/wQwMC4GCCsGAQUFBwMDBgor
# BgEEAYI3AgEVBgorBgEEAYI3AgEWBgorBgEEAYI3CgMNMB0GA1UdDgQWBBRJXSel
# Al3xO1xbqhne59EDlSlLDTAfBgNVHSMEGDAWgBTQTg9AmWy4SxlvOyi44OOIBzSq
# tzCCAUIGA1UdIASCATkwggE1MIIBMQYLKwYBBAGBtTcBAgIwggEgMC4GCCsGAQUF
# BwIBFiJodHRwOi8vd3d3LnN0YXJ0c3NsLmNvbS9wb2xpY3kucGRmMDQGCCsGAQUF
# BwIBFihodHRwOi8vd3d3LnN0YXJ0c3NsLmNvbS9pbnRlcm1lZGlhdGUucGRmMIG3
# BggrBgEFBQcCAjCBqjAUFg1TdGFydENvbSBMdGQuMAMCAQEagZFMaW1pdGVkIExp
# YWJpbGl0eSwgc2VlIHNlY3Rpb24gKkxlZ2FsIExpbWl0YXRpb25zKiBvZiB0aGUg
# U3RhcnRDb20gQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUG9saWN5IGF2YWlsYWJs
# ZSBhdCBodHRwOi8vd3d3LnN0YXJ0c3NsLmNvbS9wb2xpY3kucGRmMGMGA1UdHwRc
# MFowK6ApoCeGJWh0dHA6Ly93d3cuc3RhcnRzc2wuY29tL2NydGMyLWNybC5jcmww
# K6ApoCeGJWh0dHA6Ly9jcmwuc3RhcnRzc2wuY29tL2NydGMyLWNybC5jcmwwgYkG
# CCsGAQUFBwEBBH0wezA3BggrBgEFBQcwAYYraHR0cDovL29jc3Auc3RhcnRzc2wu
# Y29tL3N1Yi9jbGFzczIvY29kZS9jYTBABggrBgEFBQcwAoY0aHR0cDovL3d3dy5z
# dGFydHNzbC5jb20vY2VydHMvc3ViLmNsYXNzMi5jb2RlLmNhLmNydDAjBgNVHRIE
# HDAahhhodHRwOi8vd3d3LnN0YXJ0c3NsLmNvbS8wDQYJKoZIhvcNAQEFBQADggEB
# AKXDU2dDvp9xU1Itmf4hCpIa19/crmZaV6Dh3HQCn3TvGUQ57mIqbyYvZaHry0vB
# EBo7a66BvHOs5gvx1hg172zRQpr73bilTqVFF9U5B3FH3Nwh1Yx7J7ouvV0GMjHW
# 7LtLThDGPZvAwZ6N/9BgF6vRayvBywJS5HU+3iqKH+2HDS8X6quSfP5sAGkFMMk3
# 4meMHdQJVunOh7rg2rrl+8re6ayIx++NP20O2NNpunO5GUbOfZB//ghHXC06+XL8
# 4tfoUOhkD3lToByWEAlrXzMHrw6acswSGJeWR7wuwopQg9TvFWm2uBPGOYf304YY
# BdV5R6BsreoOkGrJQYj4NEExggO4MIIDtAIBATCBkzCBjDELMAkGA1UEBhMCSUwx
# FjAUBgNVBAoTDVN0YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBEaWdpdGFs
# IENlcnRpZmljYXRlIFNpZ25pbmcxODA2BgNVBAMTL1N0YXJ0Q29tIENsYXNzIDIg
# UHJpbWFyeSBJbnRlcm1lZGlhdGUgT2JqZWN0IENBAgICBzAJBgUrDgMCGgUAoHgw
# GAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGC
# NwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFjAjBgkqhkiG9w0BCQQx
# FgQUphL6eSM3dRfUVrrgmgWVYdksVyQwDQYJKoZIhvcNAQEBBQAEggEAl2BAC00p
# R8no8iZOX3t1xQoW7583LvE0tF6sOp/qqSJv6xcIRyiNglWdeaV7EDFNLQpDw8ct
# 1XDzw9TlpYF39E066I4OHy424lIdvqii3+5TkYzqlS4X0kRocQvriFh5OK/ejz4w
# ZwMoOWsgBbxWsIhSM9jdMt/B+d5ytWYJJeWQubMeLob3APrrM1YmXjWrwbLGgZBS
# rT7N/zxO7x2wU3Exz4XDUOsixmCjx81lrU0IOsUi0oh6jGluvS2PRknQLToIKKpO
# +C58Uzq+sXGJX9fYLI4DLPeW/TB/X6VFTOQZeejvmmvlPg8dNW4Oi3f3hwR1F6bA
# oqEg2d2JsVGmoKGCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBnMFMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UEAxMiVmVy
# aVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT2bg+9raN
# 7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTIwNjE1MTgwNzIzWjAjBgkqhkiG9w0BCQQxFgQUOr0TB/MzpVvN
# 8xQsGBb5myWh3u8wDQYJKoZIhvcNAQEBBQAEgYBXrTZmfGJ7+DgRBuFaEYdWfSLN
# DHz7Bhae/ZcsV9jINK/l656BTXe/KKNCzoHvAsOK00PsyxMmirm0G2UkJFo1YVgg
# J3PBH1mwzxYM7lIukeanJbVWHRIpjz+4oi+/HbvIbHJi3PQDf9lrL0VLfqnsfZwS
# 03ZHz+3nbvcFJ7sjSg==
# SIG # End signature block
