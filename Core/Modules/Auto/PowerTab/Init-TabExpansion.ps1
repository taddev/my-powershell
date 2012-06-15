param ($ConfigurationLocation,[switch]$NoWarn)

$InitScriptVersion = 'PowerTab version 0.98' 

if ($global:PowerTabConfig -and (-not $NoWarn)) {
  Write-Warning "$($PowerTabConfig.Version) found loaded already"
}

# Read Configuration 

$filename = 'PowerTabConfig.xml'
#"$ConfigurationLocation\$fileName"
$global:dsTabExpansion = new-object data.dataset

[void]$global:dsTabExpansion.ReadXml("$ConfigurationLocation\$fileName",'InferSchema')

$installDir = $global:dsTabExpansion.tables['Config'].select("Name = 'InstallPath'")[0].value
$DatabaseName = $global:dsTabExpansion.tables['Config'].select("Name = 'DatabaseName'")[0].value
$DatabasePath = $global:dsTabExpansion.tables['Config'].select("Name = 'DatabasePath'")[0].value

# Load the PowerTab Utility Functions

. "$installDir\TabExpansionLib.ps1" 

# Load TabExpansion database

Import-TabExpansionDataBase $DatabaseName $DatabasePath -nomessage
Import-TabExpansionConfig 'PowerTabConfig.xml' $ConfigurationLocation -no

if ($global:dsTabExpansion.Tables['Config'].select("Name = 'Version'")[0].value -ne $InitScriptVersion) {
  Write-Warning "Error while loading the PowerTab configuration !, version of configuration file not correct !"
  write-Warning "Configuration of $($global:dsTabExpansion.Tables['Config'].select(""Name = 'Version'"")[0].value) found while $InitScriptVersion was expected ! If powertab library files are updated please run PowerTabsetup.ps1 again to update configuration`n"
  write-Warning "`n If you started Setup.cmd to upgrade Powertab this is an expected situation, you can just continue and the powershell setup script will start after the profile is loaded and update the configurationdatabase to the right version`n"
  read-host "press enter to continue"
}
# Backup current tabexpansion function 

&{trap{continue}$global:dsTabExpansion.Tables.Remove('Cache')}
$dtCache = New-Object System.Data.DataTable
[void]$dtCache.Columns.add('Name',[string])
[void]$dtCache.Columns.add('Value')
$dtCache.TableName = 'Cache'
$row = $dtCache.newrow()
$row.Name = 'OldTabexpansion'
$oldTabexpansion = gc function:\tabexpansion
$row.Value = $oldTabexpansion
$dtCache.rows.add($row)
$global:dsTabExpansion.Tables.Add($dtCache)

$global:PowerTabConfig = new-object object

Add-Member -InputObject $Global:PowerTabConfig -MemberType NoteProperty -Name Version -Value $global:dsTabExpansion.Tables['Config'].select("Name = 'Version'")[0].value
# Add enable scriptproperties

    add-member `
      -InputObject $PowerTabConfig `
      -MemberType ScriptProperty `
      -Name Enabled `
      -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
            "`$v = `$dsTabExpansion.Tables['Config'].Select(""Name = 'Enabled'"")[0]
            if (`$v.type -eq 'bool'){[bool][int]`$v.Value}
            else {[$($_.type)](`$v.value)}
         ") `
      -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                "trap{write-warning `$_;continue}
                `$val = [bool]`$args[0]
                 `$val = [int]`$val
                `$dsTabExpansion.Tables['Config'].Select(""Name = 'Enabled'"")[0].Value = `$val
                 if ([bool]`$val){`$path = `$dsTabExpansion.Tables['Config'].Select(""Name = 'InstallPath'"")[0].value
                   . ""`$path\TabExpansion.ps1""
                 }else{sc function:\tabexpansion `$global:dsTabExpansion.Tables['Cache'].select(""name = 'OldTabExpansion'"")[0].value}") `
      -Force

$PowerTabColors = new-object object
Add-Member -InputObject $Global:PowerTabConfig -MemberType NoteProperty -Name Colors -Value $PowerTabColors
Add-Member -InputObject $global:PowerTabConfig.Colors -MemberType ScriptMethod -name ToString -Value {"{PowerTab Color Configuration}"} -Force

$PowerTabShortCuts = new-object object
Add-Member -InputObject $PowerTabShortCuts -MemberType ScriptMethod -name ToString -Value {"{PowerTab Shortcut Characters}"} -Force
Add-Member -InputObject $Global:PowerTabConfig -MemberType NoteProperty -Name ShortcutChars -Value $PowerTabShortcuts

$PowerTabSetup = new-object object
Add-Member -InputObject $PowerTabSetup -MemberType ScriptMethod -name ToString -Value {"{PowerTab Setup Data}"} -Force
Add-Member -InputObject $Global:PowerTabConfig -MemberType NoteProperty -Name Setup -Value $PowerTabSetup

# Make Global properties on Config Object

$global:dsTabExpansion.Tables['Config'].select("Category = 'Global'") | 
  Foreach-Object {
    add-member `
      -InputObject $PowerTabConfig `
      -MemberType ScriptProperty `
      -Name $_.Name `
      -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
            "`$v = `$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0]
            if (`$v.type -eq 'bool'){[bool][int]`$v.Value}
            else {[$($_.type)](`$v.value)}
         ") `
      -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                "trap{write-warning `$_;continue}
                `$val = [$($_.type)]`$args[0]
                 if ( '$($_.type)' -eq 'bool' ) {`$val = [int]`$val}
                `$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0].Value = `$val") `
      -Force
  }




# Make Setup properties on Config Object

$global:dsTabExpansion.Tables['Config'].select("Category = 'Setup'") | 
  Foreach-Object {
    add-member `
      -InputObject $PowerTabConfig.setup `
      -MemberType ScriptProperty `
      -Name $_.Name `
      -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
            "`$v = `$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0]
            if (`$v.type -eq 'bool'){[bool][int]`$v.Value}
            else {[$($_.type)](`$v.value)}
         ") `
      -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                "trap{write-warning `$_;continue}
                `$val = [$($_.type)]`$args[0]
                 if ( '$($_.type)' -eq 'bool' ) {`$val = [int]`$val}
                `$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0].Value = `$val") `
      -Force
  }


# Make Color properties on Config Object

$global:dsTabExpansion.Tables['Config'].select("Category = 'Colors'") | 
  Foreach-Object {
    add-member `
      -InputObject $PowerTabConfig.Colors `
      -MemberType ScriptProperty `
      -Name $_.Name `
      -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
             "`$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0].Value") `
      -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                "trap{write-warning `$_;continue}
                `$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0].Value = [consolecolor]`$args[0]") `
      -Force
  }
  Add-Member `
      -InputObject $PowerTabConfig.Colors `
      -MemberType ScriptMethod `
      -name ExportTheme `
      -Value {$this | gm -MemberType ScriptProperty | select @{name='Name';expression={$_.name}},@{name='Color';expression={$PowerTabConfig.colors."$($_.name)"}}} 
 
  Add-Member `
      -InputObject $PowerTabConfig.Colors `
      -MemberType ScriptMethod `
      -name ImportTheme `
      -Value {$args[0] |% {$PowerTabConfig.Colors."$($_.name)" = $_.Color}} 

# Make Shortcut properties on Config Object

$global:dsTabExpansion.Tables['Config'].select("Category = 'ShortcutChars'") | 
  Foreach-Object {
    add-member `
      -InputObject $PowerTabConfig.ShortcutChars `
      -MemberType ScriptProperty `
      -Name $_.Name `
      -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
             "`$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0].Value") `
      -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                "trap{write-warning `$_;continue}
                `$dsTabExpansion.Tables['Config'].Select(""Name = '$($_.name)'"")[0].Value = `$args[0]") `
      -Force
  }



# load other functions 

if ($PowerTabConfig.Enabled) {
  . "$installDir\TabExpansion.ps1"          # Load Main Tabcompletion function
}
. "$installDir\Out-DataGridView.ps1"      # Used for GUI TabExpansion
. "$installDir\ConsoleLib.ps1"            # Used for RawUi ConsoleList border
. "$installDir\Get-ScriptParameters.ps1"  # Get Parameters of Scripts


# load External Library for Share Enumeration

[void][System.Reflection.Assembly]::LoadFile("$installDir\shares.dll")

if ($PowerTabConfig.ShowBanner) {
Write-Host -f 'Yellow' "$($PowerTabConfig.Version) PowerShell TabExpansion library "
Write-Host -f 'Blue' "/\/\o\/\/ 2007 http://thePowerShellGuy.com"
Write-Host -f 'Yellow' "PowerTab Tabexpansion additions enabled : $($PowerTabConfig.Enabled)"
}


# SIG # Begin signature block
# MIIY+QYJKoZIhvcNAQcCoIIY6jCCGOYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU7yEXOuKrqB/pA3U0u29L1+vH
# VFSgghSrMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
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
# FgQU/RinC+nRm9jUUl32rg//Pvr6DqIwDQYJKoZIhvcNAQEBBQAEggEATD2dJH3w
# X/IcMlsTsVLiJIg7SdkUJvvJpz5op0xeAg+63dE0jDEyCp+BrrpmWMAvnrAa4Lln
# Fvs4o1KlHpUb+KJWI2C4w1e/+cfDAyXD7SPSUtS0YL/4+jZMt2pEdTn2tek7J3wT
# 0XbLJZ5ZDt0L0kJUtWc1qEOhfk58Sal3uFySoISMc8aN6WPResqj2PbhSvIYwBh5
# FrKhoYV9jcAn16UIaYEVweJCYkd4As27aJtBJ6IMJ+4i8KQOIHwNFzBwSk110165
# 6hFjbCyV/FOY5S+VQc1yeXbSNpvYtxZ7hBZRyxlmBqEbCAzy8Drph35UPmSVVyuM
# G3ga+xoNIv/MNqGCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBnMFMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UEAxMiVmVy
# aVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT2bg+9raN
# 7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTIwNjE1MTc1OTU4WjAjBgkqhkiG9w0BCQQxFgQU+9PNMuaNTpPi
# ZsvSewo40BlwzL0wDQYJKoZIhvcNAQEBBQAEgYCFhmpPKoTp2JnuXlgaVQ9p+SXp
# tAl/cbmYkSgbzDiqHD5jN5Mtu+VqSgNivaG+lFX43LXJpcAbxgZHs2qWD48Xbjgh
# afdnfy0Ig2nX1r7F20i94VPymhZTieDLx1Wt+I20PirtLALLjPqUChWZgw3NrtRv
# r4/i/wB2CnU/FXOhVQ==
# SIG # End signature block
