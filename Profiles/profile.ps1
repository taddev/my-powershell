<#=============================================================================
Script Name:     Universal Profile
Created On:      05/22/2009
Author:          Mark E. Schill
File:            profile.ps1
Usage:           PS> . profile.ps1
Version:         1.2
Purpose:         Serves as a PowerShell Profile that can be used on multiple systems
Requirements:    <NONE>
Last Updated:    06/05/2010
History:
         1.2 06/05/2010 - Many Numerous changes.
                 1.1 10/11/2009 - Converted to strictly 2.0 and updated layout
              1.0 05/22/2009 - Initial Revision

         ** Licensed under a Creative Commons Attribution 3.0 License **
=============================================================================#>

<# This must be configured on systems where this script is run
 New-PSDrive -name Scripts -psprovider FileSystem -root <Location of Scripts Folder> -Description "Scripts Folder"
 . Scripts:\Profiles\profile.ps1
#>

# Main Function is just like the C# Main function. I use it to be be able to put functions last.
function Main
{

  # Grab some system Information to be displayed
  $PSVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$PSHome\Powershell.exe").FileVersion
  $IPAddress = @( Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object { $_.DefaultIpGateway } )[0].IPAddress[0]
  $Cert = Get-ChildItem -Path Cert:\CurrentUser\my -CodeSigningCert

    # Set up general aliases
  New-Alias -Name "Edit" -Value "edit-file" -force

  # Add to Module path so I can just do "ipmo Module"
  if ( !($Env:PSModulepath.Contains($(Convert-Path Scripts:\Core\Modules\Manual)) ))
  {    $env:PSMODULEPATH += ";" + $(Convert-Path Scripts:\Core\Modules\Manual) }

  # Import my auto modules.
  Get-ChildItem $(Convert-Path Scripts:\Core\Modules\Auto) | Where-Object {$_.PsIsContainer} | %{ Import-Module $($_.FullName) -Force }

  ## We also want to add our scripts directory to the path
  $ENV:PATH += Get-Item Scripts:\Core\Functions | ? { $_.PsIsContainer } | % {";$($_.FullName)" }
  $ENV:PATH += Get-ChildItem Scripts:\Core\Functions\* | ? { $_.PsIsContainer } | % {";$($_.FullName)" }

  # Machine Based Rules
  switch -regex ( $env:COMPUTERNAME)
  {
    ".+"
    {    }
    "GWT-TECHLT02"
    {      Add-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue }
    "BIG-GHI"
    {    }
  }


  # Host based rules
  switch ( $Host.Name )
  {
    "PowerShellPlus Host"
    {
      $Color_Label = "Cyan"
      $Color_Value_1 = "Green"
      $Color_Value_2 = "Yellow"
      $HostWidth = $Host.UI.RawUI.WindowSize.Width
    }
    'Windows PowerShell ISE Host'
    {
      $Color_Label = "DarkCyan"
      $Color_Value_1 = "Magenta"
      $Color_Value_2 = "DarkGreen"
      $HostWidth = 80

            Import-Module ISEPack 

            $PSISE.options.FontName = "Consolas"

      # watch for changes to the Files collection of the current Tab
      register-objectevent $psise.CurrentPowerShellTab.Files collectionchanged -action {
        # iterate ISEFile objects
        $event.sender | % {
          # set private field which holds default encoding to ASCII
          $_.gettype().getfield("encoding","nonpublic,instance").setvalue($_, [text.encoding]::ascii)
        }
      } | Out-null
    }
    default
    {
      $Color_Label = "Cyan"
      $Color_Value_1 = "Green"
      $Color_Value_2 = "Yellow"
      $HostWidth = $Host.UI.RawUI.WindowSize.Width

      if ( $env:PROCESSOR_ARCHITECTURE -eq 'AMD64')
      {        $NPP = "${env:ProgramFiles(x86)}\Notepad++\Notepad++.exe" }
      else
      {        $NPP = "$env:ProgramFiles\Notepad++\Notepad++.exe" }
      if (Test-Path $NPP) { Set-Alias -Name Edit-File -Value $NPP -Force } 

      # Initialize PowerTab
      #& $(Convert-Path scripts:\Core\Includes\PowerTab\Init-TabExpansion.ps1) -ConfigurationLocation $(Convert-Path scripts:\Core\Includes\PowerTab)
    }

  }

  Record-Session # Start Session Recording
  Clear-Host

    $PreviousColor = $Host.UI.RawUI.ForegroundColor

  # Display relevant information
  Write-Host "ComputerName:`t`t" -ForegroundColor $Color_Label -nonewline
  Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $Color_Value_2
  Write-Host "IP Address:`t`t" -ForeGroundColor $Color_Label -nonewline
  Write-Host $IPAddress -ForeGroundColor $Color_Value_2
  Write-Host "UserName:`t`t" -ForegroundColor $Color_Label -nonewline
  Write-Host "$env:UserDomain\$env:UserName" -ForegroundColor $Color_Value_2
  Write-Host "PowerShell Version:`t" -ForegroundColor $Color_Label -nonewline
  Write-Host $PSVersion -ForegroundColor $Color_Value_2
  Write-Host "Code Signing Cert:`t" -ForegroundColor $Color_Label -nonewline
  Write-Host $Cert.FriendlyName -ForegroundColor $Color_Value_2

  Write-Host "Snapins:`t`t" -ForegroundColor $Color_Label -NoNewline
  $StartingPosition = $Host.UI.RawUI.CursorPosition.X
  Write-Host "".PadRight(30,"-") -ForegroundColor $Color_Label
  Get-PSSnapin | Format-Wide -autosize | Out-String -Width $( $HostWidth -$StartingPosition -1 ) -stream | Where-Object {$_} | %{ Write-Host $($(" "*$StartingPosition) + $_) -foregroundColor $Color_Value_1} 

  Write-Host "Modules:`t`t" -foregroundcolor $Color_Label -noNewLine
  $StartingPosition = $Host.UI.RawUI.CursorPosition.X
  Write-Host "".PadRight(30,"-") -ForegroundColor $Color_Label

  Get-Module | Format-Wide -AutoSize | Out-String -Width $( $HostWidth -$StartingPosition -1 ) -stream | Where-Object {$_} |  %{ Write-Host $($(" "*$StartingPosition) + $_) -foregroundColor $Color_Value_1}
  Get-Module -ListAvailable | Format-Wide -Column 3 | Out-String -Width $( $HostWidth -$StartingPosition -1 ) -stream | Where-Object {$_} |  %{ Write-Host $($(" "*$StartingPosition) + $_) -foregroundColor $Color_Value_2} 

  Write-Host "Functions:`t`t" -foregroundcolor $Color_Label -noNewLine
  $StartingPosition = $Host.UI.RawUI.CursorPosition.X
  Write-Host "".PadRight(30,"-") -ForegroundColor $Color_Label
  Get-ChildItem Scripts:\Core\Functions\* -Recurse | Select-Object Name | Format-Wide -AutoSize | Out-String -Width $( $HostWidth -$StartingPosition -1 ) -stream | Where-Object {$_} |  %{ Write-Host $($(" "*$StartingPosition) + $_) -foregroundColor $Color_Value_1}  

  $Host.UI.RawUI.ForegroundColor =$PreviousColor
  Write-Host ""
  Write-Host ""

  # This should go OUTSIDE the prompt function, it doesn't need re-evaluation
    # We're going to calculate a prefix for the window title
  if(!$global:WindowTitlePrefix) {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
    {
      $global:WindowTitlePrefix = "PoSh (ADMIN) -"
    } else {
      $global:WindowTitlePrefix = "PoSh - "
    }
  }

  $LogicalDisk = @()
  gwmi Win32_LogicalDisk -filter "DriveType='3'" | % {
    $LogicalDisk += @($_ | Select @{n="Name";e={$_.Caption}},
    @{n="Volume Label";e={$_.VolumeName}},
    @{n="Used (GB)";e={"{0:N2}" -f ( ($_.Size/1GB) - ($_.FreeSpace/1GB) )}},
    @{n="Free (GB)";e={"{0:N2}" -f ($_.FreeSpace/1GB)}},
    @{n="Size (GB)";e={"{0:N2}" -f ($_.Size/1GB)}},
    @{n="Free (%)";e={if($_.Size) { "{0:N2}" -f ( ($_.FreeSpace/1GB) / ($_.Size/1GB) * 100 )}else{"NAN"} }} )
  } 

  $Host.UI.RawUI.ForegroundColor = $Color_Value_2
  $LogicalDisk | format-table -AutoSize | out-string
  $Host.UI.RawUI.ForegroundColor = $PreviousColor

  Get-SystemUptime
  Write-Host

  cd scripts:\

}

#Record all Powershell activities
function Record-Session
 {
  $MyPath = "$((Get-PSDrive Scripts).Root)\_Transcripts"
  if ( ! (Test-Path $MyPath ) ) { mkdir $MyPath > $null }
  $ComputerName = $env:ComputerName
  $Date = Get-Date -Format "yyyy-MM-dd"
  switch ( $Host.Name )
  {
    "PowerShellPlus Host"
    {      Start-Transcript -path "$MyPath\$ComputerName-$Date.log" -ea silentlycontinue }
    'Windows PowerShell ISE Host'
        {
            # PowerShell ISE does not support transcription
        }
        default
    {      Start-Transcript -path "$MyPath\$ComputerName-$Date.log" -append -ea silentlycontinue }

  }
}

function prompt
{

  Write-Host "$([char]0x0A7) " -NoNewline -ForegroundColor $Color_Label
  Write-Host ([net.Dns]::GetHostName()) -NoNewline -ForegroundColor $Color_Value_1
  Write-Host ' {' -NoNewline -ForegroundColor "Red"
  Write-Host (shorten-path (pwd).path) -NoNewline -ForegroundColor $Color_Label
  Write-Host '}' -NoNewline -ForegroundColor "Red"
  return ' '
}


function shorten-path([string] $path) {
  $loc = $path.Replace($HOME, '~')
  # remove prefix for UNC paths
  $loc = $loc -replace '^[^:]+::', ''
  # make path shorter like tabs in Vim,
    # handle paths starting with \\ and . correctly
  return ($loc -replace '\\(\.?)([^\\]{3})[^\\]*(?=\\)','\$1$2')
} 

function Get-SystemUptime ($computer = "$env:computername") {
  $lastboot = [System.Management.ManagementDateTimeconverter]::ToDateTime("$((gwmi  Win32_OperatingSystem -computername $computer).LastBootUpTime)")
  $uptime = (Get-Date) - $lastboot
  Write-Host "System Uptime for $computer is: " -NoNewline -ForegroundColor $Color_Value_2
  Write-Host $uptime.days -NoNewline -ForegroundColor $Color_Label
  Write-Host " days " -NoNewline -ForegroundColor $Color_Value_2
  Write-Host $uptime.hours -NoNewline -ForegroundColor $Color_Label
  Write-Host " hours " -NoNewline -ForegroundColor $Color_Value_2
  Write-Host $uptime.minutes -NoNewline -ForegroundColor $Color_Label
  Write-Host " minutes " -NoNewline -ForegroundColor $Color_Value_2
  Write-Host $uptime.seconds -NoNewline -ForegroundColor $Color_Label
  Write-Host " seconds" -ForegroundColor $Color_Value_2
}

# Call "Main" Function
. Main


# SIG # Begin signature block
# MIIY+QYJKoZIhvcNAQcCoIIY6jCCGOYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU9X9+eDRgLey6/MG9yipXK7K/
# Cr+gghSrMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
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
# FgQUeabJSR7grEqKeNjXArPdGossKcIwDQYJKoZIhvcNAQEBBQAEggEABDhRDD45
# IgpRfPmbg4VZKwXOaR2hWuu8prntTgKvkz/r05PXLWBpV6umx3OFK/0InmO3BpvG
# FN9ytfqzRQgR1DHlpNxIU8RZ7sQQjL95igIjfmOKPD+hpqerqwN8XzXBFVCpTh99
# 8PcyBO77fY9otAcEaU4I72oQv/vH2FHOxIB8C6iQf68IJa9PD9AI9CRA6ly/H2Xb
# pXJ65IZ3dLsvn/SEapbkHL5QIvZ9ei8P7sfWO7ygjUSsfBallr0rd5QgpxUC9TA1
# rPxMPUzKQnZglpqFs56tqKHbyR4g519ju9bIP/TTBK2aqIxPww+xCVlsY8Eepr3V
# /OWRt8Oa/y7aN6GCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBnMFMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UEAxMiVmVy
# aVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT2bg+9raN
# 7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTIwNjE1MTc1NjIxWjAjBgkqhkiG9w0BCQQxFgQUsMs4uApqf8s3
# A1m86wviS2blseYwDQYJKoZIhvcNAQEBBQAEgYCbdGpDWA/y+KKsjhNmwqbudcwJ
# 9qRc7pRxjO7j06pGUdNaNcP7EG+tv/PvjPeYjuX3cOsPSSoDOy7LV+0ro+dkI0fg
# p3OyySBNuTBaPFRNhvcAf2QOGTO/AitMJJcIvfa2tw9rBaZPbu/XKQM2wf/bk+my
# AjRkk7EoUwwJRSJjxA==
# SIG # End signature block
