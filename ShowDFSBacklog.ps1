# Description: simplify the DFS backlog command
#
# Author: Tad DeVries
# Email: taddevries@gmail.com
# FileName: ShowDFSBacklog.ps1
#
# Copyright (C) 2010 Tad DeVries
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.
#
# License Info: http://www.gnu.org/licenses/gpl-2.0.html
#
#


function ShowDFSBacklog-Receiving
{
    param( [string]$Site=$(Throw "Please enter a site abbreviation") );

    $sendingmember = $Site + "fs1";
    $rgname = "schools.rcas.org\" + $Site + "\data";

    dfsrdiag.exe backlog /sendingmember:$sendingmember /rgname:$rgname /rfname:data

    <#
    .SYNOPSIS
    Display the backlog of files rcasdfs1 is receiving from its partners.

    .DESCRIPTION
    Display the backlog of files rcasdfs1 is receiving from its partners.

    The command is run as follows.
    dfsrdiag.exe backlog /sendingmember:$sendingmember /rgname:$rgname /rfname:data

    .INPUTS
    Site abbreviation.

    .EXAMPLE
    ShowDFSBacklog-Receiving ch

    .EXAMPLE
    sbr ch

    .LINK
    http://splunk.net
    #>
}

function ShowDFSBacklog-Sending
{
    param( [string]$Site=$(Throw "Please enter a site abbreviation") );

    $receivingmember = $Site + "fs1";
    $rgname = "schools.rcas.org\" + $Site + "\data";

    dfsrdiag.exe backlog /sendingmember:rcasdfs1 /receivingmember:$receivingmember /rgname:$rgname /rfname:data

    <#
    .SYNOPSIS
    Display the backlog of files rcasdfs1 is sending to its partners.

    .DESCRIPTION
    Display the backlog of files rcasdfs1 is sending to its partners.

    The command is run as follows.
    dfsrdiag.exe backlog /sendingmember:rcasdfs1 /receivingmember:$receivingmember /rgname:$rgname /rfname:data

    .INPUTS
    Site abbreviation.

    .EXAMPLE
    ShowDFSBacklog-Sending ch

    .EXAMPLE
    sbs ch

    .LINK
    http://splunk.net
    #>
}

Set-Alias sbr ShowDFSBacklog-Receiving
Set-Alias sbs ShowDFSBacklog-Sending

# SIG # Begin signature block
# MIIY1AYJKoZIhvcNAQcCoIIYxTCCGMECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/WR/Fr/5CSB68NJq0ScpUueK
# CIWgghSGMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
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
# CSqGSIb3DQEJBDEWBBRAHCENdaYPqA72NZ/+rIQdLZ0fdDANBgkqhkiG9w0BAQEF
# AASCAQBQRyBs8NuozvKy9jRZFuexlUyDQTqZ1IqMGzEzAWjq1r4T51mBPXF3BtUe
# UexQdKuLh5XLXSoi03Ykqvf5u3O9yWAtnWYG0mUcZe1jRB8XsHS4a5GlfZTFDHCt
# lqs1j77UmotIYc7fJeVIQH/WzZBb9t9NCBt5wpqJ3fJ5uISE/SIJdKaVD//81rOS
# gnP26Kc67tgltfPoj+q5FkAei/IJOLOqjaLl1fEP00ETR+eTrLEusHV4nUtisjmt
# vJk4HGoxmyqVNjPPWoswM5YFtxXe2yuwtzjO1IJhcglmc6KPXrRLXmQrMt5AribY
# ZZADeeqbpuEvPKzkM+aGfwS2m741oYIBfzCCAXsGCSqGSIb3DQEJBjGCAWwwggFo
# AgEBMGcwUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBAhA4Jdf6
# +GGvnvSQ5ya11lrVMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0xMDEwMjMyMjMyMTJaMCMGCSqGSIb3DQEJBDEW
# BBTf3jthzTmSytJcIM0ag5Qj4+PejTANBgkqhkiG9w0BAQEFAASBgH+HlVG0K5cG
# rxFLZNu67D5XoXLXXo6fRxveoULmDT4P93XShG2mtfg53Dpl/CBdhWbuPED/An23
# Zyn4n4//y+7prVhm3CpkboPxn0yD65way6ydh1IEZUHCEJp+BEkOk5XbPikSHj7b
# Jw75vX/bb+SegOlhSJCRUo3WyXGVAUEW
# SIG # End signature block
