# TabExpansionUtil.ps1
#
# 


#########################
## Private functions
#########################

Function Out-DataGridView {
    [CmdletBinding()]
    param(
		[Parameter(Position = 0)]
        [String]
        $ReturnField
        ,
		[Parameter(ValueFromPipeline = $true)]
        [Object[]]
        $InputObject
    )

    begin {
        [Object[]]$Objects = @()
    }

    process {
        $Objects += $InputObject
    }

    end {
        # Make DataTable from Input
        $dt = New-Object System.Data.DataTable
        $First = $true
        foreach ($Item in $Objects) {
            $dr = $dt.NewRow()
            $Item.PSObject.get_Properties() | ForEach-Object {
                if ($first) {
                    $col =  New-Object System.Data.DataColumn
                    $col.ColumnName = $_.Name.ToString()
                    $dt.Columns.Add($col)
                }
                if ($_.Value -eq $null) {
                    $dr.Item($_.Name) = "[empty]"
                } elseif ($_.IsArray) {
                    $dr.Item($_.Name) =[String]::Join($_.Value ,";")
                } else {
                    $dr.Item($_.Name) = $_.Value
                }
            }
            $dt.Rows.Add($dr)
            $First = $false
        }

        # Show Datatable in Form
        $form = New-Object System.Windows.Forms.Form
        $form.Size = new-Object System.Drawing.Size @(1000,600)
        $dg = New-Object System.Windows.Forms.DataGridView
        $dg.DataSource = $dt.PSObject.BaseObject
        $dg.Dock = [System.Windows.Forms.DockStyle]::Fill
        $dg.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
        $dg.SelectionMode = 'FullRowSelect'
        $dg.add_DoubleClick({
            $script:ret = $this.SelectedRows | ForEach-Object {$_.DataBoundItem["$ReturnField"]}
            $form.Close()
        })

        $form.Text = "$($MyInvocation.Line)"
        $form.KeyPreview = $true
        $form.add_KeyDown({
            if ($_.KeyCode -eq 'Enter') {
                $script:ret = $dg.SelectedRows | ForEach-Object {$_.DataBoundItem["$ReturnField"]}
                $form.Close()
            } elseif ($_.KeyCode -eq 'Escape') {
                $form.Close()
            }
        })

        $form.Controls.Add($dg)
        $form.add_Shown({$form.Activate(); $dg.AutoResizeColumns()})
        $script:ret = $null
        [Void]$form.ShowDialog()
        $script:ret
    }
}

############

Function Resolve-Command {
    [CmdletBinding()]
    param(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name
        ,
        [Switch]
        $CommandInfo
    )

    process {
        $Command = ""

        ## Get command info, the where clause prevents problems with "?" wildcard
        if ($Name -match "\\") {
            ## Full name usage
            $Module = $Name.Substring(0, $Name.Indexof("\"))
            $CommandName = $Name.Substring($Name.Indexof("\") + 1, $Name.length - ($Name.Indexof("\") + 1))
            if ($Module = Get-Module $Module) {
                $Command = @(Get-Command $CommandName -Module $Module -ErrorAction SilentlyContinue)[0]
                if (-not $Command) {
                    ## Try to look up command with prefix
                    $Prefix = Get-CommandPrefix $Module
                    $Verb = $CommandName.Substring(0, $CommandName.Indexof("-"))
                    $Noun = $CommandName.Substring($CommandName.Indexof("-") + 1, $CommandName.length - ($CommandName.Indexof("-") + 1))
                    $Command = @(Get-Command "$Verb-$Prefix$Noun" -ErrorAction SilentlyContinue)[0]
                }
                if (-not $Command) {
                    ## Try looking in the module's exported command list
                    $Command = $Module.ExportedCommands[$CommandName]
                }
            }
        }
        if (-not $Command) {
            if ($Name.Contains("?")) {
                $Command = @(Get-Command $Name | Where-Object {$_.Name -eq $Name})[0]
            } else {
                $Command = @(Get-Command $Name)[0]
            }
        }

        if ($Command.CommandType -eq "Alias") {
            $Command = $Command.ResolvedCommand	
        }

        ## Return result
        if ($CommandInfo) {
            $Command
        } else {
            if ($Command.CommandType -eq "ExternalScript") {
                $Command.Path
            } else {
                $Command.Name
            }
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}

Function Resolve-Parameter {
    [CmdletBinding(DefaultParameterSetName = "Command")]
    param(
		[Parameter(ParameterSetName = "Command", Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Command
        ,
		[Parameter(ParameterSetName = "CommandInfo", Mandatory = $true, Position = 0)]
        [ValidateNotNull()]
        [System.Management.Automation.CommandInfo]
        $CommandInfo
        ,
		[Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name
        ,
        [Switch]
        $ParameterInfo
    )

    process {
        ## Remove leading dash if it exists
        $Name = $Name -replace '^-'

        ## Get command info
		if ($PSCmdlet.ParameterSetName -eq "Command") {
            $CommandInfo = Resolve-Command $Command -CommandInfo
        } elseif ($PSCmdlet.ParameterSetName -eq "CommandInfo") {
            if ($CommandInfo -eq $null) {return}
        }

        ## Check if this is a real parameter name and not an alias
        if ($CommandInfo.Parameters["$Name"]) {
            $Parameter = $CommandInfo.Parameters["$Name"]
        } else {
            ## Possible alias
            $Parameter = @($CommandInfo.Parameters.Values | Where-Object {$_.Aliases -contains $Name})[0]
        }

        ## If no parameter found, it could be an abreviated name (-comp instead of -ComputerName)
        if (-not $Parameter) {
            $Parameter = @($CommandInfo.Parameters.Values | Where-Object {$_.Name -like "$Name*"})[0]
        }

        ## Return result
        if ($ParameterInfo) {
            $Parameter
        } else {
            $Parameter.Name
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}

Function Resolve-PositionalParameter {
    param(
		[Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNull()]
        [Object]
        $Context
    )
    
    process {
        if ($TabExpansionCommandInfoRegistry[$Context.Command]) {
            $ScriptBlock = $TabExpansionCommandInfoRegistry[$Context.Command]
            $CommandInfo = & $ScriptBlock $Context
            if (-not $CommandInfo) {throw "foo"} ## TODO
        } elseif ($Context.CommandInfo) {
            $CommandInfo = $Context.CommandInfo
        } else {
            return $Context
        }

        foreach ($ParameterSet in $CommandInfo.ParameterSets) {
            $PositionalParameters = @($ParameterSet.Parameters |
                Where-Object {($_.Position -ge 0) -and ($Context.OtherParameters.Keys -notcontains $_.Name)} | Sort-Object Position)

            if (($Context.PositionalParameter -ge 0) -and ($Context.PositionalParameter -lt $PositionalParameters.Count)) {
                ## TODO: Try to figure out a better parameter?
                $Context.Parameter = $PositionalParameters[$Context.PositionalParameter].Name
                #$Context.PositionalParameter -= 1
                break
            } elseif ($PositionalParameters[-1].ValueFromRemainingArguments) {
                $Context.Parameter = $PositionalParameters[-1].Name
                break
            }
        }

        $Context

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}

Function Resolve-InternalCommandName {
    [CmdletBinding(DefaultParameterSetName = "Command")]
    param(
		[Parameter(ParameterSetName = "Command", Mandatory = $true, Position = 0, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Command
        ,
		[Parameter(ParameterSetName = "CommandInfo", Mandatory = $true, Position = 0)]
        [ValidateNotNull()]
        [System.Management.Automation.CommandInfo]
        $CommandInfo
    )

    process {
        ## Get command info
		if ($PSCmdlet.ParameterSetName -eq "Command") {
            $CommandInfo = Resolve-Command $Command -CommandInfo
        }

        ## Return result
        if ($Prefix = Get-CommandPrefix $CommandInfo) {
            $Verb = $CommandInfo.Name.Substring(0, $CommandInfo.Name.Indexof("-"))
            $Noun = $CommandInfo.Name.Substring($CommandInfo.Name.Indexof("-") + 1, $CommandInfo.Name.length - ($CommandInfo.Name.Indexof("-") + 1))
            $Noun = $Noun -replace [Regex]::Escape($Prefix)
            $InternalName = "$Verb-$Noun"
        } else {
            $InternalName = $CommandInfo.Name
        }

        New-Object PSObject -Property @{"InternalName"=$InternalName;"Module"=$CommandInfo.Module}

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}

Function Get-CommandPrefix {
    [CmdletBinding(DefaultParameterSetName = "Command")]
    param(
		[Parameter(ParameterSetName = "Command", Mandatory = $true, Position = 0)]
        [ValidateNotNull()]
        [String]
        $Command
        ,
		[Parameter(ParameterSetName = "CommandInfo", Mandatory = $true, Position = 0)]
        [ValidateNotNull()]
        [System.Management.Automation.CommandInfo]
        $CommandInfo
        ,
		[Parameter(ParameterSetName = "ModuleInfo", Mandatory = $true, Position = 0)]
        [ValidateNotNull()]
        [System.Management.Automation.PSModuleInfo]
        $ModuleInfo
    )

    process {
        ## Get module info
		if ($PSCmdlet.ParameterSetName -eq "Command") {
            $ModuleInfo =  (Resolve-Command $Command -CommandInfo).Module
        } elseif (($PSCmdlet.ParameterSetName -eq "CommandInfo") -and $CommandInfo.Module) {
            $ModuleInfo =  Get-Module $CommandInfo.Module
        }

        if ($ModuleInfo) {
            $CommandGroups = $ModuleInfo.ExportedFunctions.Values +
                (Get-Command -Module $ModuleInfo -CommandType Function,Filter,Cmdlet) | Group-Object {$_.Definition}
            $Prefixes = foreach ($Group in $CommandGroups) {
                $Names = $Group.Group | Select-Object -ExpandProperty Name
                $TempNoun = (@($Names)[0] -split "-")[1]
            	foreach($Name in $Names) {
            		if ($Name -match "-") {
            			$PossiblePrefix = $Name.SubString($Name.IndexOf("-") + 1, $Name.LastIndexOf($TempNoun) - $Name.IndexOf("-") - 1)
                        if ($PossiblePrefix) {
                            $PossiblePrefix
                        }
            		}
            	}
            }

            if ($Prefixes.Count) {
                $Prefixes | Select-Object -Unique
            } else {
                $Prefixes
            }
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}

############

Function Resolve-TabExpansionParameterValue {
    param(
        [String]$Value
    )

    switch -regex ($Value) {
        '^\$' {
            [String](Invoke-Expression $_)
            break
        }
        '^\(.*\)$' {
            [String](Invoke-Expression $_)
            break
        }
        Default {$Value}
    }
}

############

## Slightly modified from http://blog.sapien.com/index.php/2009/08/24/writing-form-centered-scripts-with-primalforms/
Function Get-GuiDate {
    param(
       [Int]$DisplayMode = 1, # number of months to show
       [Int]$SelectionCount = 0, # number of days that can be selected
       [DateTime]$TodayDate = $(Get-Date), # sets default selected date
       [DateTime]$DateSelected = $TodayDate, # sets default selected date
       [Int]$FirstDayofWeek = -1, # -1 used default - calendar dayofweek, NOT datetime
       [DateTime[]]$Bold = @(), # Array of bolded dates to add
       [DateTime[]]$YBold = @(), # annual bolded dates to add
       [DateTime[]]$MBold = @(), # monthly bolded dates to add
       [Int]$ScrollBy = $DisplayMode, # number of months to scroll by; 0 = screenfull
       [Switch]$WeekNumbers, # Show numeric week of year on the display
       [String]$Title = "Get-GuiDate",
       [Switch]$NoTodayCircle,
       [DateTime]$MinDate = "1753-01-01",
       [DateTime]$MaxDate = "9998-12-31"
    )

    [System.Windows.Forms.Application]::EnableVisualStyles()
    # Is this voodoo code, or not?
    [System.Windows.Forms.Application]::DoEvents()

    $cal = New-Object Windows.Forms.MonthCalendar
    $cal.SetDate($DateSelected)
    $cal.TodayDate = $TodayDate
    if ($SelectionCount -lt 1) {$SelectionCount = [int]::MaxValue}
    $cal.MaxSelectionCount = $SelectionCount
    $cal.MinDate = $MinDate
    $cal.MaxDate = $MaxDate
    $cal.ScrollChange = $ScrollBy
    $cal.ShowTodayCircle = $true
    if ($FirstDayofWeek -eq -1) {$FirstDayofWeek = [System.Windows.Forms.Day]::Default}
    $cal.FirstDayofWeek = [System.Windows.Forms.Day]$FirstDayofWeek
    $cal.ShowWeekNumbers = $WeekNumbers
    if ($NoTodayCircle) {$cal.ShowTodayCircle = $False}

    # Provides clean display geometry
    switch -regex ($DisplayMode) {
        "^1$" {$cal.CalendarDimensions = "1,1"}
        "^2$" {$cal.CalendarDimensions = "2,1"}
        "^3$" { $cal.CalendarDimensions = "3,1"}
        "^4$" {$cal.CalendarDimensions = "2,2"}
        "^[56]$" {$cal.CalendarDimensions = "3,2"}
        "^[78]$" {$cal.CalendarDimensions = "4,2"}
        "^9$" {$cal.CalendarDimensions = "3,3"}
        "^1[012]$" {$cal.CalendarDimensions = "4,3"}
        default {$cal.CalendarDimensions = "4,4"}
    }

    if ($Bold) {$cal.BoldedDates = $Bold}
    if ($YBold) {$cal.AnnuallyBoldedDates = $YBold}
    if ($MBold) {$cal.MonthlyBoldedDates = $MBold}

    $form = New-Object Windows.Forms.Form
    $form.AutoSize = $form.TopMost = $form.KeyPreview = $True
    $form.MaximizeBox = $form.MinimizeBox = $False
    $form.AutoSizeMode = "GrowAndShrink"
    $form.Controls.Add($cal)
    $form.BackColor = [System.Drawing.Color]::White
    $form.Text = $Title

    # We'll handle escape or enter to get out.
    $Escaped = $False;
    $form.Add_KeyDown([System.Windows.Forms.KeyEventHandler]{
        if ($_.KeyCode -eq "Escape") {
            $Escaped = $true; $form.Close()
        } elseif ($_.KeyCode -eq "Enter") {
            $form.Close()
        }
    })

    # Ensures the form is on top, is active, and then shows it.
    # After calling ShowDialog(), the script is blocked until
    # the form is no longer visible.
    $form.Add_Shown({$form.Activate()}) 
    [Void]$form.ShowDialog()

    # If they didn't press Escape, output the selection range
    # as a series of dates.
    if (!$Escaped) {
        for(
            $day = $cal.SelectionRange.Start;
            $day -le $cal.SelectionRange.End;
            $day = $day.AddDays(1)
            )
        {
            $day
        }
    }

    # 2009-08-27
    # -initialized $Escaped and removed $ShowTodayCircle (thanks, tojo2000) 
    # -modified $FirstDayOfWeek so casts don't occur until after Forms library loaded.
}

Function Test-IsolatedStoragePath {
    [CmdletBinding()]
    param(
        [Alias("LiteralPath")]
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Path
    )

    process {
        try {
            $UserIsoStorage = [System.IO.IsolatedStorage.IsolatedStorageFile]::GetUserStoreForAssembly()
            if ($UserIsoStorage.GetFileNames($Path)) {
                $true
            } else {
                $false
            }
        } catch {
            $false
        }
    }
}

Function Open-IsolatedStorageFile {
    [CmdletBinding()]
    param(
        [Alias("Path")]
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath
        ,
        [Switch]
        $Writable
    )

    process {
        if ($Writable) {
            $UserIsoStorage = [System.IO.IsolatedStorage.IsolatedStorageFile]::GetUserStoreForAssembly()
            if (Test-IsolatedStoragePath $LiteralPath) {
                New-Object System.IO.IsolatedStorage.IsolatedStorageFileStream($LiteralPath, [System.IO.FileMode]::Truncate, $UserIsoStorage)
            } else {
                New-Object System.IO.IsolatedStorage.IsolatedStorageFileStream($LiteralPath, [System.IO.FileMode]::Create, $UserIsoStorage)
            }
        } else {
            $UserIsoStorage = [System.IO.IsolatedStorage.IsolatedStorageFile]::GetUserStoreForAssembly()
            New-Object System.IO.IsolatedStorage.IsolatedStorageFileStream($LiteralPath, [System.IO.FileMode]::Open, $UserIsoStorage)
        }
    }
}

Function New-IsolatedStorageDirectory {
    [CmdletBinding()]
    param(
        [Alias("Path")]
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath
    )

    process {
        $UserIsoStorage = [System.IO.IsolatedStorage.IsolatedStorageFile]::GetUserStoreForAssembly()
        if (-not $UserIsoStorage.GetDirectoryNames($LiteralPath)) {$UserIsoStorage.CreateDirectory($LiteralPath)}
    }
}

Function Get-IsolatedStorage {
}


##########
# Here there be hacks (from Jaykul)
##########

Function Parse-Manifest {
    $Manifest = Get-Content "$PSScriptRoot\PowerTab.psd1" | Where-Object {$_ -notmatch '^\s*#'}
    $ModuleManifest = "Data {`n" + ($Manifest -join "`r`n") + "`n}"
    $ExecutionContext.SessionState.InvokeCommand.NewScriptBlock($ModuleManifest).Invoke()[0]
}

Function Find-Module {
    [CmdletBinding()]
    param(
        [String[]]$Name = "*"
        ,
        [Switch]$All
    )

    foreach ($n in $Name) {
        $folder = [System.IO.Path]::GetDirectoryName($n)
        $n = [System.IO.Path]::GetFileName($n)
        $ModulePaths = Get-ModulePath

        if ($folder) {
            $ModulePaths = Join-Path $ModulePaths $folder
        }

        ## Note: the order of these is important. They need to be in the order they'd be loaded by the system
        $Files = @(Get-ChildItem -Path $ModulePaths -Recurse -Filter "$n.ps?1" -EA 0; Get-ChildItem -Path $ModulePaths -Recurse -Filter "$n.dll" -EA 0)
        $Files | Where-Object {
                $parent = [System.IO.Path]::GetFileName( $_.PSParentPath )
                return $all -or ($parent -eq $_.BaseName) -or ($folder -and ($parent -eq ([System.IO.Path]::GetFileName($folder))) -and ($n -eq $_.BaseName))
            } | Group-Object PSParentPath | ForEach-Object {@($_.Group)[0]}
    }
}

# | Sort-Object {switch ($_.Extension) {".psd1"{1} ".psm1"{2}}})
Function Get-ModulePath {
    $Env:PSModulePath -split ";" | ForEach-Object {"{0}\" -f $_.TrimEnd('\','/')} | Select-Object -Unique | Where-Object {Test-Path $_}
}
# SIG # Begin signature block
# MIIbaQYJKoZIhvcNAQcCoIIbWjCCG1YCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHqpo7ZmLINsTQLPDLW6Q4C2a
# w4egghYbMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
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
# kXcMCwowggg4MIIHIKADAgECAgIHqTANBgkqhkiG9w0BAQUFADCBjDELMAkGA1UE
# BhMCSUwxFjAUBgNVBAoTDVN0YXJ0Q29tIEx0ZC4xKzApBgNVBAsTIlNlY3VyZSBE
# aWdpdGFsIENlcnRpZmljYXRlIFNpZ25pbmcxODA2BgNVBAMTL1N0YXJ0Q29tIENs
# YXNzIDIgUHJpbWFyeSBJbnRlcm1lZGlhdGUgT2JqZWN0IENBMB4XDTEyMTAxMDIz
# MzU0OFoXDTE0MTAxMzAwMjE0MFowgY8xGTAXBgNVBA0TEFMzRUM3Yzh4Y1lOMnBQ
# cXUxCzAJBgNVBAYTAlVTMRUwEwYDVQQIEwxTb3V0aCBEYWtvdGExEzARBgNVBAcT
# ClJhcGlkIENpdHkxFDASBgNVBAMTC1RhZCBEZVZyaWVzMSMwIQYJKoZIhvcNAQkB
# FhR0YWRkZXZyaWVzQGdtYWlsLmNvbTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAKb2chsYUh+l9MhIyQc+TczABVRO4rU3YOwu1t0gybek1d0KacGTtD/C
# SFWutUsrfVHWb2ybUiaTN/+P1ChqtnS4Sq/pyZ/UcBzOUoFEFlIOv5NxTjv7gm2M
# pR6LwgYx2AyfdVYpAfcbmAH0wXfgvA3i6y9PEAlVEHq3gf11Hf1qrQKKD+k7ZMHG
# ozQhmtQ9MxfF4VCG9NNSU/j7TXJG+j7sxlG0ADxwjMo+iA7R1ANs6N2seOnvcNvQ
# a3YP4SwHv0hUgz9KBXHXCdA7LG8lGlLp4s0bbyPxagZ1+Of0qnTyG4yq5qij8Wsa
# xAasi1sRYM6rO6Dn5ISaIF1lJmQIOYPezivKenDc3o9yjbb4jPDUjT7M2iK+VRfc
# FPEbcxHJ+FpUAvTYPOEeDO2LkriuRvUkkMTYiXWpqUVojLk3JDlcCRkE5cykIMdX
# irx82lxQpiZGkFrfrGQPMi6DAALX85ZUiDQ10iGyXANtubJkhAnp5hn4Q5JA4tpR
# ty6MlZh94TjeFlbXq9Y2phRi3AWqunOMAxX8gSHfbrmAa7gNkaBoVZd2tlVrV1X+
# lnnnb3yO0SuErx3bfhS++MgrisERscGgcY+vB5trw05FMGfK5YkzWZF2eIE/m70T
# 2rfmH9tUnElgJHTqEu4L8txmnNZ/j8ZzyLNY5+n8XqGghtTqeIxLAgMBAAGjggOd
# MIIDmTAJBgNVHRMEAjAAMA4GA1UdDwEB/wQEAwIHgDAuBgNVHSUBAf8EJDAiBggr
# BgEFBQcDAwYKKwYBBAGCNwIBFQYKKwYBBAGCNwoDDTAdBgNVHQ4EFgQU/zkKtNmi
# KcWBOqQkxr6qsIyjrGUwHwYDVR0jBBgwFoAU0E4PQJlsuEsZbzsouODjiAc0qrcw
# ggIhBgNVHSAEggIYMIICFDCCAhAGCysGAQQBgbU3AQICMIIB/zAuBggrBgEFBQcC
# ARYiaHR0cDovL3d3dy5zdGFydHNzbC5jb20vcG9saWN5LnBkZjA0BggrBgEFBQcC
# ARYoaHR0cDovL3d3dy5zdGFydHNzbC5jb20vaW50ZXJtZWRpYXRlLnBkZjCB9wYI
# KwYBBQUHAgIwgeowJxYgU3RhcnRDb20gQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkw
# AwIBARqBvlRoaXMgY2VydGlmaWNhdGUgd2FzIGlzc3VlZCBhY2NvcmRpbmcgdG8g
# dGhlIENsYXNzIDIgVmFsaWRhdGlvbiByZXF1aXJlbWVudHMgb2YgdGhlIFN0YXJ0
# Q29tIENBIHBvbGljeSwgcmVsaWFuY2Ugb25seSBmb3IgdGhlIGludGVuZGVkIHB1
# cnBvc2UgaW4gY29tcGxpYW5jZSBvZiB0aGUgcmVseWluZyBwYXJ0eSBvYmxpZ2F0
# aW9ucy4wgZwGCCsGAQUFBwICMIGPMCcWIFN0YXJ0Q29tIENlcnRpZmljYXRpb24g
# QXV0aG9yaXR5MAMCAQIaZExpYWJpbGl0eSBhbmQgd2FycmFudGllcyBhcmUgbGlt
# aXRlZCEgU2VlIHNlY3Rpb24gIkxlZ2FsIGFuZCBMaW1pdGF0aW9ucyIgb2YgdGhl
# IFN0YXJ0Q29tIENBIHBvbGljeS4wNgYDVR0fBC8wLTAroCmgJ4YlaHR0cDovL2Ny
# bC5zdGFydHNzbC5jb20vY3J0YzItY3JsLmNybDCBiQYIKwYBBQUHAQEEfTB7MDcG
# CCsGAQUFBzABhitodHRwOi8vb2NzcC5zdGFydHNzbC5jb20vc3ViL2NsYXNzMi9j
# b2RlL2NhMEAGCCsGAQUFBzAChjRodHRwOi8vYWlhLnN0YXJ0c3NsLmNvbS9jZXJ0
# cy9zdWIuY2xhc3MyLmNvZGUuY2EuY3J0MCMGA1UdEgQcMBqGGGh0dHA6Ly93d3cu
# c3RhcnRzc2wuY29tLzANBgkqhkiG9w0BAQUFAAOCAQEAMDdkGhWaFooFqzWBaA/R
# rf9KAQOeFSLoJrgZ+Qua9vNHrWq0TGyzH4hCJSY4Owurl2HCI98R/1RNYDWhQ0+1
# dK6HZ/OmKk7gsbQ5rqRnRqMT8b2HW7RVTVrJzOOj/QdI+sNKI5oSmTS4YN4LRmvP
# MWGwbPX7Poo/QtTJAlxXkeEsLN71fabQsavjjJORaDXDqgd6LydG7yJOlLzs2zDr
# dSBOZnP8VD9seRIZtMWqZH2tGZp3YBQSTWq4BySHdsxsIgZVZnWi1HzSjUTMtbcl
# P/CKtZKBCS7FPHJNcACouOQbA81aOjduUtIVsOnulVGT/i72Grs607e5m+Z1f4pU
# FjGCBLgwggS0AgEBMIGTMIGMMQswCQYDVQQGEwJJTDEWMBQGA1UEChMNU3RhcnRD
# b20gTHRkLjErMCkGA1UECxMiU2VjdXJlIERpZ2l0YWwgQ2VydGlmaWNhdGUgU2ln
# bmluZzE4MDYGA1UEAxMvU3RhcnRDb20gQ2xhc3MgMiBQcmltYXJ5IEludGVybWVk
# aWF0ZSBPYmplY3QgQ0ECAgepMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBR+NH/63Ze7dDZWL2+c
# wfplPQH0gTANBgkqhkiG9w0BAQEFAASCAgCR68Uc2+nIXD8aZF6Nf0sezURM48Xh
# WrSQrkfg/BshCwQs1Y/1PEOETRYW1ZnWV1sphSc2v2proQtID4LvaRudy6CZt07l
# wkb1wVflo9D4n8A3fJfmlaU9hqwatR+stDM0Lwf6XaIqtz3sijM0GX0/vn1xlQF7
# 0uCIZW3RjWkeAomauh2L5vueeJpc0XhfMc6i2sNDFcqTUrjrA6Q27h1aoNvzcwPR
# +E2xbGWjGL8UT3nYPF+R8oNzdDRb3H8qBuZ1weCIab0D7rJ/lkhDPYJVAkO/da6Y
# IEL4lNXqUj2EDECIAl6KBnFwwdmMTAQtTXL3NLIkscCD0XE9XURtj1TDY1mNN+Uv
# Th+LRhjhT8+evGVjP6JpXiKnO57St02VVSDjDI+sb8m1dpS2e2hZfCqQTY6ukrEE
# t9IW4lFUqOI/DwfknVXjfsPJMpb2eav1N5O/rEmNripllnc+C3yZxkeH2dnHP1En
# 69oqFE6DXh8iB1DZ63HPgXFsJSOuVk8qvbze3ImV2e92xQ+EmPb0kvlPfsDLVThW
# OSTi2+hiomGqnYBTB4+zFKtOJVFMQdQU6DA9Xhq18LKaJcV3hwaO3uS79ToUyAKq
# 4MytB/Z0wmNhSMXNr3+3e6YZt2dDK/hRw6B5nXkPSIa0EEE6MUDEkqUJSW3Zh9/W
# WMqDTB0Y8y9xhqGCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBnMFMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UEAxMiVmVy
# aVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT2bg+9raN
# 7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTIxMDI3MDI0NDUxWjAjBgkqhkiG9w0BCQQxFgQUYGVR+z7LSQ5B
# xuQpkLbH50EapIQwDQYJKoZIhvcNAQEBBQAEgYBtYrjPRsiuivAgfo8azF7734G6
# 1dn/CiJ/tMcD6P+gGSRsb7bi/mgEB6eGi6+vj0w2MDjzcFbRPGTFAN38/YTiMWyO
# xFQr0dE4Jy88UESlqCOr8+evs8HUkCGl3O8bnZeWZybKdJc9ex53U2FwQ+Q2PmJk
# BTm/zN5j4hDszzUADg==
# SIG # End signature block
