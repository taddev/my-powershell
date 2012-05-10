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
# MIIY1AYJKoZIhvcNAQcCoIIYxTCCGMECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUHqpo7ZmLINsTQLPDLW6Q4C2a
# w4egghSGMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
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
# CSqGSIb3DQEJBDEWBBR+NH/63Ze7dDZWL2+cwfplPQH0gTANBgkqhkiG9w0BAQEF
# AASCAQBP6/DuqS+KHnq34dR+9nzK16lT8eHhCUQsl8eGzi1zJwlrPcQgBqEsWlya
# aEBYgr5O//bth/rVhUEOE0LQEapTLVmMKEMq4PXl2t4adCUP4aMnq8en6uaCeQa3
# 58Tlkcc4LeBHMdI/Sc0GlijxrlFcbP4K+fbF8MAW7S6tj58B2m77UM4NRJRGEhhg
# ptVYPE5Hzio3cghZjry/akGWRRcZM2mTfRrR6xBbsTvxBCceeWxOoRPuNjJO6bSj
# XSxCHBEOxT5q0wAvGYuCB9mHOYFjttN11wv9Q/O42Mc/Ahl1hg2Devw/rkyKmolu
# YVzvAPU1JuXc7RAWrIBmm6YcMkucoYIBfzCCAXsGCSqGSIb3DQEJBjGCAWwwggFo
# AgEBMGcwUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBAhA4Jdf6
# +GGvnvSQ5ya11lrVMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0xMTEwMjgxOTUzMzdaMCMGCSqGSIb3DQEJBDEW
# BBTfHLTcRf5qhBck9H6S7EhoV323TjANBgkqhkiG9w0BAQEFAASBgF9Jm2D6/9XP
# a+Gh419V7sIKEDxIoKhlX4UrqHsiVlEK3Bi4K2ucBtL7b2kPGPWEqn30MwGPFnUB
# XhTNP+DcvNM+8yAPWz5mEtQnrxOili+b54e2xliNz86sNw4TjvJY/rQn82Zbst7A
# afqTXpUj2y+xkHGxrTsZOzcqTC/AVmE4
# SIG # End signature block
