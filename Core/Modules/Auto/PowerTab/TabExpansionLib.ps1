# TabExpansionLib.ps1
#
# 


#########################
## Public functions
#########################

Function Invoke-TabActivityIndicator {
    [CmdletBinding()]
    param(
        [Switch]
        $Error
    )

    end {
        if ($PowerTabConfig.TabActivityIndicator) {
            if ("ConsoleHost","PowerShellPlus Host" -contains $Host.Name) {
                if ($Error) {
                    $MessageBuffer = ConvertTo-BufferCellArray ([String[]]"[Err]") Yellow Red
                } else {
                    $MessageBuffer = ConvertTo-BufferCellArray ([String[]]"[Tab]") Yellow Blue
                }
                if ($MessageHandle) {
                    $MessageHandle.Content = $MessageBuffer
                    $MessageHandle.Show()
                } else {
                    $script:MessageHandle = New-Buffer $Host.UI.RawUI.WindowPosition $MessageBuffer
                }
                if ($Error) {
                    Start-Sleep 1
                }
            } else {
                Write-Progress "PowerTab" $Resources.invoke_tabactivityindicator_prog_status
            }
        }
    }
}


Function Remove-TabActivityIndicator {
    [CmdletBinding()]
    param()

    end {
        if ("ConsoleHost","PowerShellPlus Host" -contains $Host.Name) {
            if ($MessageHandle) {
                $MessageHandle.Clear()
                Remove-Variable -Name MessageHandle -Scope Script
            }
        } else {
            if ($PowerTabConfig.TabActivityIndicator) {
                Write-Progress "PowerTab" $Resources.invoke_tabactivityindicator_prog_status -Completed
            }
        }
    }
}


Function Invoke-TabItemSelector {
    [CmdletBinding(DefaultParameterSetName = "Values")]
    param(
        [Parameter(Position = 0)]
        [ValidateNotNull()]
        [String]
        $LastWord
        ,
        [ValidateSet("ConsoleList","Intellisense","Dynamic","Default","ObjectDefault")]
        [String]
        $SelectionHandler = "Default"
        ,
        [String]
        $ReturnWord
        ,
        [Parameter(ParameterSetName = "Values", ValueFromPipeline = $true)]
        [String[]]
        $Value
        ,
        [Parameter(ParameterSetName = "Objects", ValueFromPipeline = $true)]
        [Object[]]
        $Object
        ,
        [Switch]
        $ForceList
    )

    begin {
        [String[]]$Values = @()
        [Object[]]$Objects = @()
    }

    process {
        if ($PSCmdlet.ParameterSetName -eq "Values") {
            $Values += $Value
        } else {
            $Objects += $Object
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }

    end {
        Write-Debug "Invoke-TabItemSelector parameter set: $($PSCmdlet.ParameterSetName)"

        ## If dynamic, select an appropriate handler based on the current host
        if ($SelectionHandler -eq "Dynamic") {
            switch -exact ($Host.Name) {
                'ConsoleHost' {  ## PowerShell.exe
                    $SelectionHandler = "ConsoleList"
                    break
                }
                'PoshConsole' {
                    $SelectionHandler = "Default"
                    break
                }
                'PowerShellPlus Host' {
                    $SelectionHandler = "ConsoleList"
                    break
                }
                'Windows PowerShell ISE Host' {
                    $SelectionHandler = "Default"
                    break
                }
                default {
                    $SelectionHandler = "Default"
                    break
                }
            }
        }

        ## Block certain handlers in hosts that don't support them
        ## Example, ConsoleList and Intellisense won't work in PowerShell ISE
        [String[]]$IncompatibleHandlers = @()
        switch -exact ($Host.Name) {
            'ConsoleHost' {  ## PowerShell.exe
                $IncompatibleHandlers = @()
                break
            }
            'PoshConsole' {
                $IncompatibleHandlers = "ConsoleList","Intellisense"
                break
            }
            'PowerGUIHost' {
                $IncompatibleHandlers = "ConsoleList","Intellisense"
                break
            }
            'PowerGUIScriptEditorHost' {
                $IncompatibleHandlers = "ConsoleList","Intellisense"
                break
            }
            'PowerShellPlus Host' {
                $IncompatibleHandlers = @()
                break
            }
            'Windows PowerShell ISE Host' {
                $IncompatibleHandlers = "ConsoleList","Intellisense"
                break
            }
        }
        if ($IncompatibleHandlers -contains $SelectionHandler) {$SelectionHandler = "Default"}

        ## List of selection handlers that can handle objects
        ## TODO: Upgrade ConsoleList
        $ObjectHandlers = @("ConsoleList","ObjectDefault")

        if (($ObjectHandlers -contains $SelectionHandler) -and ($PSCmdlet.ParameterSetName -eq "Values")) {
            $Objects = foreach ($Item in $Values) {New-TabItem -Value $Item -Text $Item -Type Unknown}
        } elseif (($ObjectHandlers -notcontains $SelectionHandler) -and ($PSCmdlet.ParameterSetName -eq "Objects")) {
            $Values = foreach ($Item in $Objects) {$Item.Value}
        }

        switch -exact ($SelectionHandler) {
            'ConsoleList' {$Objects | Out-ConsoleList $LastWord $ReturnWord -ForceList:$ForceList}
            'Intellisense' {$Values | Invoke-Intellisense $LastWord}
            'ObjectDefault' {$Objects}
            'Default' {$Values}
        }
    }
}


Function New-TabItem {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Value
        ,
        [Parameter(Position = 1, ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Text = $Value
        ,
        [ValidateNotNullOrEmpty()]
        [String]
        $Type = "Unknown"
    )

    process {
        New-Object PSObject -Property @{Text=$Text; DisplayText=""; Value=$Value; Type=$Type}
    }
}

############

# .ExternalHelp TabExpansionLib-Help.xml
Function New-TabExpansionDatabase {
    [CmdletBinding()]
    param()

    end {
        $script:dsTabExpansionDatabase = New-Object System.Data.DataSet

        $dtCom = New-Object System.Data.DataTable
        [Void]($dtCom.Columns.Add('Name', [String]))
        [Void]($dtCom.Columns.Add('Description', [String]))
        $dtCom.TableName = 'COM'
        $dsTabExpansionDatabase.Tables.Add($dtCom)

        $dtCustom = New-Object System.Data.DataTable
        [Void]($dtCustom.Columns.Add('Filter', [String]))
        [Void]($dtCustom.Columns.Add('Text', [String]))
        [Void]($dtCustom.Columns.Add('Type', [String]))
        $dtCustom.TableName = 'Custom'
        $dsTabExpansionDatabase.Tables.Add($dtCustom)

        $dtTypes = New-Object System.Data.DataTable
        [Void]($dtTypes.Columns.Add('Name', [String]))
        [Void]($dtTypes.Columns.Add('DC', [String]))
        [Void]($dtTypes.Columns.Add('NS', [String]))
        $dtTypes.TableName = 'Types'
        $dsTabExpansionDatabase.Tables.Add($dtTypes)

        $dtWmi = New-Object System.Data.DataTable
        [Void]($dtWmi.Columns.Add('Name', [String]))
        [Void]($dtWmi.Columns.Add('Description', [String]))
        $dtWmi.TableName = 'WMI'
        $dsTabExpansionDatabase.Tables.Add($dtWmi)

        . (Join-Path $PSScriptRoot "TabExpansionCustomLib.ps1")
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function New-TabExpansionConfig {
    [CmdletBinding()]
    param(
        [Alias("FullName","Path")]
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath = $PowerTabConfig.Setup.ConfigurationPath
    )

    end {
        $script:dsTabExpansionConfig = InternalNewTabExpansionConfig $LiteralPath
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Import-TabExpansionDataBase {
    [CmdletBinding()]
    param(
        [Alias("FullName","Path")]
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath = $PowerTabConfig.Setup.DatabasePath
    )

    end {
        $script:dsTabExpansionDatabase = InternalImportTabExpansionDataBase $LiteralPath
        Write-Verbose ($Resources.import_tabexpansiondatabase_ver_success -f $LiteralPath)
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Export-TabExpansionDatabase {
    [CmdletBinding()]
    param(
        [Alias("FullName","Path")]
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath = $PowerTabConfig.Setup.DatabasePath
    )

    end {
        try {
            if (-not $PowerTabConfig.Setup.DatabasePath) {
                $BlankDatabasePath = $true
                Write-Verbose "Setting DatabasePath to $LiteralPath"  ## TODO: localize
                $PowerTabConfig.Setup.DatabasePath = $LiteralPath
            }

            if ($LiteralPath -eq "IsolatedStorage") {
                New-IsolatedStorageDirectory "PowerTab"
                $IsoFile = Open-IsolatedStorageFile "PowerTab\TabExpansion.xml" -Writable
                $dsTabExpansionDatabase.WriteXml($IsoFile)
            } else {
                if (-not (Test-Path (Split-Path $LiteralPath))) {
                    New-Item (Split-Path $LiteralPath) -ItemType Directory > $null
                }
                $dsTabExpansionDatabase.WriteXml($LiteralPath)
            }

            Write-Verbose ($Resources.export_tabexpansiondatabase_ver_success -f $LiteralPath)
        } finally {
            if ($BlankDatabasePath) {
                Write-Verbose "Reverting DatabasePath"  ## TODO: localize
                $PowerTabConfig.Setup.DatabasePath = ""
            }
        }
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Import-TabExpansionConfig {
    [CmdletBinding()]
    param(
        [Alias("FullName","Path")]
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath = $PowerTabConfig.Setup.ConfigurationPath
    )

    end {
        $Config = InternalImportTabExpansionConfig $LiteralPath

        ## Load Version
        [System.Version]$CurVersion = (Parse-Manifest).ModuleVersion
        $Version = $Config.Tables['Config'].Select("Name = 'Version'")[0].Value -as [System.Version]

        ## Upgrade if needed
        if ($Version -lt $CurVersion) {
            ## Upgrade config and database
            UpgradeTabExpansionDatabase ([Ref]$Config) ([Ref](New-Object System.Data.DataSet)) $Version
        } elseif ($Version -gt $CurVersion) {
            ## TODO: config is from a later version
        }

        $script:dsTabExpansionConfig = $Config

        ## Set version
        $PowerTabConfig.Version = $CurVersion

        Write-Verbose ($Resources.import_tabexpansionconfig_ver_success -f $LiteralPath)
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Export-TabExpansionConfig {
    [CmdletBinding()]
    param(
        [Alias("FullName","Path")]
        [Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath = $PowerTabConfig.Setup.ConfigurationPath
    )

    end {
        try {
            if (-not $PowerTabConfig.Setup.ConfigurationPath) {
                $BlankConfigurationPath = $true
                Write-Verbose "Setting ConfigurationPath to $LiteralPath"  ## TODO: localize
                $PowerTabConfig.Setup.ConfigurationPath = $LiteralPath
            }
            if (-not $PowerTabConfig.Setup.DatabasePath) {
                $BlankDatabasePath = $true
                if ($LiteralPath -eq "IsolatedStorage") {
                    $DatabasePath = $LiteralPath
                } else {
                    $DatabasePath = Join-Path (Split-Path $LiteralPath) TabExpansion.xml
                }
                Write-Verbose "Setting DatabasePath to $DatabasePath"  ## TODO: localize
                $PowerTabConfig.Setup.DatabasePath = $DatabasePath
            }

            if ($LiteralPath -eq "IsolatedStorage") {
                New-IsolatedStorageDirectory "PowerTab"
                $IsoFile = Open-IsolatedStorageFile "PowerTab\PowerTabConfig.xml" -Writable
                $dsTabExpansionConfig.Tables['Config'].WriteXml($IsoFile)
            } else {
                if (-not (Test-Path (Split-Path $LiteralPath))) {
                    New-Item (Split-Path $LiteralPath) -ItemType Directory > $null
                }
                $dsTabExpansionConfig.Tables['Config'].WriteXml($LiteralPath)
            }

            Write-Verbose ($Resources.export_tabexpansionconfig_ver_success -f $LiteralPath)
        } finally {
            if ($BlankConfigurationPath) {
                Write-Verbose "Reverting ConfigurationPath"  ## TODO: localize
                $PowerTabConfig.Setup.ConfigurationPath = ""
            }
            if ($BlankDatabasePath) {
                Write-Verbose "Reverting DatabasePath"  ## TODO: localize
                $PowerTabConfig.Setup.DatabasePath = ""
            }
            if ($IsoFile) {
                $IsoFile.Close()
            }
        }
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Import-TabExpansionTheme {
    [CmdletBinding(DefaultParameterSetName = "Name")]
    param(
        [Parameter(ParameterSetName = "Name", Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name
        ,
        [Alias("FullName","Path")]
        [Parameter(ParameterSetName = "LiteralPath", ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath
    )

    end {
		if ($PSCmdlet.ParameterSetName -eq "Name") {
            Import-Csv (Join-Path $PSScriptRoot "ColorThemes\Theme${Name}.csv") | ForEach-Object {$PowerTabConfig.Colors."$($_.Name)" = $_.Color}
        } else {
            Import-Csv $LiteralPath | ForEach-Object {$PowerTabConfig.Colors."$($_.Name)" = $_.Color}
        }
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Export-TabExpansionTheme {
    [CmdletBinding(DefaultParameterSetName = "Name")]
    param(
        [Parameter(ParameterSetName = "Name", Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name
        ,
        [Alias("FullName","Path")]
        [Parameter(ParameterSetName = "LiteralPath", ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $LiteralPath
    )

    process {
		if ($PSCmdlet.ParameterSetName -eq "Name") {
            $ExportPath = Join-Path $PSScriptRoot "ColorThemes\Theme${Name}.csv"
        } else {
            $ExportPath = $LiteralPath
        }
        $Colors = $PowerTabConfig.Colors | Get-Member -MemberType ScriptProperty |
            Select-Object @{Name='Name';Expression={$_.Name}},@{Name='Color';Expression={$PowerTabConfig.Colors."$($_.Name)"}} |
            Export-Csv $ExportPath -NoType

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}

############

# .ExternalHelp TabExpansionLib-Help.xml
Function Update-TabExpansionDataBase {
	[CmdletBinding(SupportsShouldProcess = $true, SupportsTransactions = $false,
		ConfirmImpact = "Low", DefaultParameterSetName = "")]
	param(
        [Switch]
        $Force
    )

    end {
        if ($Force -or $PSCmdlet.ShouldProcess($Resources.update_tabexpansiondatabase_type_conf_description,
            $Resources.update_tabexpansiondatabase_type_conf_inquire, $Resources.update_tabexpansiondatabase_type_conf_caption)) {
            Update-TabExpansionType
        }
        if ($Force -or $PSCmdlet.ShouldProcess($Resources.update_tabexpansiondatabase_wmi_conf_description,
            $Resources.update_tabexpansiondatabase_wmi_conf_inquire, $Resources.update_tabexpansiondatabase_wmi_conf_caption)) {
            Update-TabExpansionWmi
        }
        if ($Force -or $PSCmdlet.ShouldProcess($Resources.update_tabexpansiondatabase_com_conf_description,
            $Resources.update_tabexpansiondatabase_com_conf_inquire, $Resources.update_tabexpansiondatabase_com_conf_caption)) {
            Update-TabExpansionCom
        }
        if ($Force -or $PSCmdlet.ShouldProcess($Resources.update_tabexpansiondatabase_computer_conf_description,
            $Resources.update_tabexpansiondatabase_computer_conf_inquire, $Resources.update_tabexpansiondatabase_computer_conf_caption)) {
            Remove-TabExpansionComputer
            Add-TabExpansionComputer -NetView
        }
    }
}
Set-Alias udte Update-TabExpansionDataBase


# .ExternalHelp TabExpansionLib-Help.xml
Function Update-TabExpansionType {
	[CmdletBinding()]
    param()

    end {
        $dsTabExpansionDatabase.Tables['Types'].Clear()
        $Assemblies = [AppDomain]::CurrentDomain.GetAssemblies()
        $Assemblies | ForEach-Object {
                $i++; $Assembly = $_
                [Int]$AssemblyProgress = ($i * 100) / $Assemblies.Length
                Write-Progress "Adding Assembly $($_.GetName().Name):" $AssemblyProgress -PercentComplete $AssemblyProgress
                trap{$Types = $Assembly.GetExportedTypes() | Where-Object {$_.IsPublic -eq $true}; continue}; $Types = $_.GetTypes() |
                    Where-Object {$_.IsPublic -eq $true}
                $Types | Foreach-Object {$j = 0} {
                        $j++
                        if (($j % 200) -eq 0) {
                            [Int]$TypeProgress = ($j * 100) / $Types.Length
                            Write-Progress "Adding types:" $TypeProgress -PercentComplete $TypeProgress -Id 1
                        }
                        $dc = & {trap{continue;0}; $_.FullName.Split(".").Count - 1}
                        $ns = $_.NameSpace
                        [Void]$dsTabExpansionDatabase.Tables['Types'].Rows.Add($_.FullName, $dc, $ns)
                    }
            }
        Write-Progress "Adding types percent complete:" 100 -Id 1 -Completed

        # Add NameSpaces Without types
        $NL = $dsTabExpansionDatabase.Tables['Types'] | ForEach-Object {$i = 0} {
                $i++
                if (($i % 500) -eq 0) {
                    [Int]$TypeProgress = ($i * 100) / $dsTabExpansionDatabase.Tables['Types'].Rows.Count
                    Write-Progress "Adding namespaces:" $TypeProgress -PercentComplete $TypeProgress -Id 1
                } 
                $Split = [Regex]::Split($_.Name,'\.')
                if ($Split.Length -gt 2) {
                    0..($Split.Length - 3) | ForEach-Object {$ofs='.'; "$($Split[0..($_)])"}
                }
            } | Sort-Object -Unique
        $nl | ForEach-Object {[Void]$dsTabExpansionDatabase.Tables['Types'].Rows.Add("Dummy", $_.Split('.').Count, $_)}
        Write-Progress "Adding NameSpaces percent complete:" 100 -Id 1 -Completed
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Add-TabExpansionType {
	[CmdletBinding()]
    param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNull()]
        [System.Reflection.Assembly]
        $Assembly
    )

    process {
        $Assembly | ForEach-Object {
                $i++; $ass = $_
                trap{$Types = $ass.GetExportedTypes() | Where-Object {$_.IsPublic -eq $true}; continue}; $Types = $_.GetTypes() |
                    Where-Object {$_.IsPublic -eq $true}
                $Types | ForEach-Object {$j = 0} {
                        $j++;
                        if (($j % 200) -eq 0) {
                            [Int]$TypeProgress = ($j * 100) / $Types.Length
                            Write-Progress "Adding types:" $TypeProgress -PercentComplete $TypeProgress -Id 1
                        } 
                        $dc = & {trap{continue;0}; $_.FullName.Split(".").Count - 1} 
                        $ns = $_.NameSpace 
                        [Void]$dsTabExpansionDatabase.Tables['Types'].Rows.Add($_.FullName, $dc, $ns)
                    }
            }
        Write-Progress "Adding types percent complete:" "100" -Id 1 -Completed

        # Add NameSpaces Without types
        $NL = $dsTabExpansionDatabase.Tables['Types'].select("ns = '$($ass.GetName().name)'") |
            ForEach-Object {$i = 0} {$i++
                if (($i % 500) -eq 0) {
                    [Int]$TypeProgress = ($i * 100) / $dsTabExpansionDatabase.Tables['Types'].Rows.Count
                    Write-Progress "Adding namespaces:" $TypeProgress -PercentComplete $TypeProgress -Id 1
                }
                $Split = [Regex]::Split($_.Name,'\.')
                if ($Split.Length -gt 2) {
                    0..($Split.Length - 3) | ForEach-Object {$ofs='.'; "$($Split[0..($_)])"}
                }
            } | Sort-Object -Unique
        $nl | ForEach-Object {[Void]$dsTabExpansionDatabase.Tables['Types'].Rows.Add("Dummy",$_.Split('.').Count, $_)}
        Write-Progress "Adding NameSpaces percent complete:" 100 -Id 1 -Completed

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            Write-Progress "Adding NameSpaces percent complete:" 100 -Id 1 -Completed
            break
        }
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Update-TabExpansionWmi {
	[CmdletBinding()]
    param()

    end {
        $dsTabExpansionDatabase.Tables['WMI'].Clear()

        # Set Enumeration Options
        $Options = New-Object System.Management.EnumerationOptions
        $Options.EnumerateDeep = $true
        $Options.UseAmendedQualifiers = $true

        $i = 0 ; Write-Progress $Resources.update_tabexpansiondatabase_wmi_activity $i
        foreach ($Class in (([WmiClass]'').PSBase.GetSubclasses($Options))) {
            $i++ ; if ($i % 10 -eq 0) {Write-Progress $Resources.update_tabexpansiondatabase_wmi_activity $i}
            $Description = try { $Class.GetQualifierValue('Description') } catch { }
            [Void]$dsTabExpansionDatabase.Tables['WMI'].Rows.Add($Class.Name, $Description )
        }
        Write-Progress $Resources.update_tabexpansiondatabase_wmi_activity $i -Completed
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Update-TabExpansionCom {
	[CmdletBinding()]
    param()

    end {
        $dsTabExpansionDatabase.Tables['COM'].Clear()

        $i = 0 ; Write-Progress $Resources.update_tabexpansiondatabase_com_activity $i
        foreach ($Class in (Get-WmiObject Win32_ClassicCOMClassSetting -Filter "VersionIndependentProgId LIKE '%'" | Sort-Object VersionIndependentProgId)) {
            $i++ ; if ($i % 10 -eq 0) {Write-Progress $Resources.update_tabexpansiondatabase_com_activity $i}
            [Void]$dsTabExpansionDatabase.Tables['COM'].Rows.Add($Class.VersionIndependentProgId, $Class.Description)
        }
        Write-Progress $Resources.update_tabexpansiondatabase_com_activity $i -Completed
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Add-TabExpansionComputer {
	[CmdletBinding(SupportsShouldProcess = $false, SupportsTransactions = $false,
		ConfirmImpact = "None", DefaultParameterSetName = "Name")]
	param(
        [Alias("Name")]
		[Parameter(ParameterSetName = "Name", Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $ComputerName
        ,
		[Parameter(ParameterSetName = "OU", Position = 0, Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateNotNull()]
        [System.DirectoryServices.DirectoryEntry]
        $OU
        ,
		[Parameter(ParameterSetName = "NetView")]
        [Switch]
        $NetView
    )

    process {
        $count = 0
        if ($PSCmdlet.ParameterSetName -eq "Name") {
            Add-TabExpansion $ComputerName $ComputerName "Computer"
        } elseif ($PSCmdlet.ParameterSetName -eq "OU") {
            foreach ($Computer in ($OU.PSBase.get_Children() | Select-Object @{Name='Name';Expression={$_.cn[0]}})) {
                $count++; if ($count % 5 -eq 0) {Write-Progress $Resources.update_tabexpansiondatabase_computer_activity $count}
                Add-TabExpansion $Computer.Name $Computer.Name Computer
            }
        } elseif ($PSCmdlet.ParameterSetName -eq "NetView") {
            foreach ($Line in (net view)) {
                if ($Line -match '\\\\(.*?) ') {
                    $Computer = $Matches[1]
                    $count++; if ($count % 5 -eq 0) {Write-Progress $Resources.update_tabexpansiondatabase_computer_activity $count}
                    Add-TabExpansion $Computer $Computer Computer
                }
            }
        }
        if ($PSCmdlet.ParameterSetName -ne "Name") {
            Write-Progress $Resources.update_tabexpansiondatabase_computer_activity $count -Completed
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            if ($PSCmdlet.ParameterSetName -ne "Name") {
                Write-Progress $Resources.update_tabexpansiondatabase_computer_activity $count -Completed
            }
            break
        }
    }
}


# .ExternalHelp TabExpansionLib-Help.xml
Function Remove-TabExpansionComputer {
	[CmdletBinding()]
    param()

    end {
        foreach ($Computer in $dsTabExpansionDatabase.Tables['Custom'].Select("Type LIKE 'Computer'")) {
            $Computer.Delete()
        }
    }
}

############

# .ExternalHelp TabExpansionLib-Help.xml
Function Get-TabExpansion {
	[CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [String]
        $Filter = "*"
        ,
        [Parameter(Position = 1)]
        [String]
        $Type = "*"
    )

    ## TODO: Make Type a dynamic validateset?
    ## TODO: escape special characters?

    process {
        ## Split filter on internal wildcards as DataTables do not support them
        $Filters = @($Filter -split '(?<=.)[\*%](?=.)')
        if ($Filters.Count -gt 1) {
            $Filters[0] = $Filters[0] + "*"  ## First item
            $Filters[-1] = "*" + $Filters[-1]  ## Last item

            if ($Filters.Count -gt 2) {
                foreach ($Index in 1..($Filters.Count - 2)) {
                    $Filters[$Index] = "*" + $Filters[$Index] + "*"
                }
            }
        }

        ## Run query
        if ("COM","Types","WMI" -contains $Type){
            ## Construct query from multiple filters
            $Query = "Name LIKE '$($Filters[0])'"
            foreach ($Filter in $Filters[1..($Filters.Count - 1)]) {
                $Query += " AND Name LIKE '$Filter'"
            }
            $dsTabExpansionDatabase.Tables[$Type].Select($Query)
        } else {
            ## Construct query from multiple filters
            $Query = "Filter LIKE '$($Filters[0])'"
            foreach ($Filter in $Filters[1..($Filters.Count - 1)]) {
                $Query += " AND Filter LIKE '$Filter'"
            }
            $dsTabExpansionDatabase.Tables["Custom"].Select("$Query AND Type LIKE '$Type'")
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}
Set-Alias gte Get-TabExpansion


# .ExternalHelp TabExpansionLib-Help.xml
Function Add-TabExpansion {
	[CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Filter
        ,
        [Parameter(Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Text
        ,
        [Parameter(Position = 2)]
        [ValidateNotNull()]
        [String]
        $Type = 'Custom'
    )

    ## TODO: Add -PassThru support
    process {
        [Void]$dsTabExpansionDatabase.Tables['Custom'].Rows.Add($Filter, $Text, $Type)

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}
Set-Alias ate Add-TabExpansion


# .ExternalHelp TabExpansionLib-Help.xml
Function Remove-TabExpansion {
	[CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Filter
    )

    ## TODO: Add type
    process {
        $Filter = $Filter -replace "\*","%"

        foreach ($Item in $dsTabExpansionDatabase.Tables['Custom'].Select("Filter LIKE '$Filter'")) {
            $Item.Delete()
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}
Set-Alias rte Remove-TabExpansion


# .ExternalHelp TabExpansionLib-Help.xml
Function Invoke-TabExpansionEditor {
	[CmdletBinding()]
    param()

    end {
        [System.Version]$CurVersion = (Parse-Manifest).ModuleVersion

        $Form = New-Object System.Windows.Forms.Form
        $Form.Size = New-Object System.Drawing.Size @(500,300)
        $Form.Text = "PowerTab $CurVersion PowerShell TabExpansion Library"

        $DataGrid = New-Object System.Windows.Forms.DataGrid
        $DataGrid.CaptionText = "Custom TabExpansion Database Editor"
        $DataGrid.AllowSorting = $true
        $DataGrid.DataSource = $dsTabExpansionDatabase.PSObject.BaseObject
        $DataGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
        $Form.Controls.Add($DataGrid)
        $StatusBar = New-Object System.Windows.Forms.Statusbar
        $StatusBar.Text = " /\/\o\/\/ 2007 http://thePowerShellGuy.com"
        $Form.Controls.Add($StatusBar)

        ## Show the Form
        $Form.Add_Shown({$Form.Activate(); $DataGrid.Expand(0)})
        [Void]$Form.ShowDialog()
    }
}
Set-Alias itee Invoke-TabExpansionEditor

############

# .ExternalHelp TabExpansionLib-Help.xml
Function Register-TabExpansion {
	[CmdletBinding()]
    param(
		[Parameter(Position = 0, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Name
        ,
		[Parameter(Position = 1, Mandatory = $true, ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNull()]
        [ScriptBlock]
        $Handler
        ,
        [ValidateSet("Command","CommandInfo","Parameter","ParameterName")]
        [String]
        $Type = "Command"
        ,
        [Switch]
        $Force
    )
    
    process {
        if ($Type -eq "Parameter") {
            if (-not $TabExpansionParameterRegistry[$Name] -or $Force) {
                $TabExpansionParameterRegistry[$Name] = $Handler
            }
        } elseif ($Type -eq "ParameterName") {
            if (-not $TabExpansionParameterNameRegistry[$Name] -or $Force) {
                $TabExpansionParameterNameRegistry[$Name] = $Handler
            }
        } elseif ($Type -eq "CommandInfo") {
            if (-not $TabExpansionCommandInfoRegistry[$Name] -or $Force) {
                $TabExpansionCommandInfoRegistry[$Name] = $Handler
            }
        } else {
            if (-not $TabExpansionCommandRegistry[$Name] -or $Force) {
                $TabExpansionCommandRegistry[$Name] = $Handler
            }
        }

        trap [System.Management.Automation.PipelineStoppedException] {
            ## Pipeline was stopped
            break
        }
    }
}
Set-Alias rgte Register-TabExpansion



#########################
## Private functions
#########################

Function Initialize-PowerTab {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$ConfigurationPath = $PowerTabConfig.Setup.ConfigurationPath
    )

    ## Load Configuration
    if ($ConfigurationPath -and ((Test-Path $ConfigurationPath) -or ($ConfigurationPath -eq "IsolatedStorage"))) {
        $Config = InternalImportTabExpansionConfig $ConfigurationPath
    } else {
        ## TODO: Throw error or create new config?
        #$Config = InternalNewTabExpansionConfig $ConfigurationPath
    }

    ## Load Version
    [System.Version]$CurVersion = (Parse-Manifest).ModuleVersion
    $Version = $Config.Tables['Config'].Select("Name = 'Version'")[0].Value -as [System.Version]

    ## Load Database
    if ($Version -lt ([System.Version]'0.99.3.0')) {
        $DatabaseName = $Config.Tables['Config'].select("Name = 'DatabaseName'")[0].Value
        $DatabasePath = Join-Path ($Config.Tables['Config'].select("Name = 'DatabasePath'")[0].Value) $DatabaseName
    } else {
        $DatabasePath = $Config.Tables['Config'].select("Name = 'DatabasePath'")[0].Value
    }
    if (!(Split-Path $DatabasePath)) {
        $DatabasePath = Join-Path $PSScriptRoot $DataBasePath
    }

    $Database = InternalImportTabExpansionDataBase $DatabasePath

    ## Upgrade if needed
    if ($Version -lt $CurVersion) {
        ## Upgrade config and database
        UpgradeTabExpansionDatabase ([Ref]$Config) ([Ref]$Database) $Version
    } elseif ($Version -gt $CurVersion) {
        ## TODO: config is from a later version
    }

    ## Config and database are good
    $script:dsTabExpansionConfig = $Config
    $script:dsTabExpansionDatabase = $Database

    ## Create the user interface for the PowerTab settings
    CreatePowerTabConfig

    ## Set version
    $PowerTabConfig.Version = $CurVersion
}


Function UpgradeTabExpansionDatabase {
    [CmdletBinding()]
    param(
        [Ref]$Config
        ,
        [Ref]$Database
        ,
        [System.Version]$Version
    )

    <#
    For future releases, add new if conditions only if an upgrade path is needed due to changes
    in the database or config structure.  Or to add default values for new config settings.
    #>

    if ($Version -lt [System.Version]'0.99.3.0') {
        ## Upgrade versions from the first version of PowerTab
        Write-Host "Upgrading from version $Version"
        UpgradePowerTab99 $Config $Database
        $Version = '0.99.3.0'
    }
    if ($Version -lt [System.Version]'0.99.5.0') {
        ## Upgrade versions from the first version of PowerTab
        Write-Host "Upgrading from version $Version"
        UpgradePowerTab993 $Config $Database
        $Version = '0.99.5.0'
    }
}


Function UpgradePowerTab99 {
    [CmdletBinding()]
    param(
        [Ref]$Config
        ,
        [Ref]$Database
    )

    $Config.Value.Tables['Config'].Select("Name = 'InstallPath' AND Category = 'Setup'") | ForEach-Object {$_.Delete()}
    if ($Database.Value.Tables['Config']) {
        $Database.Value.Tables.Remove('Config')
        trap {continue}
    }
    if ($Database.Value.Tables['Cache']) {
        $Database.Value.Tables.Remove('Cache')
        trap {continue}
    }
    $ConfigurationPath = $Config.Value.Tables['Config'].Select("Name = 'ConfigurationPath'")[0].Value
    $Config.Value.Tables['Config'].Select("Name = 'ConfigurationPath'")[0].Value = Join-Path $ConfigurationPath "PowerTabConfig.xml"
    $DatabasePath = $Config.Value.Tables['Config'].Select("Name = 'DatabasePath'")[0].Value
    $DatabaseName = $Config.Value.Tables['Config'].Select("Name = 'DatabaseName'")[0].Value
    $Config.Value.Tables['Config'].Select("Name = 'DatabasePath'")[0].Value = Join-Path $DatabasePath $DatabaseName
    $Config.Value.Tables['Config'].Select("Name = 'DatabaseName' AND Category = 'Setup'") | ForEach-Object {$_.Delete()}
}


Function UpgradePowerTab993 {
    [CmdletBinding()]
    param(
        [Ref]$Config
        ,
        [Ref]$Database
    )

    $Config.Value.Tables['Config'].Select("Name = 'SpaceCompleteFileSystem'") | ForEach-Object {$_.Delete()}
    ## Add VisualStudioTabBehavior
    $row = $Config.Value.Tables['Config'].NewRow()
    $row.Name = 'VisualStudioTabBehavior'
    $row.Type = 'Bool'
    $row.Category = 'Global'
    $row.Value = [Int]($False)
    $Config.Value.Tables['Config'].Rows.Add($row)
}


Function InternalNewTabExpansionConfig {
    [CmdletBinding()]
    param(
        [String]$ConfigurationPath
    )
    
    if ($ConfigurationPath) {
        if ($ConfigurationPath -eq "IsolatedStorage") {
            $DatabasePath = $ConfigurationPath
        } else {
            $DatabasePath = Join-Path (Split-Path $ConfigurationPath) "TabExpansion.xml"
        }
    }

    $Config = New-Object System.Data.DataSet

    $dtConfig = New-Object System.Data.DataTable
    [Void]$dtConfig.Columns.Add('Category', [String])
    [Void]$dtConfig.Columns.Add('Name', [String])
    [Void]$dtConfig.Columns.Add('Value')
    [Void]$dtConfig.Columns.Add('Type')
    $dtConfig.TableName = 'Config'

    ## Add global configuration
    @{
        Version = (Parse-Manifest).ModuleVersion
        DefaultHandler = 'Dynamic'
        AlternateHandler = 'Dynamic'
        CustomUserFunction = 'Write-Warning'
        CustomCompletionChars = ']:)'
    }.GetEnumerator() | Foreach-Object {
            $row = $dtConfig.NewRow()
            $row.Name = $_.Name
            $row.Type = 'String'
            $row.Category = 'Global'
            $row.Value = $_.Value
            $dtConfig.Rows.Add($row)
        }
    @($dtConfig.Select("Name = 'Version'"))[0].Category = 'Version'

    ## Add color configuration
    $Items = `
        'BorderColor',
        'BorderBackColor',
        'BackColor',
        'TextColor',
        'SelectedBackColor',
        'SelectedTextColor',
        'BorderTextColor',
        'FilterColor'
    $DefaultColors = `
        'Blue',
        'DarkBlue',
        'DarkGray',
        'Yellow',
        'DarkRed',
        'Red',
        'Yellow',
        'DarkGray'
    0..($Items.GetUpperBound(0)) | Foreach-Object {
            $row = $dtConfig.NewRow()
            $row.Name = $items[$_]
            $row.Category = 'Colors'
            $row.Type = 'ConsoleColor'
            $row.Value = [ConsoleColor]($DefaultColors[$_])
            $dtConfig.Rows.Add($row)
        }

    ## Add shortcut configuration
    @{
        Alias   = '@'
        Partial = '%'
        Native  = '!'
        Invoke  = '&'
        Custom  = '^'
        CustomFunction  = '#'
    }.GetEnumerator() | Foreach-Object {
            $row = $dtConfig.NewRow()
            $row.Name = $_.Name
            $row.Type = 'String'
            $row.Category = 'ShortcutChars'
            $row.Value = $_.Value
            $dtConfig.Rows.Add($row)
        }

    ## Add setup configuration
    @{
        ConfigurationPath = $ConfigurationPath
        DatabasePath = $DatabasePath
    }.GetEnumerator() | Foreach-Object {
            $row = $dtConfig.NewRow()
            $row.Name = $_.Name
            $row.Type = 'String'
            $row.Category = 'Setup'
            $row.Value = $_.Value
            $dtConfig.Rows.Add($row)
        }

    $Options = @{
            Enabled = $True
            ShowBanner = $True
            TabActivityIndicator = $True
            DoubleTabEnabled = $False
            DoubleTabLock = $False
            AliasQuickExpand = $False
            FileSystemExpand = $True
            ShowAccessorMethods = $True
            DoubleBorder = $True
            CustomFunctionEnabled = $False
            IgnoreConfirmPreference = $False
            ## ConsoleList
            CloseListOnEmptyFilter = $True
            DotComplete = $True
            AutoExpandOnDot = $True
            BackSlashComplete = $True
            AutoExpandOnBackSlash = $True
            CustomComplete = $True
            SpaceComplete = $True
            VisualStudioTabBehavior = $False
        }
    $Options.GetEnumerator() | Foreach-Object {
            $row = $dtConfig.NewRow()
            $row.Name = $_.Name
            $row.Type = 'Bool'
            $row.Category = 'Global'
            $row.Value = [Int]($_.Value)
            $dtConfig.Rows.Add($row)
        }

    @{
        ## ConsoleList
        MinimumListItems   = '2'
        FastScrollItemcount = '10'
    }.GetEnumerator() | ForEach-Object {
            $row = $dtConfig.NewRow()
            $row.Name = $_.Name
            $row.Type = 'Int'
            $row.Category = 'Global'
            $row.Value = $_.Value
            $dtConfig.Rows.Add($row)
        }

    $Config.Tables.Add($dtConfig)
    $Config
}


Function InternalImportTabExpansionDataBase {
    [CmdletBinding()]
    param(
        [String]$LiteralPath
    )

    $Database = New-Object System.Data.DataSet
    if (($LiteralPath -eq "IsolatedStorage") -and (Test-IsolatedStoragePath "PowerTab\TabExpansion.xml")) {
        $UserIsoStorage = [System.IO.IsolatedStorage.IsolatedStorageFile]::GetUserStoreForAssembly()
        $IsoFile = New-Object System.IO.IsolatedStorage.IsolatedStorageFileStream("PowerTab\TabExpansion.xml",
            [System.IO.FileMode]::Open, $UserIsoStorage)
        [Void]$Database.ReadXml($IsoFile)
    } elseif (Test-Path $LiteralPath) {
        if (![System.IO.Path]::IsPathRooted($_)) {
            $LiteralPath = Resolve-Path $LiteralPath
        }
        [Void]$Database.ReadXml($LiteralPath)
    }

    if (!$Database.Tables["COM"]) {
        $dtCom = New-Object System.Data.DataTable
        [Void]($dtCom.Columns.Add('Name', [String]))
        [Void]($dtCom.Columns.Add('Description', [String]))
        $dtCom.TableName = 'COM'
        $Database.Tables.Add($dtCom)
    }
    if (!$Database.Tables["Custom"]) {
        $dtCustom = New-Object System.Data.DataTable
        [Void]($dtCustom.Columns.Add('Filter', [String]))
        [Void]($dtCustom.Columns.Add('Text', [String]))
        [Void]($dtCustom.Columns.Add('Type', [String]))
        $dtCustom.TableName = 'Custom'
        $Database.Tables.Add($dtCustom)
    }
    if (!$Database.Tables["Types"]) {
        $dtTypes = New-Object System.Data.DataTable
        [Void]($dtTypes.Columns.Add('Name', [String]))
        [Void]($dtTypes.Columns.Add('DC', [String]))
        [Void]($dtTypes.Columns.Add('NS', [String]))
        $dtTypes.TableName = 'Types'
        $Database.Tables.Add($dtTypes)
    }
    if (!$Database.Tables["WMI"]) {
        $dtWmi = New-Object System.Data.DataTable
        [Void]($dtWmi.Columns.Add('Name', [String]))
        [Void]($dtWmi.Columns.Add('Description', [String]))
        $dtWmi.TableName = 'WMI'
        $Database.Tables.Add($dtWmi)
    }

    $Database
}


Function InternalImportTabExpansionConfig {
    [CmdletBinding()]
    param(
        [String]$LiteralPath
    )

    $Config = New-Object System.Data.DataSet
    if ($LiteralPath -eq "IsolatedStorage") {
        $UserIsoStorage = [System.IO.IsolatedStorage.IsolatedStorageFile]::GetUserStoreForAssembly()
        $IsoFile = New-Object System.IO.IsolatedStorage.IsolatedStorageFileStream("PowerTab\PowerTabConfig.xml",
            [System.IO.FileMode]::Open, $UserIsoStorage)
        [Void]$Config.ReadXml($IsoFile, 'InferSchema')
    } elseif (Test-Path $LiteralPath) {
        if (![System.IO.Path]::IsPathRooted($_)) {
            $LiteralPath = Resolve-Path $LiteralPath
        }
        [Void]$Config.ReadXml($LiteralPath, 'InferSchema')
    } else {
        $Config = InternalNewTabExpansionConfig $LiteralPath
    }

    $Version = $Config.Tables['Config'].Select("Name = 'Version'")[0].Value -as [System.Version]
    if ($Version -eq $null) {$Config.Tables['Config'].Select("Name = 'Version'")[0].Value = '0.99.0.0'}

    $Config
}


Function CreatePowerTabConfig {
    [CmdletBinding()]
    param()
    
    $script:PowerTabConfig = New-Object PSObject

    Add-Member -InputObject $PowerTabConfig -MemberType ScriptProperty -Name Version `
        -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
            "`$dsTabExpansionConfig.Tables['Config'].Select(`"Name = 'Version'`")[0].Value") `
        -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
            "trap {Write-Warning `$_; continue}
            `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = 'Version'`")[0].Value = [String]`$args[0]")

    ## Add Enable ScriptProperty
    Add-Member -InputObject $PowerTabConfig -MemberType ScriptProperty -Name Enabled `
        -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
            "`$v = `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = 'Enabled'`")[0]
            [Bool][Int]`$v.Value") `
        -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
            "trap {Write-Warning `$_; continue}
            [Int]`$val = [Bool]`$args[0]
            `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = 'Enabled'`")[0].Value = `$val
            if ([Bool]`$val) {
                . `"`$PSScriptRoot\TabExpansion.ps1`"
            } else {
                Set-Content Function:\TabExpansion -Value `$OldTabExpansion
            }") `
        -Force

    Add-Member -InputObject $PowerTabConfig -MemberType NoteProperty -Name Colors -Value (New-Object PSObject)
    Add-Member -InputObject $PowerTabConfig.Colors -MemberType ScriptMethod -Name ToString -Value {"{PowerTab Color Configuration}"} -Force

    Add-Member -InputObject $PowerTabConfig -MemberType NoteProperty -Name ShortcutChars -Value (New-Object PSObject)
    Add-Member -InputObject $PowerTabConfig.ShortcutChars -MemberType ScriptMethod -Name ToString -Value {"{PowerTab Shortcut Characters}"} -Force

    Add-Member -InputObject $PowerTabConfig -MemberType NoteProperty -Name Setup -Value (New-Object PSObject)
    Add-Member -InputObject $PowerTabConfig.Setup -MemberType ScriptMethod -Name ToString -Value {"{PowerTab Setup Data}"} -Force

    ## Make global properties on config object
    $dsTabExpansionConfig.Tables['Config'].Select("Category = 'Global'") | Where-Object {$_.Name -ne "Enabled"} | ForEach-Object {
            Add-Member -InputObject $PowerTabConfig -MemberType ScriptProperty -Name $_.Name `
                -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
                    "`$v = `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0]
                    if (`$v.Type -eq 'Bool') {
                        [Bool][Int]`$v.Value
                    } else {
                        [$($_.Type)](`$v.Value)
                    }") `
                -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                    "trap {Write-Warning `$_; continue}
                    `$val = [$($_.Type)]`$args[0]
                     if ('$($_.Type)' -eq 'bool') {`$val = [Int]`$val}
                    `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0].Value = `$val") `
                -Force
        }

    ## Make color properties on config object
    $dsTabExpansionConfig.Tables['Config'].Select("Category = 'Colors'") | Foreach-Object {
            Add-Member -InputObject $PowerTabConfig.Colors -MemberType ScriptProperty -Name $_.Name `
                -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
                 "`$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0].Value") `
                -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                    "trap {Write-Warning `$_; continue}
                    `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0].Value = [ConsoleColor]`$args[0]") `
                -Force
        }

    ## Make shortcut properties on config object
    $dsTabExpansionConfig.Tables['Config'].Select("Category = 'ShortcutChars'") | Foreach-Object {
            Add-Member -InputObject $PowerTabConfig.ShortcutChars -MemberType ScriptProperty -Name $_.Name `
                -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
                    "`$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0].Value") `
                -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                    "trap {Write-Warning `$_; continue}
                    `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0].Value = `$args[0]") `
                -Force
        }

    ## Make Setup properties on Config Object
    $dsTabExpansionConfig.Tables['Config'].Select("Category = 'Setup'") | Foreach-Object {
            Add-Member -InputObject $PowerTabConfig.Setup -MemberType ScriptProperty -Name $_.Name `
                -Value $ExecutionContext.InvokeCommand.NewScriptBlock(
                    "`$v = `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0]
                    if (`$v.Type -eq 'Bool') {
                        [Bool][Int]`$v.Value
                    } else {
                        [$($_.Type)](`$v.Value)
                    }") `
                -SecondValue $ExecutionContext.InvokeCommand.NewScriptBlock(
                    "trap {Write-Warning `$_; continue}
                    `$val = [$($_.Type)]`$args[0]
                     if ('$($_.Type)' -eq 'bool') {`$val = [Int]`$val}
                    `$dsTabExpansionConfig.Tables['Config'].Select(`"Name = '$($_.Name)'`")[0].Value = `$val") `
                -Force
        }
}

# SIG # Begin signature block
# MIIbaQYJKoZIhvcNAQcCoIIbWjCCG1YCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUwJQqlW6VRZ2oiipqbYJ5KBxs
# HZqgghYbMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
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
# AQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQK1C4JOZvFfFdLGpnO
# OCJlPuWZfDANBgkqhkiG9w0BAQEFAASCAgCB45k6ob1TiCS0eS5PXlPJFHvLgHsr
# GOb6HxkJGiE225brSxYgY+Zc49RKHuscgGDOVBgo9zZVW3taSi2MBopB/4CZuKv1
# oWyuAjXlUvIPGAgZ6udfOnRtS89VFeHLRTiaCPDp9/jaUw0qf6/yieez4Sfhp6GZ
# j1A7IKhusQK8oLj9+tI0nw2uPj0hfkJUoBR6l3Hvx19XI90UHgNFqSniTRDFMWJY
# UQs68wzi/AIf5tzsNtNcRLeXOXALq5LfrWFYGwrMArfkJ0nyzxy6w0N9zahkvev0
# igXsD4YqXfA/yVJzGCKxfHUOqNdAegwJyQdzqFz+jeuenGrL+0rn3tgByr+jxt+g
# +9BX44JeDtTHDOpjLZ927N7vkOyxd+H7huJJ6f7VRNTXzrhhcdXhTmj0DdNVKNaP
# gZHiuUXw3wf0qlRMnXgG7le5r+q6k+akQ0pWI2gLdgUdfiNI7Q9jnASUhrWP00Nz
# 7hrgu77my8wrZb/JdUU5MYO2XiWbz61OFC28IUuPWTXIca4HUneCea582jNxQaJY
# 4KFvww141gMczCpBZFZWJ+KHPX6qX1LX6nSZJPKfWWhBjHjzMx9j7R499e73kss8
# jl6KRzmxFb+rxxvk5v2LZ0bo+YTiRkCXOeUva7898mCaeDsYQ8GEcIqaOFqHCnk6
# qfVg7EJ/lJffP6GCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBnMFMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UEAxMiVmVy
# aVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT2bg+9raN
# 7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTIxMDI3MDI0NDUxWjAjBgkqhkiG9w0BCQQxFgQUSYGEP3ixXm1i
# fqOvOaggs+e7oqwwDQYJKoZIhvcNAQEBBQAEgYAJQ1s90oleSyKx8K/ZGYUSnII+
# twOuK1tXg297cSmQAo2kSPCd2aQReFx7Y8rfiI9fSfTV/vD0fr85XSDW12FIBcyB
# vwDHuHLgk2sFIRoAtVrVXbc7QcotLPAbWYXSJoWAdM1AgK7c7jcwCy1Tjewv/OhF
# /QaYtNVhlgFBwP2pmQ==
# SIG # End signature block
