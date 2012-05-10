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
# MIIY1AYJKoZIhvcNAQcCoIIYxTCCGMECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUwJQqlW6VRZ2oiipqbYJ5KBxs
# HZqgghSGMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
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
# CSqGSIb3DQEJBDEWBBQK1C4JOZvFfFdLGpnOOCJlPuWZfDANBgkqhkiG9w0BAQEF
# AASCAQBiqNigrs+8+9TrimB1RbkDI+qC0RUEV3DMS921OkNcYClapRegvL6TO3sD
# N7zuhkwjVXPPO9wMuENMuJ2p2xdoGxdPhN88Ae8x7Tvp3YekUDZKCBP2nB4N4LM1
# z1gsnl9DgnW+jp71d066drC8MokH6XyPz6Ob1p/zJaIPx5mZ8n4uHTSYxWR78wSg
# g65WPPO+ZKM0tUeVdfyTtShnMmqnPyKjPIGcWhtIDrvnW2yhhs2A6KrR4bbySLUh
# TSlcSL6tVzHO7ebX2LiR+qFsQ1OBMZAF7cgJGAzPmTZKIXUCgGUDojqJnOnozmH4
# RTrgT7Y7NPhNeX+EQYVtmrhpJ2v6oYIBfzCCAXsGCSqGSIb3DQEJBjGCAWwwggFo
# AgEBMGcwUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBAhA4Jdf6
# +GGvnvSQ5ya11lrVMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0xMTEwMjgxOTUzMzdaMCMGCSqGSIb3DQEJBDEW
# BBQeYp4YYPNDawPyCKBwKha0NGmN3TANBgkqhkiG9w0BAQEFAASBgEha+giZaZ/L
# 10nBckoJuEdMU1b6IPa+skTLfiVZZ5JnuO82hZ+kq4XeUNQO4+r5VLJN8AK5kBKl
# PFb8PW1mq+Hlrw1rWFdiJCO7H98fNd1Cid4KmQIgoc9OTFJyQdjIhDsG7eMJakd5
# 7vJH6FCwXZDvSC99TQstRyrOSzeH3xB+
# SIG # End signature block
