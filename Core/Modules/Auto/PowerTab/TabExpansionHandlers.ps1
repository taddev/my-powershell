
<########################
## Notes

- The closures in this file make the code run outside the PowerTab module's context.  This avoids
some problems like Get-Module only seeing the modules loaded within PowerTab, or private functions
showing up from Get-Command.
########################>


#########################
## Command handlers
#########################

## Alias
& {
    $AliasHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-Alias -Name "$Argument*" | Select-Object -ExpandProperty Name
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Export-Alias" $AliasHandler -Type "Command"
    Register-TabExpansion "Get-Alias" $AliasHandler -Type "Command"
}

## Get-Command (-Module mainly)
Register-TabExpansion "Get-Command" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    if (($Context.Parameter -eq "ArgumentList") -and ($Context.PositionalParameter -eq 0)) {
        $Context.Parameter = "Name"  ## Fix for odd default parameter set on Get-Command
    }
    switch -exact ($Context.Parameter) {
        'Module' {
            $TabExpansionHasOutput.Value = $true
            Get-Module "$Argument*" | Select-Object -ExpandProperty Name | Sort-Object
        }
        'Name' {
            $TabExpansionHasOutput.Value = $true
            $Parameters = @{}
            if ($Context.OtherParameters["Module"]) {
                $Parameters["Module"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["Module"]
            }
            if ($Context.OtherParameters["CommandType"]) {
                $Parameters["CommandType"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["CommandType"]
            } else {
                $Parameters["CommandType"] = "Alias","Function","Filter","Cmdlet"
            }
            Get-Command "$Argument*" @Parameters | Select-Object -ExpandProperty Name
        }
        'Noun' {
            ## TODO
            ## TODO: [workitem:9]
        }
        'Verb' {
            $TabExpansionHasOutput.Value = $true
            Get-Verb "$Argument*" | Select-Object -ExpandProperty Verb | Sort-Object
        }
    }
}.GetNewClosure()

## Reset-ComputerMachinePassword
Register-TabExpansion "Reset-ComputerMachinePassword" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Server' {
            if ($Argument -match "^\w") {
                $TabExpansionHasOutput.Value = $true
                Get-TabExpansion "$Argument*" Computer | Select-Object -ExpandProperty "Text"
            }
        }
    }
}.GetNewClosure()

## ComputerRestore
& {
    $ComputerRestoreHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Drive' {
                $TabExpansionHasOutput.Value = $true
                Get-PSDrive -PSProvider FileSystem "$Argument*" | New-TabItem -Value {$_.Root} -Text {$_.Root} -Type Drive
            }
        }
    }.GetNewClosure()
    
    $ComputerRestorePointHandler = {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'RestorePoint' {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                ## TODO: Display more info
                ## TODO: [workitem:10]
                foreach ($Point in Get-ComputerRestorePoint -EA Stop) {
                    $Text = "{0}: {1}" -f ([String]$Point.SequenceNumber),[DateTime]::ParseExact($Point.CreationTime, "yyyyMMddHHmmss.ffffff-000", $null)
                    New-TabItem -Value $Point.SequenceNumber -Text $Text -Type ComputerRestorePoint
                }
            }
        }
    }.GetNewClosure()

    Register-TabExpansion "Disable-ComputerRestore" $ComputerRestoreHandler -Type Command
    Register-TabExpansion "Enable-ComputerRestore" $ComputerRestoreHandler -Type Command
    Register-TabExpansion "Get-ComputerRestorePoint" $ComputerRestorePointHandler -Type Command
    Register-TabExpansion "Restore-Computer" $ComputerRestorePointHandler -Type Command
}

## Counter
& {
    $CounterHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Counter' {
                $TabExpansionHasOutput.Value = $true
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                Get-Counter -ListSet * @Parameters | Select-Object -ExpandProperty PathsWithInstances | 
                    Where-Object {$_ -like "*$Argument*"} | Sort-Object
            }
            'ListSet' {
                $TabExpansionHasOutput.Value = $true
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                Get-Counter -ListSet "$Argument*" @Parameters | Select-Object -ExpandProperty CounterSetName
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Get-Counter" $CounterHandler -Type "Command"
    Register-TabExpansion "Import-Counter" $CounterHandler -Type "Command"
}

## Event
& {
    $GetEventHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'SourceIdentifier' {
                $TabExpansionHasOutput.Value = $true
                Get-Event "$Argument*" | Select-Object -ExpandProperty SourceIdentifier | Sort-Object
            }
            'EventIdentifier' {
                $TabExpansionHasOutput.Value = $true
                Get-Event | Select-Object -ExpandProperty EventIdentifier
            }
        }
    }.GetNewClosure()
    $EventHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Class' {
                $TabExpansionHasOutput.Value = $true
                ## TODO: escape special characters?
                Get-TabExpansion "$Argument*" WMI | Select-Object -ExpandProperty Name
            }
            'EventName' {
                if ($Context.OtherParameters["InputObject"]) {
                    $TabExpansionHasOutput.Value = $true
                    Invoke-Expression $Context.OtherParameters["InputObject"] | Get-Member | 
                        Where-Object {$_.MemberType -eq "Event" -and $_.Name -like "$Argument*"} | Select-Object -ExpandProperty Name
                }
            }
            'Namespace' {
                $TabExpansionHasOutput.Value = $true
                if ($Argument -notlike "ROOT\*") {
                    $Argument = "ROOT\$Argument"
                }
                if ($Context.OtherParameters["ComputerName"]) {
                    $ComputerName = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                } else {
                    $ComputerName = "."
                }
                
                $ParentNamespace = $Argument -replace '\\[^\\]*$'
                $Namespaces = New-Object System.Management.ManagementClass "\\$ComputerName\${ParentNamespace}:__NAMESPACE"
                $Namespaces = foreach ($Namespace in $Namespaces.PSBase.GetInstances()) {"{0}\{1}" -f $Namespace.__NameSpace,$Namespace.Name}
                $Namespaces | Where-Object {$_ -like "$Argument*"} | Sort-Object
            }
            'SourceIdentifier' {
                ## TODO:
                ## TODO: [workitem:11]
            }
        }
    }.GetNewClosure()
    
    ## TODO: Needs work
    Register-TabExpansion "Get-Event" $GetEventHandler -Type "Command"
    Register-TabExpansion "Get-EventSubscriber" $EventHandler -Type "Command"
    Register-TabExpansion "New-Event" $EventHandler -Type "Command"
    Register-TabExpansion "Register-ObjectEvent" $EventHandler -Type "Command"
    Register-TabExpansion "Register-EngineEvent" $EventHandler -Type "Command"
    Register-TabExpansion "Register-WmiEvent" $EventHandler -Type "Command"
    Register-TabExpansion "Remove-Event" $EventHandler -Type "Command"
    Register-TabExpansion "Unregister-Event" $EventHandler -Type "Command"
    Register-TabExpansion "Wait-Event" $EventHandler -Type "Command"
}

## EventLog
& {
    $EventLogHandler = {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Category' {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                $Categories = (
                    @{'Id'='0';'Name'='None'},
                    @{'Id'='1';'Name'='Devices'},
                    @{'Id'='2';'Name'='Disk'},
                    @{'Id'='3';'Name'='Printers'},
                    @{'Id'='4';'Name'='Services'},
                    @{'Id'='5';'Name'='Shell'},
                    @{'Id'='6';'Name'='System Event'},
                    @{'Id'='7';'Name'='Network'}
                )
                $Categories | Where-Object {$_.Name -like "$Argument*"} |
                    New-TabItem -Value {$_.Id} -Text {$_.Name} -Type EventLogCategory
            }
            'LogName' {
                $TabExpansionHasOutput.Value = $true
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                Get-EventLog -List -AsString @Parameters | Where-Object {$_ -like "$Argument*"} |
                    New-TabItem -Value {$_} -Text {$_} -Type EventLog
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Clear-EventLog" $EventLogHandler -Type "Command"
    Register-TabExpansion "Get-EventLog" $EventLogHandler -Type "Command"
    Register-TabExpansion "Limit-EventLog" $EventLogHandler -Type "Command"
    Register-TabExpansion "Remove-EventLog" $EventLogHandler -Type "Command"
    Register-TabExpansion "Write-EventLog" $EventLogHandler -Type "Command"
}

## Get-Help
& {
    $HelpHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                if ($Argument -like "about_*") {
                    $Commands = Get-Help "$Argument*" | Select-Object -ExpandProperty Name
                    if ($Commands) {
                        $TabExpansionHasOutput.Value = $true
                        $Commands
                    }
                } else {
                    $Commands = Get-Command "$Argument*" -CommandType Function,Filter,Cmdlet,ExternalScript | Select-Object -ExpandProperty Name
                    if ($Commands) {
                        $TabExpansionHasOutput.Value = $true
                        $Commands
                    }
                }
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Get-Help" $HelpHandler -Type "Command"
    Register-TabExpansion "help" $HelpHandler -Type "Command"
}

## Get-HotFix
Register-TabExpansion "Get-HotFix" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Id' {
            $TabExpansionHasOutput.Value = $true
            $Parameters = @{}
            if ($Context.OtherParameters["ComputerName"]) {
                $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
            }
            Get-HotFix @Parameters | Where-Object {$_.HotFixID -like "$Argument*"} | Select-Object -ExpandProperty HotFixID
        }
    }
}.GetNewClosure()

## ItemProperty
& {
    $ItemPropertyHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                $TabExpansionHasOutput.Value = $true
                $Path = "."
                if ($Context.OtherParameters["Path"]) {
                    $Path = Resolve-TabExpansionParameterValue $Context.OtherParameters["Path"]
                }
                Get-ItemProperty -Path $Path -Name "$Argument*" | Get-Member | Where-Object {
                    (("Property","NoteProperty") -contains $_.MemberType) -and
                    (("PSChildName","PSDrive","PSParentPath","PSPath","PSProvider") -notcontains $_.Name)
                } | Select-Object -ExpandProperty Name -Unique
            }
        }
    }.GetNewClosure()

    Register-TabExpansion "Clear-ItemProperty" $ItemPropertyHandler -Type "Command"
    Register-TabExpansion "Copy-ItemProperty" $ItemPropertyHandler -Type "Command"
    Register-TabExpansion "Get-ItemProperty" $ItemPropertyHandler -Type "Command"
    Register-TabExpansion "Move-ItemProperty" $ItemPropertyHandler -Type "Command"
    Register-TabExpansion "Remove-ItemProperty" $ItemPropertyHandler -Type "Command"
    Register-TabExpansion "Rename-ItemProperty" $ItemPropertyHandler -Type "Command"
    Register-TabExpansion "Set-ItemProperty" $ItemPropertyHandler -Type "Command"
}

## Job
& {
    $JobHandler = {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Id' {
                $TabExpansionHasOutput.Value = $true
                Get-Job | Select-Object -ExpandProperty Id
            }
            'InstanceId' {
                $TabExpansionHasOutput.Value = $true
                Get-Job | Select-Object -ExpandProperty InstanceId
            }
            'Location' {
                ## TODO: Receive-Job
                ## TODO: [workitem:12]
            }
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-Job -Name "$Argument*" | Select-Object -ExpandProperty Name
            }
            'Job' {
                if ($Argument -notlike '$*') {
                    $TabExpansionHasOutput.Value = $true
                    $QuoteSpaces.Value = $false
                    foreach ($Job in Get-Job -Name "$Argument*") {'(Get-Job "{0}")' -f $Job.Name}
                }
            }
        }
    }.GetNewClosure()

    Register-TabExpansion "Get-Job" $JobHandler -Type "Command"
    Register-TabExpansion "Receive-Job" $JobHandler -Type "Command"
    Register-TabExpansion "Remove-Job" $JobHandler -Type "Command"
    Register-TabExpansion "Stop-Job" $JobHandler -Type "Command"
    Register-TabExpansion "Wait-Job" $JobHandler -Type "Command"
}

## Get-Module
Register-TabExpansion "Get-Module" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Name' {
            $Parameters = @{}
            if ($Context.OtherParameters["All"]) {
                $Parameters["All"] = $true
            }
            if ($Context.OtherParameters["ListAvailable"]) {
                $Parameters["ListAvailable"] = $true
            }
            $Modules = @(Get-Module "$Argument*" @Parameters | Sort-Object Name)
            if ($Modules.Count -gt 0) {
                $TabExpansionHasOutput.Value = $true
                $Modules | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Module
            }
        }
    }
}.GetNewClosure()

## Import-Module
Register-TabExpansion "Import-Module" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Name' {
            if ($Argument -notmatch '^\.') {
                $Modules = @(Find-Module "$Argument*" | Sort-Object BaseName)
                if ($Modules.Count -gt 0) {
                    $TabExpansionHasOutput.Value = $true
                    $Modules | New-TabItem -Value {$_.BaseName} -Text {$_.BaseName} -Type Module
                }
            }
        }
    }
}

## Remove-Module
Register-TabExpansion "Remove-Module" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Name' {
            $TabExpansionHasOutput.Value = $true
            Get-Module "$Argument*" | Select-Object -ExpandProperty Name | Sort-Object |
                New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Module
        }
    }
}.GetNewClosure()

## Group-Object
Register-TabExpansion "Group-Object" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Property' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
            }
        }
    }
}

## New-Object
Register-TabExpansion "New-Object" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'ComObject' {
            ## TODO: Maybe cache these like we do with .NET types and WMI object names?
            ## TODO: [workitem:13]
            $TabExpansionHasOutput.Value = $true
            Get-TabExpansion "$Argument*" COM | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type COMObject
        }
        'TypeName' {
            if ($Argument -notmatch '^\.') {
                ## TODO: Find way to differentiate namespaces from types
                $TabExpansionHasOutput.Value = $true
                $Dots = $Argument.Split(".").Count - 1
                $res = @()
                $res += $dsTabExpansionDatabase.Tables['Types'].Select("NS like '$Argument*' and DC = $($Dots + 1)") |
                    Select-Object -Unique -ExpandProperty NS | New-TabItem -Value {$_} -Text {"$_."} -Type Namespace
                $res += $dsTabExpansionDatabase.Tables['Types'].Select("NS like 'System.$Argument*' and DC = $($Dots + 2)") |
                    Select-Object -Unique -ExpandProperty NS | New-TabItem -Value {$_} -Text {"$_."} -Type Namespace
                if ($Dots -gt 0) {
                    $res += $dsTabExpansionDatabase.Tables['Types'].Select("Name like '$Argument*' and DC = $Dots") |
                        Select-Object -ExpandProperty Name | New-TabItem -Value {$_} -Text {$_} -Type Type
                    $res += $dsTabExpansionDatabase.Tables['Types'].Select("Name like 'System.$Argument*' and DC = $($Dots + 1)") |
                        Select-Object -ExpandProperty Name | New-TabItem -Value {$_} -Text {$_} -Type Type
                }
                $res | Where-Object {$_}
            }
        }
    }
}

## Select-Object
Register-TabExpansion "Select-Object" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Property' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
            }
        }
    }
}

## Sort-Object
Register-TabExpansion "Sort-Object" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Property' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Expression={$_.}}'
                '@{Expression={$_.}; Ascending=$true}'
                '@{Expression={$_.}; Descending=$true}'
            }
        }
    }
}

## Out-Printer
Register-TabExpansion "Out-Printer" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Name' {
            $TabExpansionHasOutput.Value = $true
            Get-WMIObject Win32_Printer -Filter "Name LIKE '$Argument%'" | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Printer
        }
    }
}.GetNewClosure()

## Process
& {
    $ProcessHandler = {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Id' {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                if ($Argument -match '^[0-9]+$') {
                    Get-Process | Where-Object {$_.Id.ToString() -like "$Argument*"} |
                        New-TabItem -Value {$_.Id} -Text {"{0:-4} {1}" -f ([String]$_.Id),$_.Name} -Type Process
                } else {
                    Get-Process | Where-Object {$_.Name -like "$Argument*"} |
                        New-TabItem -Value {$_.Id} -Text {"{0:-4} {1}" -f ([String]$_.Id),$_.Name} -Type Process
                }
            }
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-Process -Name "$Argument*" | Get-Unique | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Process
            }
        }
    }.GetNewClosure()
    
    $GetProcessHandler = {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Id' {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                if ($Argument -match '^[0-9]+$') {
                    Get-Process @Parameters | Where-Object {$_.Id.ToString() -like "$Argument*"} |
                        New-TabItem -Value {$_.Id} -Text {"{0:-4} <# {1} #>" -f ([String]$_.Id),$_.Name} -Type Process
                } else {
                    Get-Process @Parameters | Where-Object {$_.Name -like "$Argument*"} |
                        New-TabItem -Value {$_.Id} -Text {"{0:-4} <# {1} #>" -f ([String]$_.Id),$_.Name} -Type Process
                }
            }
            'Name' {
                $TabExpansionHasOutput.Value = $true
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                Get-Process -Name "$Argument*" @Parameters | Get-Unique | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Process
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Debug-Process" $ProcessHandler -Type "Command"
    Register-TabExpansion "Get-Process" $GetProcessHandler -Type "Command"
    Register-TabExpansion "Stop-Process" $ProcessHandler -Type "Command"
    Register-TabExpansion "Wait-Process" $ProcessHandler -Type "Command"
}

## PSBreakpoint
& {
    $PSBreakpointHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Breakpoint' {
                ## TODO:
                ## TODO: [workitem:15]
            }
            'Command' {
                ## TODO:
                ## TODO: [workitem:15]
            }
            'Id' {
                ## TODO:
                Get-PSBreakpoint | Select-Object -ExpandProperty Id
            }
            'Line' {
                ## TODO:
                ## TODO: [workitem:15]
            }
            'Script' {
                ## TODO: Display relative paths
                $Scripts = Get-ChildItem "$Argument*" -Include *.ps1 | Select-Object -ExpandProperty FullName
                if ($Scripts) {
                    $TabExpansionHasOutput.Value = $true
                    $Scripts
                }
            }
            'Variable' {
                if ($Argument -notlike '$*') {
                    $TabExpansionHasOutput.Value = $true
                    Get-Variable "$Argument*" -Scope Global | Select-Object -ExpandProperty Name
                }
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Disable-PSBreakpoint" $PSBreakpointHandler -Type "Command"
    Register-TabExpansion "Enable-PSBreakpoint" $PSBreakpointHandler -Type "Command"
    Register-TabExpansion "Get-PSBreakpoint" $PSBreakpointHandler -Type "Command"
    Register-TabExpansion "Set-PSBreakpoint" $PSBreakpointHandler -Type "Command"
}

## PSDrive
& {
    $PSDriveHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-PSDrive "$Argument*" | Select-Object -ExpandProperty Name
            }
        }
    }.GetNewClosure()
    $NewPSDriveHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Scope' {
                $TabExpansionHasOutput.Value = $true
                "Global","Local","Script","0"
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Get-PSDrive" $PSDriveHandler -Type "Command"
    Register-TabExpansion "New-PSDrive" $NewPSDriveHandler -Type "Command"
    Register-TabExpansion "Remove-PSDrive" $PSDriveHandler -Type "Command"
}

## PSSession
& {
    $PSSessionHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'ConfigurationName' {
                ## TODO:  But can we?
            }
            'Id' {
                $TabExpansionHasOutput.Value = $true
                Get-PSSession | Select-Object -ExpandProperty Id
            }
            'InstanceId' {
                $TabExpansionHasOutput.Value = $true
                Get-PSSession | Where-Object {$_.InstanceId -like "$Argument*"} | Select-Object -ExpandProperty InstanceId
            }
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-PSSession -Name "$Argument*" | Select-Object -ExpandProperty Name
            }
        }
    }.GetNewClosure()
    $ImportPSSessionHandler = {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'CommandName' {
                ## TODO:
                ## TODO: [workitem:16]
            }
            'FormatTypeName' {
                ## TODO:
                ## TODO: [workitem:16]
            }
            'Module' {
                ## TODO: Grab from session instead?
                $TabExpansionHasOutput.Value = $true
                (Get-Module -ListAvailable "$Argument*") + (Get-PSSnapin "$Argument*") | Select-Object -ExpandProperty Name | Sort-Object
            }
            'Session' {
                if ($Argument -notlike '$*') {
                    $TabExpansionHasOutput.Value = $true
                    $QuoteSpaces.Value = $false
                    Get-PSSession -Name "$Argument*" | ForEach-Object {'(Get-PSSession -Name "{0}")' -f $_.Name}
                }
            }
        }
    }.GetNewClosure()

    Register-TabExpansion "Invoke-Command" $PSSessionHandler -Type "Command"  ## if we can get other parameters
    Register-TabExpansion "Enter-PSSession" $PSSessionHandler -Type "Command"
    Register-TabExpansion "Export-PSSession" $ImportPSSessionHandler -Type "Command"
    Register-TabExpansion "Get-PSSession" $PSSessionHandler -Type "Command"
    Register-TabExpansion "Import-PSSession" $ImportPSSessionHandler -Type "Command"
    Register-TabExpansion "Remove-PSSession" $PSSessionHandler -Type "Command"
}

## PSSessionConfiguration
& {
    $PSSessionConfigurationHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'ConfigurationTypeName' {
                ## TODO: Find way to differentiate namespaces from types
                $TabExpansionHasOutput.Value = $true
                $Dots = $Argument.Split(".").Count - 1
                $res = @()
                $res += $dsTabExpansionDatabase.Tables['Types'].Select("ns like '$Argument*' and dc = $($Dots + 1)") |
                    Select-Object -Unique -ExpandProperty ns
                if ($Dots -gt 0) {
                    $res += $dsTabExpansionDatabase.Tables['Types'].Select("name like '$Argument*' and dc = $Dots") |
                        Select-Object -ExpandProperty Name
                }
                $res
            }
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-PSSessionConfiguration "$Argument*" | Select-Object -ExpandProperty Name
            }
        }
    }

    Register-TabExpansion "Disable-PSSessionConfiguration" $PSSessionConfigurationHandler -Type "Command"
    Register-TabExpansion "Enable-PSSessionConfiguration" $PSSessionConfigurationHandler -Type "Command"
    Register-TabExpansion "Get-PSSessionConfiguration" $PSSessionConfigurationHandler -Type "Command"
    Register-TabExpansion "Register-PSSessionConfiguration" $PSSessionConfigurationHandler -Type "Command"
    Register-TabExpansion "Set-PSSessionConfiguration" $PSSessionConfigurationHandler -Type "Command"
    Register-TabExpansion "Unregister-PSSessionConfiguration" $PSSessionConfigurationHandler -Type "Command"
}

## Add-PSSnapin
Register-TabExpansion "Add-PSSnapin" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Name' {
            $TabExpansionHasOutput.Value = $true
            $Loaded = @(Get-PSSnapin)
            Get-PSSnapin "$Argument*" -Registered | Where-Object {$Loaded -notcontains $_} |
                Select-Object -ExpandProperty Name | Sort-Object
        }
    }
}.GetNewClosure()

## Get-PSSnapin
Register-TabExpansion "Get-PSSnapin" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Name' {
            $TabExpansionHasOutput.Value = $true
            $Parameters = @{"ErrorAction" = "SilentlyContinue"}
            if ($Context.OtherParameters["Registered"]) {
                $Parameters["Registered"] = $true
            }
            Get-PSSnapin "$Argument*" @Parameters | Select-Object -ExpandProperty Name | Sort-Object
        }
    }
}.GetNewClosure()

## Service
& {
    $ServiceHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'DisplayName' {
                $TabExpansionHasOutput.Value = $true
                Get-Service -DisplayName "*$Argument*" | Select-Object -ExpandProperty DisplayName
            }
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-Service -Name "$Argument*" | Select-Object -ExpandProperty Name
            }
        }
    }.GetNewClosure()
    
    $GetServiceHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'DisplayName' {
                $TabExpansionHasOutput.Value = $true
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                Get-Service -DisplayName "*$Argument*" @Parameters | Select-Object -ExpandProperty DisplayName
            }
            'Name' {
                $TabExpansionHasOutput.Value = $true
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                Get-Service -Name "$Argument*" @Parameters | Select-Object -ExpandProperty Name
            }
        }
    }.GetNewClosure()
    
    $SetServiceHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                $TabExpansionHasOutput.Value = $true
                $Parameters = @{}
                if ($Context.OtherParameters["ComputerName"]) {
                    $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                }
                Get-Service -Name "$Argument*" @Parameters | Select-Object -ExpandProperty Name
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Get-Service" $GetServiceHandler -Type "Command"
    Register-TabExpansion "Restart-Service" $ServiceHandler -Type "Command"
    Register-TabExpansion "Resume-Service" $ServiceHandler -Type "Command"
    Register-TabExpansion "Set-Service" $SetServiceHandler -Type "Command"
    Register-TabExpansion "Start-Service" $ServiceHandler -Type "Command"
    Register-TabExpansion "Stop-Service" $ServiceHandler -Type "Command"
    Register-TabExpansion "Suspend-Service" $ServiceHandler -Type "Command"
}

## Set-StrictMode
Register-TabExpansion "Set-StrictMode" -Type Command {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Version' {
            $TabExpansionHasOutput.Value = $true
            "1.0","2.0","Latest" | Where-Object {$_ -like "$Argument*"}
        }
    }
}.GetNewClosure()

## TraceSource
& {
    $TraceSourceHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-TraceSource "$Argument*" | Select-Object -ExpandProperty Name
            }
            'RemoveListener' {
                $TabExpansionHasOutput.Value = $true
                "Host","Debug","*" | Where-Object {$_ -like "$Argument*"}
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Get-TraceSource" $TraceSourceHandler -Type "Command"
    Register-TabExpansion "Set-TraceSource" $TraceSourceHandler -Type "Command"
    Register-TabExpansion "Trace-Command" $TraceSourceHandler -Type "Command"
}

## Get-Verb
Register-TabExpansion "Get-Verb" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'Verb' {
            $TabExpansionHasOutput.Value = $true
            Get-Verb "$Argument*" | Select-Object -ExpandProperty Verb | Sort-Object
        }
    }
}.GetNewClosure()

## Variable
& {
    $VariableHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                if ($Argument -notlike '$*') {
                    $TabExpansionHasOutput.Value = $true
                    Get-Variable "$Argument*" -Scope "Global" | Select-Object -ExpandProperty Name
                }
            }
            'Scope' {
                $TabExpansionHasOutput.Value = $true
                "Global","Local","Script" | Where-Object {$_ -like "$Argument*"}
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Clear-Variable" $VariableHandler -Type "Command"
    Register-TabExpansion "Get-Variable" $VariableHandler -Type "Command"
    Register-TabExpansion "Remove-Variable" $VariableHandler -Type "Command"
    Register-TabExpansion "Set-Variable" $VariableHandler -Type "Command"
}

## Get-WinEvent
Register-TabExpansion "Get-WinEvent" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    $Parameters = @{"ErrorAction" = "SilentlyContinue"}
    if ($Context.OtherParameters["ComputerName"]) {
        $Parameters["ComputerName"] = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
    }
    switch -exact ($Context.Parameter) {
        'FilterHashTable' {
            $TabExpansionHasOutput.Value = $true
            $QuoteSpaces.Value = $false
            '@{LogName="*"}'
            '@{ProviderName="*"}'
            '@{Keywords=""}'
            '@{ID=""}'
            '@{Level=""}'
        }
        'ListLog' {
            $TabExpansionHasOutput.Value = $true
            ## TODO: Make it easier to access detailed Microsoft-* logs?
            Get-WinEvent -ListLog "$Argument*" @Parameters | Select-Object -ExpandProperty LogName
        }
        'ListProvider' {
            $TabExpansionHasOutput.Value = $true
            Get-WinEvent -ListProvider "$Argument*" @Parameters | Select-Object -ExpandProperty Name #| Sort-Object
        }
        'LogName' {
            $TabExpansionHasOutput.Value = $true
            ## TODO: Make it easier to access detailed Microsoft-* logs?
            Get-WinEvent -ListLog "$Argument*" @Parameters | Select-Object -ExpandProperty LogName
        }
        'ProviderName' {
            $TabExpansionHasOutput.Value = $true
            Get-WinEvent -ListProvider "$Argument*" @Parameters | Select-Object -ExpandProperty Name #| Sort-Object
        }
    }
}.GetNewClosure()

## WMI
& {
    $WmiObjectHandler = {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Class' {
                $TabExpansionHasOutput.Value = $true
                ## TODO: escape special characters?
                Get-TabExpansion "$Argument*" WMI | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type WMIClass
            }
            'Locale' {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::InstalledWin32Cultures) |
                    Where-Object {$_.Name -like "$Argument*"} | Sort-Object -Property Name |
                        New-TabItem -Value {$_.LCID} -Text {$_.Name} -Type Locale
            }
            'Name' {
                ## TODO: ??? (Method Name)
                ## TODO: [workitem:17]
            }
            'Namespace' {
                $TabExpansionHasOutput.Value = $true
                if ($Argument -notlike "ROOT\*") {
                    $Argument = "ROOT\$Argument"
                }
                if ($Context.OtherParameters["ComputerName"]) {
                    $ComputerName = Resolve-TabExpansionParameterValue $Context.OtherParameters["ComputerName"]
                } else {
                    $ComputerName = "."
                }
                
                $ParentNamespace = $Argument -replace '\\[^\\]*$'
                $Namespaces = New-Object System.Management.ManagementClass "\\$ComputerName\${ParentNamespace}:__NAMESPACE"
                $Namespaces = foreach ($Namespace in $Namespaces.PSBase.GetInstances()) {"{0}\{1}" -f $Namespace.__NameSpace,$Namespace.Name}
                $Namespaces | Where-Object {$_ -like "$Argument*"} | Sort-Object | New-TabItem -Value {$_} -Text {$_} -Type WMINamespace
            }
            'Path' {
                ## TODO: ???
                ## TODO: [workitem:17]
            }
            'Property' {
                ## TODO: ???
                ## TODO: [workitem:17]
            }
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Get-WmiObject" $WmiObjectHandler -Type "Command"
    Register-TabExpansion "Invoke-WmiMethod" $WmiObjectHandler -Type "Command"
    Register-TabExpansion "Register-WmiEvent" $WmiObjectHandler -Type "Command"
    Register-TabExpansion "Remove-WmiObject" $WmiObjectHandler -Type "Command"
    Register-TabExpansion "Set-WmiInstance" $WmiObjectHandler -Type "Command"
}

## WSMan & WSManInstance & WSManAction
& {
    ## TODO: [workitem:18]
}

## Format-Custom
Register-TabExpansion "Format-Custom" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'GroupBy' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
                '@{Name=""; Expression={$_.}; FormatString=""}'
            }
        }
        'Property' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Expression={$_.}}'
                '@{Expression={$_.}; Depth=3}'
            }
        }
        'View' {
            ## TODO: Need to figure out what type of object will be coming in
            ## TODO: [workitem:19]
        }
    }
}.GetNewClosure()

## Format-List
Register-TabExpansion "Format-List" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'GroupBy' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
                '@{Name=""; Expression={$_.}; FormatString=""}'
            }
        }
        'Property' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
                '@{Name=""; Expression={$_.}; FormatString=""}'
            }
        }
        'View' {
            ## TODO: Need to figure out what type of object will be coming in
            ## TODO: [workitem:19]
        }
    }
}.GetNewClosure()

## Format-Table
Register-TabExpansion "Format-Table" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'GroupBy' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
                '@{Name=""; Expression={$_.}; FormatString=""}'
            }
        }
        'Property' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
                '@{Name=""; Expression={$_.}; FormatString=""}'
                '@{Name=""; Expression={$_.}; Width=9}'
                '@{Name=""; Expression={$_.}; Width=9; Alignment="Left"}'
                '@{Name=""; Expression={$_.}; Width=9; Alignment="Center"}'
                '@{Name=""; Expression={$_.}; Width=9; Alignment="Right"}'
            }
        }
        'View' {
            ## TODO: Need to figure out what type of object will be coming in
            ## TODO: [workitem:19]
        }
    }
}.GetNewClosure()

## Format-Wide
Register-TabExpansion "Format-Wide" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
    $Argument = $Context.Argument
    switch -exact ($Context.Parameter) {
        'GroupBy' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Name=""; Expression={$_.}}'
                '@{Name=""; Expression={$_.}; FormatString=""}'
            }
        }
        'Property' {
            if ($Argument -like "@*") {
                $TabExpansionHasOutput.Value = $true
                $QuoteSpaces.Value = $false
                '@{Expression={$_.}}'
                '@{Expression={$_.}; FormatString=""}'
            }
        }
        'View' {
            ## TODO: Need to figure out what type of object will be coming in
            ## TODO: [workitem:19]
        }
    }
}.GetNewClosure()

## Function
Register-TabExpansion "function" -Type "Command" {
    param($Context, [ref]$TabExpansionHasOutput)
    $Argument = $Context.Argument
    if ($Context.PositionalParameters -eq 0) {
        $TabExpansionHasOutput.Value = $true
        if ($Argument -match '^[a-zA-Z]*$') {
            Get-Verb "$Argument*" | Select-Object -ExpandProperty Verb | Sort-Object
        }
    }
}.GetNewClosure()


#########################
## Parameter handlers
#########################

## -ComputerName and -Server
& {
    $ComputerNameHandler =  {
        param($Argument, [ref]$TabExpansionHasOutput)
        if ($Argument -notmatch '^\$') {
            $TabExpansionHasOutput.Value = $true
            Get-TabExpansion "$Argument*" Computer | New-TabItem -Value {$_.Text} -Text {$_.Text} -Type Computer
        }
    }.GetNewClosure()

    Register-TabExpansion "ComputerName" $ComputerNameHandler -Type Parameter
    Register-TabExpansion "Server" $ComputerNameHandler -Type Parameter
}

## Parameters that take the name of a variable
& {
    $VariableHandler = {
        param($Argument, [ref]$TabExpansionHasOutput)
        if ($Argument -notlike '^\$') {
            $TabExpansionHasOutput.Value = $true
            Get-Variable "$Argument*" -Scope Global | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Variable
        }
    }.GetNewClosure()
    
    Register-TabExpansion "ErrorVariable" $VariableHandler -Type Parameter
    Register-TabExpansion "OutVariable" $VariableHandler -Type Parameter
    Register-TabExpansion "Variable" $VariableHandler -Type Parameter
    Register-TabExpansion "WarningVariable" $VariableHandler -Type Parameter
}

## Parameters that take the name of a culture
& {
    $CultureHandler = {
        param($Argument, [ref]$TabExpansionHasOutput)
        if ($Argument -notlike '^\$') {
            $TabExpansionHasOutput.Value = $true
            [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::InstalledWin32Cultures) |
                Where-Object {$_.Name -like "$Argument*"} | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Culture | Sort-Object Name
        }
    }.GetNewClosure()
    
    Register-TabExpansion "Culture" $CultureHandler -Type Parameter
    Register-TabExpansion "UICulture" $CultureHandler -Type Parameter
}

## -PSDrive
Register-TabExpansion "PSDrive" -Type Parameter {
    param($Argument, [ref]$TabExpansionHasOutput)
    if ($Argument -notlike '^\$') {
        $TabExpansionHasOutput.Value = $true
        Get-PSDrive "$Argument*" | Select-Object -ExpandProperty Name
    }
}.GetNewClosure()

## -PSProvider
Register-TabExpansion "PSProvider" -Type Parameter {
    param($Argument, [ref]$TabExpansionHasOutput)
    if ($Argument -notlike '^\$') {
        $TabExpansionHasOutput.Value = $true
        Get-PSProvider "$Argument*" | Select-Object -ExpandProperty Name
    }
}.GetNewClosure()


#########################
## Parameter Name handlers
#########################

## iexplore.exe
& {
    Register-TabExpansion iexplore.exe -Type ParameterName {
        param($Context, $Parameter)
        $Parameters = "-extoff","-embedding","-k","-nohome"
        $Parameters | Where-Object {$_ -like "$Parameter*"}
    }.GetNewClosure()

    Function iexploreexeparameters {
        param(
            [Switch]$extoff
            ,
            [Switch]$embedding
            ,
            [Switch]$k
            ,
            [Switch]$nohome
            ,
            [Parameter(Position = 0)]
            [String]$URL
        )
    }

    $IExploreCommandInfo = Get-Command iexploreexeparameters
    Register-TabExpansion iexplore.exe -Type CommandInfo {
        param($Context)
        $IExploreCommandInfo
    }.GetNewClosure()

    Register-TabExpansion iexplore.exe -Type Command {
        param($Context, [ref]$TabExpansionHasOutput, [ref]$QuoteSpaces)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'URL' {
                $Argument = [Regex]::Escape($Argument)
                $Favorites = Get-ChildItem "$env:USERPROFILE/Favorites/*" -Include *.url -Recurse
                $Favorites = $Favorites | Where-Object {($_.Name -match $Argument) -or ($_ | Select-String "^URL=.*$Argument")} |
                    New-TabItem -Value {($_ | Select-String "^URL=").Line -replace "^URL="} -Text {$_.Name -replace '\.url$'} -Type URL

                if ($Favorites) {
                    $TabExpansionHasOutput.Value = $true
                    $QuoteSpaces.Value = $false
                    $Favorites
                }
            }
        }
    }.GetNewClosure()
}

## powershell.exe
& {
    Register-TabExpansion powershell.exe -Type ParameterName {
        param($Context, $Parameter)
        $Parameters = "-Command","-EncodedCommand","-ExecutionPolicy","-File","-InputFormat","-NoExit","-NoLogo",
            "-NonInteractive","-NoProfile","-OutputFormat","-PSConsoleFile","-Sta","-Version","-WindowStyle"
        $Parameters | Where-Object {$_ -like "$Parameter*"}
        <#
        PowerShell[.exe] [-PSConsoleFile <file> | -Version <version>]
        [-NoLogo] [-NoExit] [-Sta] [-NoProfile] [-NonInteractive]
        [-InputFormat {Text | XML}] [-OutputFormat {Text | XML}]
        [-WindowStyle <style>] [-EncodedCommand <Base64EncodedCommand>]
        [-File <filePath> <args>] [-ExecutionPolicy <ExecutionPolicy>]
        [-Command { - | <script-block> [-args <arg-array>]
                      | <string> [<CommandParameters>] } ]
        #>
    }.GetNewClosure()

    Function powershellexeparameters {
        param(
            [String]$Command
            ,
            [String]$EncodedCommand
            ,
            [Microsoft.PowerShell.ExecutionPolicy]$ExecutionPolicy
            ,
            [String]$File
            ,
            [ValidateSet("Text","XML")]
            [String]$InputFormat
            ,
            [Switch]$NoExit
            ,
            [Switch]$NonInteractive
            ,
            [Switch]$NoLogo
            ,
            [Switch]$NoProfile
            ,
            [ValidateSet("Text","XML")]
            [String]$OutputFormat
            ,
            [String]$PSConsoleFile
            ,
            [Switch]$Sta
            ,
            [ValidateSet("1.0","2.0")]
            [String]$Version
            ,
            [ValidateSet("Normal","Minimized","Maximized","Hidden")]
            [String]$WindowStyle
        )
    }

    $PowershellCommandInfo = Get-Command powershellexeparameters
    Register-TabExpansion powershell.exe -Type CommandInfo {
        param($Context)
        $PowershellCommandInfo
    }.GetNewClosure()
}


#########################
## PowerTab function handlers
#########################

## Themes
& {
    $ThemeHandler = {
        param($Context, [ref]$TabExpansionHasOutput)
        $Argument = $Context.Argument
        switch -exact ($Context.Parameter) {
            'Name' {
                $TabExpansionHasOutput.Value = $true
                Get-ChildItem (Join-Path $PSScriptRoot "ColorThemes\Theme${Argument}*") -Include *.csv |
                    ForEach-Object {$_.Name -replace '^Theme([^\.]+)\.csv$','$1'}
            }
        }
    }

    Register-TabExpansion "Import-TabExpansionTheme" $ThemeHandler -Type Command
    Register-TabExpansion "Export-TabExpansionTheme" $ThemeHandler -Type Command
}


# SIG # Begin signature block
# MIIY1AYJKoZIhvcNAQcCoIIYxTCCGMECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUm3PkP4lID9z8gx4i/N5LC2ax
# zUmgghSGMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
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
# CSqGSIb3DQEJBDEWBBQ0GD5qkCvZ2cw/1hWQkP/jVk2RuzANBgkqhkiG9w0BAQEF
# AASCAQB4IGo5cq1fDOTuIaPQ6OYxrLlMIjCSvjnXFL85T8UB9Fxba7C+8ONSDLLj
# 7M2T6qqricsuhvaqwps5Z8kV1zNWRBqTnI6mvl1StVifxvEN6Ohj5yOHrhgkeZRq
# NEIxOS+Bdd7BRPxDRdKJ8O0wFXs28XQgsCkGq3RKaRFeaZXzjDvX21dxY1LrWSjr
# GeaDDPzGJVqbHVW5N9ctOi35sh8SCqhlL0tCeG8gr1nq7z/RWJfkTTYF2ZHuANu/
# 73dBI3Cc0lJBeThRsm9F5NCEK6PRpHXXI94IlPsrTWCtuMS/6K0SObu0+dRkHvsV
# uMB2aha70KxrP4VTr8Uf8XlmajwToYIBfzCCAXsGCSqGSIb3DQEJBjGCAWwwggFo
# AgEBMGcwUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBAhA4Jdf6
# +GGvnvSQ5ya11lrVMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0xMTEwMjgxOTUzMzdaMCMGCSqGSIb3DQEJBDEW
# BBTbARNlRa0xYZivzjAwYjSHSpGDyjANBgkqhkiG9w0BAQEFAASBgDx5WFj/MLKk
# bhYH/CxO9Mds2XAMA0dpvVf+sgQb1avrzZ4/UkSG12eySVgZIZFArsSuWlOAIHvJ
# SZnrDmUb8G8jBEDRnRxG7iGH1D2deso8QM87/w3QOvlXPiEBCRUzB5Oya4p4IB1h
# FDFI1w5QapTI1csSdsfEXsisJyGgv+wU
# SIG # End signature block
