# TabExpansionCore.ps1
#
# 


# .ExternalHelp TabExpansionCore-Help.xml
Function Invoke-TabExpansion {
	[CmdletBinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true)]
        [AllowEmptyString()]
        [String]
        $Line
        ,
        [Parameter(Position = 1, Mandatory = $true)]
        [AllowEmptyString()]
        [String]
        $LastWord
        ,
        [Switch]
        $ForceList
    )

    try {
    ## Save global errors in the script scoped Error ( doesn't appear to be used )
    # This also addresses the problem of $Error[0].<TAB> not working since it uses the script scoped $Error as well.
    $script:Error.Clear()
    $script:Error.AddRange($global:Error)
    $global:Error.Clear()
    $global:Error.AddRange($PowerTabError)

    if ($PowerTabConfig -eq $null) {
        ## Something happened to the PowerTabConfig object, recreate it
        CreatePowerTabConfig
    }

    $ParseErrors = $null
    $Tokens = [System.Management.Automation.PSParser]::Tokenize($Line, [ref]$ParseErrors)

    ## Figure out the context of this tab expansion request
    $_TokenTypes = [System.Management.Automation.PSTokenType]
    $ScopeStack = New-Object System.Collections.Stack
    $CurrentContext = New-TabContext
    $LastToken = $null
    $LastParameter = ""
    ## TODO: Save all values from a list for a parameter
    foreach ($Token in $Tokens) {
        if (($Token.Type -eq $_TokenTypes::Command) -and !($CurrentContext.Command)) {
            $CurrentContext.CommandInfo = try {Resolve-Command $Token.Content -CommandInfo -ErrorAction "Stop"} catch {}
            if ($CurrentContext.CommandInfo) {
                $CurrentContext.Command = $CurrentContext.CommandInfo.Name
            } else {
                $CurrentContext.Command = $Token.Content
            }
            $CurrentContext.isCommandMode = $true
        } elseif (($Token.Type -eq $_TokenTypes::Keyword) -and !($CurrentContext.Command) -and
                (@("function") -contains $Token.Content)) {
            $CurrentContext.Command = $Token.Content
            $CurrentContext.isCommandMode = $true
        } elseif ($Token.Type -eq $_TokenTypes::CommandParameter) {
            if ($CurrentContext.Parameter) {
                $CurrentContext.OtherParameters[$CurrentContext.Parameter] = $CurrentContext.Argument
            } elseif ($CurrentContext.Argument) {
                while ($CurrentContext.PositionalParameters.Count -le $CurrentContext.PositionalParameter) {
                    $CurrentContext.PositionalParameters += @("")
                }
                $CurrentContext.PositionalParameters[$CurrentContext.PositionalParameter] = $CurrentContext.Argument
            }
            $CurrentContext.Parameter = try {
                if ($TabExpansionCommandInfoRegistry[$CurrentContext.Command]) {
                    $ScriptBlock = $TabExpansionCommandInfoRegistry[$CurrentContext.Command]
                    $CommandInfo = & $ScriptBlock $CurrentContext
                    if ($CommandInfo) {
                        Resolve-Parameter $CommandInfo $Token.Content
                    } else {
                        Resolve-Parameter $CurrentContext.CommandInfo $Token.Content
                    }
                } else {
                    Resolve-Parameter $CurrentContext.CommandInfo $Token.Content
                }
            } catch {$Token.Content -replace '^-'}
            if (-not $CurrentContext.Parameter) {$CurrentContext.Parameter = $Token.Content -replace '^-'}
            $CurrentContext.Argument = ""
            $CurrentContext.isParameterValue = $false

            ## Check if parameter is a switch
            try {
                if ($TabExpansionCommandInfoRegistry[$CurrentContext.Command]) {
                    $ScriptBlock = $TabExpansionCommandInfoRegistry[$CurrentContext.Command]
                    $CommandInfo = & $ScriptBlock $CurrentContext
                    if ($CommandInfo) {
                        $Parameter = Resolve-Parameter $CommandInfo $CurrentContext.Parameter -ParameterInfo
                    } else {
                        ## TODO:
                        throw "foo"
                    }
                } else {
                    $Parameter = Resolve-Parameter $CurrentContext.CommandInfo $CurrentContext.Parameter -ParameterInfo
                }
                if ($Parameter.ParameterType -eq [System.Management.Automation.SwitchParameter]) {
                    $CurrentContext.OtherParameters[$CurrentContext.Parameter] = $true
                    $CurrentContext.Parameter = ""
                    $CurrentContext.Argument = ""
                    $CurrentContext.isParameterValue = $false
                }
            } catch {}
        } elseif (($Token.Type -eq $_TokenTypes::StatementSeparator) -or ($Token.Content -eq "|")) {
            $CurrentContext = New-TabContext
        } elseif (($Token.Type -eq $_TokenTypes::Operator) -and ($Token.Content -eq "=")) {
            $CurrentContext.isAssignment = $true
        } elseif ($CurrentContext.isCommandMode -and (($_TokenTypes::CommandArgument, $_TokenTypes::Member,
                $_TokenTypes::Number, $_TokenTypes::Operator, $_TokenTypes::String, $_TokenTypes::Variable) -contains $Token.Type)) {
            if (($Token.Type -eq $_TokenTypes::Operator) -and ("," -eq $Token.Content)) {
                ## TODO: Keep track of full list of values for a parameter?
            }

            if ($LastToken.EndColumn -ne $Token.StartColumn) {
                ## The parameter check is for the case of a switch appearing just befor a positional parameter:  Func -Switch tex<TAB>
                if ((-not $Token.Content.StartsWith("-")) -and (($CurrentContext.Argument -ne "") -or ($CurrentContext.Parameter -eq ""))) {
                    if ($CurrentContext.Parameter) {
                        $CurrentContext.OtherParameters[$CurrentContext.Parameter] = $CurrentContext.Argument
                    } elseif ($CurrentContext.Argument) {
                        while ($CurrentContext.PositionalParameters.Count -le $CurrentContext.PositionalParameter) {
                            $CurrentContext.PositionalParameters += @("")
                        }
                        $CurrentContext.PositionalParameters[$CurrentContext.PositionalParameter] = $CurrentContext.Argument
                    }
                    ## Found value for positional parameter
                    $CurrentContext.Parameter = ""
                    $CurrentContext.Argument = ""
                    $CurrentContext.PositionalParameter += 1
                    $CurrentContext.isParameterValue = $true
                } elseif ($Token.Content.StartsWith("-")) {
                    if ($CurrentContext.Parameter) {
                        $CurrentContext.OtherParameters[$CurrentContext.Parameter] = $CurrentContext.Argument
                    } elseif ($CurrentContext.Argument) {
                        while ($CurrentContext.PositionalParameters.Count -le $CurrentContext.PositionalParameter) {
                            $CurrentContext.PositionalParameters += @("")
                        }
                        $CurrentContext.PositionalParameters[$CurrentContext.PositionalParameter] = $CurrentContext.Argument
                    }
                    $CurrentContext.Parameter = ""
                    $CurrentContext.Argument = ""
                    $CurrentContext.isParameterValue = $false
                    continue
                }
            }

            ## Decide if this token could be handed off to parameter tab completion
            if (($_TokenTypes::CommandArgument, $_TokenTypes::Number, $_TokenTypes::String) -contains $Token.Type) {
                $CurrentContext.Argument = $Token.Content
                $CurrentContext.isParameterValue = $true
            } elseif ($Token.Type -eq $_TokenTypes::Variable) {
                $CurrentContext.Argument = '$' + $Token.Content
                $CurrentContext.isParameterValue = $true
            } elseif (($Token.Type -eq $_TokenTypes::Operator) -and ("," -eq $Token.Content)) {
                $CurrentContext.Argument = ""
                $CurrentContext.isParameterValue = $true
            } else {
                $CurrentContext.Argument = ""
                $CurrentContext.isParameterValue = $false
            }
        } elseif ($Token.Type -eq $_TokenTypes::GroupStart) {
            $ScopeStack.Push($CurrentContext)
            $CurrentContext = New-TabContext
        } elseif ($Token.Type -eq $_TokenTypes::GroupEnd) {
            $CurrentContext = $ScopeStack.Pop()
            if ($CurrentContext.Parameter) {
                $CurrentContext.OtherParameters[$CurrentContext.Parameter] = $CurrentContext.Argument
            }
            $CurrentContext.Parameter = ""
            $CurrentContext.Argument = ""
            $CurrentContext.isParameterValue = $false
        }

        $LastToken = $Token
    }
    ## Special case, last word is "@", this causes a parsing error so we don't see the token
    if ($LastWord -eq '@') {
        if (-not $CurrentContext.Parameter) {
            $CurrentContext.PositionalParameter += 1
        }
        $CurrentContext.Argument = $LastWord
        $CurrentContext.isParameterValue = $true
    }
    ## Special case, blank value for parameter
    if (-not $LastWord) {
        if ($LastToken.Content -eq ",") {
            ## Don't do anything, build on existing list
        } elseif ($CurrentContext.Argument -ne $LastWord) {
            ## Special case, blank value for positional parameter after another parameter (foo -bar "test" _)
            if ($CurrentContext.Parameter) {
                $CurrentContext.OtherParameters[$CurrentContext.Parameter] = $CurrentContext.Argument
                $CurrentContext.Parameter = ""
            } elseif ($CurrentContext.Argument) {
                while ($CurrentContext.PositionalParameters.Count -le $CurrentContext.PositionalParameter) {
                    $CurrentContext.PositionalParameters += @("")
                }
                $CurrentContext.PositionalParameters[$CurrentContext.PositionalParameter] = $CurrentContext.Argument
            }
            if ($CurrentContext.Command) {
                $CurrentContext.PositionalParameter += 1
            }
        } else {
            if (-not $CurrentContext.Parameter) {
                $CurrentContext.PositionalParameter += 1
            }
        }

        if ($CurrentContext.Command) {
            $CurrentContext.Argument = ""
            $CurrentContext.isParameterValue = $true
        }
    }

    ## Resolve name of positional parameter
    if ((-not $CurrentContext.Parameter) -and ($LastWord -ne "-")) {
        try {
            $CurrentContext = Resolve-PositionalParameter $CurrentContext
        } catch {}
    }

    ## Add additional context information
    Add-Member -InputObject $CurrentContext -Name Line -Value $Line -MemberType NoteProperty
    Add-Member -InputObject $CurrentContext -Name LastWord -Value $LastWord -MemberType NoteProperty
    Add-Member -InputObject $CurrentContext -Name LastToken -Value $LastToken.Type -MemberType NoteProperty
    ##  Special debug logging
    if ($PowerTabLog) {
        $CurrentContext | Select-Object Line,LastWord,LastToken,Command,Parameter,Argument,PositionalParameter,
            PositionalParameters,OtherParameters,isCommandMode,isAssignment,isParameterValue | Out-String |
            Add-Content (Join-Path $env:USERPROFILE "PowerTab.log")
    }

    ## Indicate we are busy
    Invoke-TabActivityIndicator

    try {
        ## Detect DoubleTab if enabled
        if ($PowerTabConfig.DoubleTabEnabled) {
            Start-Sleep -m 400
            $DoubleTab = ($Host.UI.RawUI.KeyAvailable)
        } else {
            $DoubleTab = $PowerTabConfig.DoubleTabLock  
        }

        ## Check DoubleTab and set selection handler
        if ($DoubleTab) {
            $SelectionHandler = $PowerTabConfig.AlternateHandler
        } else {
            $SelectionHandler = $PowerTabConfig.DefaultHandler
        }

        ## Resolve internal (no prefix) and fully qualified command names
        $FullCommandName = ""
        try {
            $InternalCommand = Resolve-InternalCommandName $CurrentContext.CommandInfo
            $InternalCommandName = $InternalCommand.InternalName
            if ($InternalCommand.Module) {
                $FullCommandName = $InternalCommand.Module.Name + "\" + $InternalCommand.InternalName
            }
        } catch {
            $InternalCommandName = $CurrentContext.Command
        }

        [Bool]$TabExpansionHasOutput = $false
        [Bool]$QuoteSpaces = $true
        [Object[]]$PossibleValues = @()
        if ($CurrentContext.isParameterValue) {
            ## Tab complete parameter value

            ## Command registry
            if ((-not $TabExpansionHasOutput) -and $TabExpansionCommandRegistry[$FullCommandName]) {
                $ScriptBlock = $TabExpansionCommandRegistry[$FullCommandName]
                $PossibleValues = & $ScriptBlock $CurrentContext ([ref]$TabExpansionHasOutput) ([ref]$QuoteSpaces)
            }
            if ((-not $TabExpansionHasOutput) -and $TabExpansionCommandRegistry[$InternalCommandName]) {
                $ScriptBlock = $TabExpansionCommandRegistry[$InternalCommandName]
                $PossibleValues = & $ScriptBlock $CurrentContext ([ref]$TabExpansionHasOutput) ([ref]$QuoteSpaces)
            }

            ## Parameter registry
            if ((-not $TabExpansionHasOutput) -and $TabExpansionParameterRegistry[$CurrentContext.Parameter]) {
                $ScriptBlock = $TabExpansionParameterRegistry[$CurrentContext.Parameter]
                $PossibleValues = & $ScriptBlock $CurrentContext.Argument ([ref]$TabExpansionHasOutput) ([ref]$QuoteSpaces)
            }

            ## Enum and ValidateSet() support
            if (-not $TabExpansionHasOutput) {
                try {
                    ## Get parameter info
                    if ($TabExpansionCommandInfoRegistry[$InternalCommandName]) {
                        $ScriptBlock = $TabExpansionCommandInfoRegistry[$InternalCommandName]
                        $CommandInfo = & $ScriptBlock $CurrentContext
                        if ($CommandInfo) {
                            $ParameterInfo = Resolve-Parameter $CommandInfo $CurrentContext.Parameter -ParameterInfo
                        } else {
                            $ParameterInfo = Resolve-Parameter $CurrentContext.CommandInfo $CurrentContext.Parameter -ParameterInfo
                        }
                    } else {
                        $ParameterInfo = Resolve-Parameter $CurrentContext.CommandInfo $CurrentContext.Parameter -ParameterInfo
                    }

                    ## Enum
                    if ($ParameterInfo.ParameterType.BaseType -eq [System.Enum]) {
                        $TabExpansionHasOutput = $true
                        $PossibleValues = [Enum]::GetNames($ParameterInfo.ParameterType) | Where-Object {$_ -like ($CurrentContext.Argument + "*")} | 
                            New-TabItem -Value {$_} -Text {$_} -Type Enum
                    }

                    ## ValidateSet
                    if (-not $TabExpansionHasOutput) {
                        $ValidateSet = $ParameterInfo.Attributes | Where-Object {$_ -is [System.Management.Automation.ValidateSetAttribute]}
                        if ($ValidateSet) {
                            $TabExpansionHasOutput = $true
                            $PossibleValues = $ValidateSet | Select-Object -ExpandProperty ValidValues |
                                Where-Object {$_ -like ($CurrentContext.Argument + "*")} | 
                                New-TabItem -Value {$_} -Text {$_} -Type Enum
                        }
                    }
                } catch {}
            }

            ## Ensure that variables get handled
            if ($PossibleValues -eq $null) {$PossibleValues = @()}
            if ($TabExpansionHasOutput -and ($PossibleValues.Count -eq 0) -and ($LastWord -match '^[@\$]')) {
                $TabExpansionHasOutput = $false
            }
        } elseif (($LastWord -match "^-") -and $CurrentContext.isCommandMode) {
            ## Tab complete parameter name

            ## Parameter name registry
            if ((-not $TabExpansionHasOutput) -and $TabExpansionParameterNameRegistry[$FullCommandName]) {
                $ScriptBlock = $TabExpansionParameterNameRegistry[$FullCommandName]
                $PossibleValues = & $ScriptBlock $CurrentContext $LastWord | New-TabItem -Value {$_} -Text {$_} -Type Parameter
                if ($PossibleValues) {
                    $TabExpansionHasOutput = $true
                }
            }
            if ((-not $TabExpansionHasOutput) -and $TabExpansionParameterNameRegistry[$InternalCommandName]) {
                $ScriptBlock = $TabExpansionParameterNameRegistry[$InternalCommandName]
                $PossibleValues = & $ScriptBlock $CurrentContext $LastWord | New-TabItem -Value {$_} -Text {$_} -Type Parameter
                if ($PossibleValues) {
                    $TabExpansionHasOutput = $true
                }
            }

            ## Command info
            if (-not $TabExpansionHasOutput) {
                if ($CurrentContext.CommandInfo) {
                    $ParameterName = $LastWord -replace "^-"
                    $PossibleValues = foreach ($Parameter in $CurrentContext.CommandInfo.Parameters.Values) {
                        if ($Parameter.Name -like "$ParameterName*") {
                            $Value = "-" + $Parameter.Name
                            New-TabItem -Value $Value -Text ("$Value [$($Parameter.ParameterType)]") -Type Parameter
                        }
                    }
                    $TabExpansionHasOutput = $true
                }
            }
        } elseif ($LastToken.Type -eq $_TokenTypes::GroupStart) {
            ## Tab complete method signatures
            $MethodTokens = @([System.Management.Automation.PSParser]::Tokenize($LastWord, [ref]$ParseErrors))
            if ($MethodTokens[-2].Type -eq $_TokenTypes::Member) {
                $MethodObject = $LastWord.SubString(0, $MethodTokens[-1].Start)
                $PossibleValues = foreach ($Overload in Invoke-Expression "$MethodObject.OverloadDefinitions") {
                    $Parameters = $Overload -replace '^\S+ .+\((.+)?\)','$1'
                    if ($Parameters) {
                        $Parameters = foreach ($Parameter in ($Parameters -split ", ")) {
                            $Type = ($Parameter -split " ")[0]
                            $Name = ($Parameter -split " ")[1]
                            '[{0}] ${1}' -f $Type,$Name
                        }
                        $Parameters = $Parameters -join ", "
                    }
                    $Value = "{0}{1})" -f $LastWord,$Parameters
                    New-TabItem -Value $Value -Text $Value -Type MethodSignature
                }
                $TabExpansionHasOutput = $true
            }
        }

        if ($PowerTabConfig -eq $null) {
            ## Something happened to the PowerTabConfig object, recreate it
            CreatePowerTabConfig
        }

        if ($TabExpansionHasOutput) {
            $PossibleValues | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler | ForEach-Object {
                if ($_ -is [String]) {
                    if ($QuoteSpaces -and ($_ -match " ") -and ($_ -notmatch "^[`"'].*[`"']`$") -and
                        (($LastToken.Type -eq $_TokenTypes::CommandArgument) -or ($LastWord -eq ""))) {
                        '"' + ($_ -replace '([\$"`])','`$1') + '"'
                    } else {
                        $_
                    }
                } else {
                    if ($QuoteSpaces -and ($_.Value -match " ") -and ($_.Value -notmatch "^[`"'].*[`"']`$") -and
                        (($LastToken.Type -eq $_TokenTypes::CommandArgument) -or ($LastWord -eq ""))) {
                        $_.Value = '"' + ($_.Value -replace '([\$"`])','`$1') + '"'
                    } else {
                        $_
                    }
                }
            }
            if ($PossibleValues.Count -lt 1) {
                ## No values
                if ($LastWord) {
                    ## Return what the user typed
                    $LastWord
                } else {
                    ## Send blank to prevent default PowerShell tab expansion
                    ""
                }
            }
        } else {
            Invoke-PowerTab -Line $Line -LastWord $LastWord -Context $CurrentContext -ForceList:$ForceList
        }
    } catch {
        Invoke-TabActivityIndicator -Error
        ""
    } finally {
        ## Remove busy indication on ready or error
        Remove-TabActivityIndicator
    }

    } finally {
        ## Save PowerTab errors and replace global errors
        $script:PowerTabError = $global:Error.Clone()
        $global:Error.Clear()
        $global:Error.AddRange($script:Error)
    }
}

Function New-TabContext {
    $Properties = @{
            "Command" = ""
            "CommandInfo" = $null
            "Parameter" = ""
            "Argument" = $null
            "ArgumentList" = @()
            "PositionalParameter" = -1
            "PositionalParameters" = @()
            "OtherParameters" = @{}
            "isCommandMode" = $false
            "isAssignment" = $false
            "isParameterValue" = $false
        }
    New-Object PSObject -Property $Properties
}


Function Invoke-PowerTab {
    param(
        $Line,
        $LastWord,
        $Context,
        [Switch]$ForceList
    ) 

    &{
    $TabExpansionHasOutput = $false

    if ($PowerTabConfig.IgnoreConfirmPreference) {
        $OriginalConfirmPreference = $ConfirmPreference
        $ConfirmPreference = 'High'
    }

    ## Helper variables 
    $_Method = [System.Management.Automation.PSMemberTypes]'Method,CodeMethod,ScriptMethod,ParameterizedProperty'
    $_ScopeNames = @("global", "local", "script", "private")

    ## Parse commandline
    $LineBlocks = [Regex]::Split($Line, '[|;]')
    $LastBlock = $LineBlocks[-1]

    ## Helper Functions
    Function Resolve-Member {
        param(
            $Object,
            $Pattern
        )

        ## Check for multilevel members
        $LevelCount = $Pattern.Split('.').Count
        if ($LevelCount -gt 1) { 
            $OFS = '.'; $Object += ".$($Pattern.Split('.')[0..($LevelCount -2)])"
        }

        ## Resolve Members
        $val = $Object
        $pat = $Pattern.Split('.')[($Level -1)] + '*'
        . {
            if ('PSBase' -like $pat) {$val + '.PSBase'}
            Invoke-Expression "Get-Member -InputObject ($val)" | Where-Object {
                $n = $_.Name
                if (-not $PowerTabConfig.ShowAccessorMethods) {
                    $n -like $pat -and $n -notmatch '^[gs]et_'
                } else {
                    $n -like $pat
                }
            } | ForEach-Object {
                if ($_.MemberType -band $_Method) {
                    ## Return a method...
                    $Value = $val + '.' + $_.Name + '('
                    New-TabItem -Value $Value -Text $Value -Type Method
                } else {
                    ## Return a property...
                    $Value = $val + '.' + $_.Name
                    New-TabItem -Value $Value -Text $Value -Type Property
                }
            }
        }
    }

    Function QuoteVariable {
        param(
            [String]$Name
        )

        ## Escape certain characters
        $Name = $Name -replace '([{}`])','`$1'

        ## If a variable name contains any of these characters it needs to be in braces
        $_varsRequiringQuotes = ('-`&@''#{}()$,;|<> .\/' + "`t").ToCharArray()
        if ($Name.IndexOfAny($_varsRequiringQuotes) -ge 0) {
            "{$Name}"
        } else {
            $Name
        }
    }

    &{
        ## Main tabcompletion , check line for patterns, select completion method and invoke handler

        ## Evaluate last block
        switch -regex ($LastBlock)  {
            ## Handle multilevel property and method expansion on simple () Blocks
            '(^| )\((.+)\)\.(.*)' {
                &{ trap {continue}
                    Resolve-Member -Object "($($Matches[2]))" -Pattern $Matches[3] |
                        Invoke-TabItemSelector $LastBlock -SelectionHandler $SelectionHandler | ForEach-Object {
                            $TabExpansionHasOutput = $true
                            if ($_.IndexOf(' ') -ge 0 ) {
                                ([Regex]::Split($_,' '))[-1].Trim()
                            } else {
                                $_
                            }
                        }
                }
                if ($TabExpansionHasOutput) {
                    ## Tab expansion handled, don't do anything more
                    return
                }
            }
        }

        ## Evaluate last word 
        switch -regex ($LastWord)  {
            ## Handle inline type search, e.g. new-object .identityreference<tab> or .identityre<tab> (oisin) 
            '^(\[*?)\.(\w+)$' {
                $TypeName = $Matches[2]
                Get-TabExpansion "*.${TypeName}*" Types | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Type |
                    Invoke-TabItemSelector $LastWord.Replace('[', '') -SelectionHandler $SelectionHandler |
                    ForEach-Object {if ($Matches[1] -eq '[') {"[$_]"} else {$_}}
                break
            }

            ## Members of script block
            '(.*\})\.([^\.]*)$' {
                $ScriptBlock = $Matches[1]
                $Member = $Matches[2]
                {} | Get-Member | Where-Object {
                    $n = $_.Name
                    if (-not $PowerTabConfig.ShowAccessorMethods) {
                        $n -like "$Member*" -and $n -notmatch '^[gs]et_'
                    } else {
                        $n -like "$Member*"
                    }
                } | ForEach-Object {
                    if ($_.MemberType -band $_Method) {
                        ## Return a method...
                        $Value = $ScriptBlock + '.' + $_.Name + '('
                        New-TabItem -Value $Value -Text $Value -Type Method
                    } else {
                        ## Return a property...
                        $Value = $ScriptBlock + '.' + $_.Name
                        New-TabItem -Value $Value -Text $Value -Type Property
                    }
                } | Invoke-TabItemSelector "$ScriptBlock.$Member" -SelectionHandler $SelectionHandler
                break
            }

            ## Completion on Shares (commented lines without DLL but need admin rights )
            '^(\\\\|//)(?<Computer>[^\\/]+)[\\/](?<Share>[^\\/]*)$' {
                #gwmi win32_share -computer $ComputerName -filter "name like '$($ShareName)%'" | Foreach-Object {"\\$($ComputerName)\$($_.name)"}
                #([adsi]"WinNT://$($ComputerName)/LanmanServer,FileService" ).psbase.children |? {$_.name -like "$($ShareName)*"}  |% {$_.name}
                $ComputerName = $Matches.Computer
                $ShareName = $Matches.Share
                [Trinet.Networking.ShareCollection]::GetShares($ComputerName) | Where-Object {$_.NetName -like "$ShareName*"} |
                    Sort-Object NetName | New-TabItem -Value {"\\$ComputerName\" + $_.NetName} -Text {"\\$ComputerName\" + $_.NetName} -Type Share |
                    Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## Completion on computers in database            
			'^(\\\\|//)(?<Computer>[^\\/]*)$' {
                $Computers = foreach ($Computer in Get-TabExpansion "$($Matches.Computer)*" Computer) {
                    $Value = "\\" + $Computer.Text
                    New-TabItem -Value $Value -Text $Value -Type Computer
                }
                $Computers | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## Variable "In-Place" expansion on \[tab]
            '.*(\$[^\\]+)\\$' {
                $val = $ExecutionContext.InvokeCommand.ExpandString($Matches[1])
                ($Matches[0] -replace ([Regex]::Escape($Matches[1])), $val).TrimEnd('\')
                break
            }

            ## Replace <command>-? with call to Get-Help
            '.*\-\?' {"Get-Help " + $LastWord.Replace('-?', ' -')}

            ## PSDrive expansion on :[tab]
            '^:$' {
                Get-PSDrive | New-TabItem -Value {$_.Name + ':'} -Text {$_.Name + ':'} -Type PSDrive |
                    Invoke-TabItemSelector '' -SelectionHandler $SelectionHandler
            }

            ## History completion against either #<pattern> or #<id>
            '^#(.*)' {
                $Pattern = $Matches[1]
                if ($Pattern -match '^[0-9]+$') {
                    @(Get-History -Id $Pattern -ErrorAction SilentlyContinue)[0].CommandLine
                } else {
					Get-History -Count 32767 | Where-Object {$_.CommandLine -like "$Pattern*"} | Sort Id -Descending |
                        Select-Object -ExpandProperty CommandLine -Unique | New-TabItem -Value {$_} -Text {$_} -Type History |
                        Invoke-TabItemSelector $Pattern -SelectionHandler $SelectionHandler
                }
                break
            }

            ## About Topics completion
            'about_(.*)' {
                Get-Help "about_$($Matches[1])*" | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Help |
                    Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## DataGrid GUI Shortcuts
            '^a_(.*)' {Get-Help "about_$($Matches[1])*" | Select-Object Name,Synopsis,Length | Out-DataGridView Name | Foreach-Object {Get-Help $_}}
            '^w_(.*)' {Get-TabExpansion "win32_$($Matches[1])*" WMI | Select-Object "Name" | Out-DataGridView Name}
            '^t_(.*)' {Get-TabExpansion "*$($Matches[1])*" Types | Select-Object "Name" | Out-DataGridView Name}
            '^c_(.*)' {Get-TabExpansion "$($Matches[1])*" | Select-Object "Text" | Out-DataGridView Text}
            '^f_' {Get-ChildItem function: | Select-Object Name | Out-DataGridView Name}
            '^d_' {Get-ChildItem | Select-Object Mode,LastWriteTime,Length,Name,FullName | Out-DataGridView FullName}
            '^h_' {Get-History -Count 100 | Out-DataGridView Commandline}

            ## WMI completion
            '(win32_.*|cim_.*|MSFT_.*)' {
                Get-TabExpansion "$($Matches[1])*" WMI | New-TabItem -Value {$_.Name} -Text {$_.Name} -Type WMIObject |
                    Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## Handle property and method expansion on variables
            '\$(.+)\.(.*)' {
                Resolve-Member -Object ('$' + $Matches[1]) -Pattern $Matches[2] |
                    Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## Translate typename to New-Object statement  [String][tab]
            '^(\[.*\])=(\w*)$' {
                "New-Object $($Matches[1].Replace('[','').Replace(']',''))"
                break
            }

            ## Handle Static methods of Types
            '(\[.*\])::(.*)$' { 
                $LastType = $Matches[1]
                $Level = $Matches[2].Split('.').Count

                if ($Level -gt 1) {
                    $LastType += ('::' + $Matches[2].Split('.')[0])
                    $Pattern = $Matches[2].SubString(($Matches[2].IndexOf('.') + 1))
                    Resolve-Member -Object $LastType -Pattern $Pattern | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                } else {
                    $Pattern = $Matches[2].Split('.')[($Level -1)] + '*'
                    Invoke-Expression "$($Matches[1]) | Get-Member -Static" | Where-Object {
                        $_.Name -like $Pattern -and $_.Name -notmatch '^[ge]et_'} | ForEach-Object {
                            if ($_.MemberType -band $_Method) {
                                $Value = "${LastType}::$($_.Name)" + '('
                                New-TabItem -Value $Value -Text $Value -Type StaticMethod
                            } else {
                                $Value = "${LastType}::$($_.Name)"
                                New-TabItem -Value $Value -Text $Value -Type StaticProperty
                            }
                        } | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                }
                break
            }

            ## Handle Static methods of types in Variables
            '(\$.*)::(.*)$' {
                $LastType = $Matches[1]
                $Level = $Matches[2].Split('.').Count
                $gt = ''

                if ((Invoke-Expression "$LastType.GetType().Name") -ne 'RuntimeType') {$gt = '.GetType'}
                if ($Level -gt 1) {
                    $LastType += ('::' + $Matches[2].Split('.')[0])
                    $Pattern = $Matches[2].SubString(($Matches[2].IndexOf('.') + 1))
                    Resolve-Member -Object $LastType -Pattern $Pattern | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                } else {
                    $Pattern = $Matches[2].Split('.')[($Level -1)] + '*'
                    Invoke-Expression "$($Matches[1]) | Get-Member -Static" | Where-Object {
                        $_.Name -like $Pattern -and $_.Name -notmatch '^[ge]et_'} | ForEach-Object {
                            if ($_.MemberType -band $_Method) {
                                $Value = "${LastType}${gt}::$($_.Name)" + '('
                                New-TabItem -Value $Value -Text $Value -Type StaticMethod
                            } else {
                                $Value = "${LastType}${gt}::$($_.Name)"
                                New-TabItem -Value $Value -Text $Value -Type StaticProperty
                            }
                        } | Invoke-TabItemSelector $LastWord.Replace('::',"${gt}::") -SelectionHandler $SelectionHandler
                }
                break
            }

            ## Handle enums and Constructors of Types
            '(\[.*\]):*$' {
                $LastType = $Matches[1].Split(',')[-1]
                $BaseType = (Invoke-Expression "$LastType.BaseType.FullName")
                . {
                    if ($BaseType -eq 'System.Enum') {
                        $Names = Invoke-Expression "[Enum]::GetNames($LastType)" |
                            New-TabItem -Value {"${LastType}::$_"} -Text {"${LastType}::$_"} -Type Enum |
                            Invoke-TabItemSelector ($LastType + '::') -SelectionHandler $SelectionHandler
                    } else {
                        $Constructors = Invoke-Expression "$LastType.GetConstructors()" | ForEach-Object {
                            $Regex = New-Object Regex('\((.*)\)')
                            $ParamTypes = foreach ($Type in $Regex.Match($_).Groups[1].Value.Split(',')) {
                                "[$($Type.Trim())]"
                            }
                            $Param = [String]::Join(' , ',$ParamTypes)
                            "New-Object $($LastType.Trim('[]'))($Param)".Replace('([])','()')
                        }
                        if ($Constructors) {
                            $Constructors | New-TabItem -Value {$_} -Text {$_} -Type Constructor |
                                Invoke-TabItemSelector $LastType -SelectionHandler $SelectionHandler
                        } else {
                            $LastWord
                        }
                    }
                }
                break
            }

            ## Handle members of Types (runtype)
            '^(\[.*\]).(\w*)$' {
                $LastType = $Matches[1].Split(',')[-1]
                Invoke-Expression "$LastType | Get-Member" | Where-Object {$_.Name -like "$($Matches[2])*"} | ForEach-Object {
                    if ($_.MemberType -band $_Method) {
                        $Value = "$LastType.$($_.Name)" + '('
                        New-TabItem -Value $Value -Text $Value -Type Method
                    } else {
                        $Value = "$LastType.$($_.Name)"
                        New-TabItem -Value $Value -Text $Value -Type Property
                    }
                } | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            } 

            ## Handle namespace and type names 
            '^\[(.*)$' {
                $Matched = $Matches[1]
                $Dots = $Matches[1].Split(".").Count - 1
                $res = @()
                $res += foreach ($Namespace in $dsTabExpansionDatabase.Tables['Types'].Select("NS like '$($Matched)%' and DC = $($Dots + 1)") |
                        Select-Object -Unique NS) {
                    $Value = "[$($Namespace.NS)"
                    New-TabItem -Value $Value -Text $Value -Type Namespace
                }
                $res += foreach ($Namespace in $dsTabExpansionDatabase.Tables['Types'].Select("NS like 'System.$($Matched)%' and DC = $($Dots + 2)") |
                        Select-Object -Unique NS) {
                    $Value = "[$($Namespace.NS)"
                    New-TabItem -Value $Value -Text $Value -Type Namespace
                }
                if ($Dots -gt 0) {
                    $res += foreach ($Type in $dsTabExpansionDatabase.Tables['Types'].Select("Name like '$($Matched)%' and DC = $Dots")) {
                        $Value = "[$($Type.Name)]"
                        New-TabItem -Value $Value -Text $Value -Type Type
                    }
                    $res += foreach ($Type in $dsTabExpansionDatabase.Tables['Types'].Select("Name like 'System.$($Matched)%' and DC = $($Dots + 1)")) {
                        $Value = "[$($Type.Name)]"
                        New-TabItem -Value $Value -Text $Value -Type Type
                    }
                }
                $res | Where-Object {$_} | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler -ForceList:$ForceList
                break
            }

            ## Handle expansions for both "Scope Variable Name" and "Type Variable Names"
            '^\$(\w+):(\w*)$' {
                $Type = $Matches[1]     # function, variable, global, etc.
                $TypeName = $Matches[2] # e.g. in '$function:C', value will be 'C'

                if ($_ScopeNames -contains $Type) {
                    # Scope variable name expansion ($global:, $script:, etc.)
                    $Variables = foreach ($ScopeVariable in (Get-Variable "$TypeName*" -Scope $Type)) {
                        '$' + (QuoteVariable ($Type + ":" + $ScopeVariable.Name))
                    }
                } else {
                    # Type variable name expansion ($function:, $variable:, $env:, etc.)
                    $Variables = foreach ($t in (Get-ChildItem ($Type + ":" + $TypeName + '*') | Sort-Object Name)) {
                        '$' + (QuoteVariable ($Type + ":" + $t.Name))
                    }
                }
                $Variables | New-TabItem -Value {$_} -Text {$_} -Type Variable | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## Handle variable name expansion
            '^([\$@])(\w*)$' {
                $VarName = $Matches[2]
                $Variables = foreach ($Variable in Get-Variable "$VarName*" -Scope Global) {
                    $Matches[1] + (QuoteVariable $Variable.Name)
                }
                $Variables | New-TabItem -Value {$_} -Text {$_} -Type Variable | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## Completion on cmdlets, function, aliases and native commands with defined shortcuts and custom additions from database

            ## Native commands / scripts in path
            "(.*)$([Regex]::Escape($PowerTabConfig.ShortcutChars.Native))`$" { 
                & {
                    Get-Command -CommandType ExternalScript -Name "$($Matches[1])*" |
                        New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Script
                    Get-Command -CommandType Application -Name "$($Matches[1])*" |
                        Where-Object {($env:PATHEXT).Split(";") -contains $_.Extension} |
                        New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Application
                } | Invoke-TabItemSelector $Matches[1] -SelectionHandler $SelectionHandler
                break
            }

            ## Aliases
            "(.+)$([Regex]::Escape($PowerTabConfig.ShortcutChars.Alias))`$" {
                & {
                    Get-Command -CommandType Alias -Name $Matches[1] | New-TabItem -Value {$_.Definition} -Text {$_.Definition} -Type Alias
                    Get-TabExpansion $Matches[1] Alias | New-TabItem -Value {$_.Text} -Text {$_.Text} -Type Alias
                } | Invoke-TabItemSelector $Matches[1] -SelectionHandler $SelectionHandler
                break
            }

            ## Custom
            "(.*)$([Regex]::Escape($PowerTabConfig.ShortcutChars.Custom))`$" {
                Get-TabExpansion "$($Matches[1])*" Custom | New-TabItem -Value {$_.Text} -Text {$_.Text} -Type Unknown |
                    Invoke-TabItemSelector $Matches[1] -SelectionHandler $SelectionHandler
                break
            }

            ## Invoke
            "(.+)$([Regex]::Escape($PowerTabConfig.ShortcutChars.Invoke))`$" {
                Get-TabExpansion "$($Matches[1])*" Invoke | ForEach-Object {
                    $ExecutionContext.InvokeCommand.InvokeScript($_.Text) | New-TabItem -Value {$_} -Text {$_} -Type Unknown
                } | Invoke-TabItemSelector $Matches[1] -SelectionHandler $SelectionHandler
                break
            }

            ## Call function
            "(.*)$([Regex]::Escape($PowerTabConfig.ShortcutChars.CustomFunction))`$" {
                if ($PowerTabConfig.CustomFunctionEnabled) {
                    $Matches[1] | ForEach-Object {
                        $ExecutionContext.InvokeCommand.InvokeScript("$($PowerTabConfig.CustomUserFunction) '$_'") | 
                            New-TabItem -Value {$_} -Text {$_} -Type Unknown
                    } | Invoke-TabItemSelector $Matches[1] -SelectionHandler $SelectionHandler
                }
                break
            }

            ## Partial functions or cmdlets
            "(.*)$([Regex]::Escape($PowerTabConfig.ShortcutChars.Partial))`$" {
                ## TODO: Give different types?
                Get-Command -CommandType Function,ExternalScript,Filter,Cmdlet -Name "$($Matches[1])*" |
                    New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Command |
                    Invoke-TabItemSelector $Matches[1] -SelectionHandler $SelectionHandler
                break
            }

            ## Functions or cmdlets on dash
            '(.+-.*)' {
                ## TODO: Give different types?
                Get-Command -CommandType Function,ExternalScript,Filter,Cmdlet -Name "$($Matches[1])*" |
                    New-TabItem -Value {$_.Name} -Text {$_.Name} -Type Command |
                    Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                break
            }

            ## Alternate alias
            '(.+)$' {
                if ($DoubleTab -or $PowerTabConfig.AliasQuickExpand) {
                    & {
                        Get-Command -CommandType Alias -Name $Matches[1] | New-TabItem -Value {$_.Definition} -Text {$_.Definition} -Type Alias
                        Get-TabExpansion $Matches[1] Alias | New-TabItem -Value {$_.Text} -Text {$_.Text} -Type Alias
                    } | Invoke-TabItemSelector $LastWord -SelectionHandler $SelectionHandler
                } else {
                    Get-TabExpansion $Matches[1] Alias | New-TabItem -Value {$_.Text} -Text {$_.Text} -Type Alias |
                        Invoke-TabItemSelector -SelectionHandler $SelectionHandler
                }
                break
            }
        } ## End of switch -regex $LastWord
    } | Where-Object {$_} | ForEach-Object {$TabExpansionHasOutput = $true; $_}

    ## Filesystem Completion
    if ((-not $TabExpansionHasOutput) -and $PowerTabConfig.FileSystemExpand) {
        $PowerTabFileSystemMode = $true
        $LastWord = $LastWord -replace '`'

        if (("Push-Location","Set-Location") -contains $Context.Command) {
            $ChildItems = @(Get-ChildItem "$LastWord*" | Where-Object {$_.PSIsContainer})
        } else {
            $ChildItems = @(Get-ChildItem "$LastWord*")
        }
        if (-not $ChildItems) {$LastWord; return}

        #if ((@($childitems).count -eq 1) -and ($lastword.endswith('\')) ) {$childitems = $childitems,@{name='..'}}
        $PathSlices = [Regex]::Split($LastWord, '\\|/')
        if ($PathSlices.Count -eq 1) {$PathSlices = ,"." + $PathSlices}
        $Container = [String]::Join('\', $PathSlices[0..($PathSlices.Count -2)])

        $LastPath = $Container + "\$([Regex]::Split($LastWord,'\\|/|:')[-1])"

        ## Fixes paths for registry keys, certificates and other unusual paths
        ## Improved fix for a problem identified by idvorkin (http://poshcode.org/1586)
        $ChildItems = foreach ($Item in $ChildItems) {
            $Child = switch ($Item.GetType().FullName) {
                "System.Security.Cryptography.X509Certificates.X509Certificate2" {$Item.Thumbprint;break}
                "Microsoft.Powershell.Commands.X509StoreLocation" {$Item.Location;break}
                "Microsoft.Win32.RegistryKey" {$Item.Name.Split("\")[-1];break}
                default {$Item.Name}
            }
            $Type = switch ($Item.GetType().FullName) {
                "System.IO.DirectoryInfo" {"Directory";break}
                "System.IO.FileInfo" {"File";break}
                "System.Management.Automation.AliasInfo" {"Alias";break}
                "System.Management.Automation.FilterInfo" {"Command";break}
                "System.Management.Automation.FunctionInfo" {"Command";break}
                "System.Security.Cryptography.X509Certificates.X509Certificate2" {"Certificate";break}
                "Microsoft.Powershell.Commands.X509StoreLocation" {"CertificateStore";break}
                "Microsoft.Win32.RegistryKey" {"RegistryKey";break}
                default {$_}
            }
            New-TabItem "$Container\$Child" "$Container\$Child" -Type $Type
        }
        $ChildItems | Invoke-TabItemSelector $LastPath -SelectionHandler $SelectionHandler -Return $LastWord -ForceList:$ForceList | ForEach-Object {
            ## If a path contains any of these characters it needs to be in quotes
            $_charsRequiringQuotes = ('`&@''#{}()$,; ' + "`t").ToCharArray()
        } {
            $Quote = ''
            $Invoke = ''

            if ($LastBlock -notmatch ".*['`"]`$") {  ## Don't quote if it looks like the path is already quoted
                if ($_ -is [String]) {
                    ## Remove quotes from beginning and end of string
                    $_ = $_ -replace '^"|"$'
                    ## Escape certain characters
                    $_ = $_ -replace '([\$"`])','`$1'

                    if ($_.IndexOfAny($_charsRequiringQuotes) -ge 0) {
                        ## Check for quotes in the last block of the input line,
                        ## if they exist, PowerShell will add them to this output
                        ## if not, then quotes can safely be added
                        if (-not (@([Char[]]$LastBlock | Where-Object {$_ -match '"|'''}).Count % 2)) {$Quote = '"'}
                        if (($LastBlock.Trim() -eq $LastWord)) {$Invoke = '& '}
                    }
                    "$Invoke$Quote$_$Quote"
                } else {
                    ## Remove quotes from beginning and end of string
                    $_.Value = $_.Value -replace '^"|"$'
                    ## Escape certain characters
                    $_.Value = $_.Value -replace '([\$"`])','`$1'

                    if ($_.Value.IndexOfAny($_charsRequiringQuotes) -ge 0) {
                        ## Check for quotes in the last block of the input line,
                        ## if they exist, PowerShell will add them to this output
                        ## if not, then quotes can safely be added
                        if (-not (@([Char[]]$LastBlock | Where-Object {$_ -match '"|'''}).Count % 2)) {$Quote = '"'}
                        if (($LastBlock.Trim() -eq $LastWord)) {$Invoke = '& '}
                    }
                    $_.Value = "$Invoke$Quote$($_.Value)$Quote"
                    $_
                }
            }
        }
    }

    } ## End of outer script block

    if ($PowerTabConfig.IgnoreConfirmPreference) {
        $ConfirmPreference = $OriginalConfirmPreference
    }

}  # end-function

# SIG # Begin signature block
# MIIY+QYJKoZIhvcNAQcCoIIY6jCCGOYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUasYzaiPl9mre+jJ/0KZ7Tbdn
# DIKgghSrMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
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
# FgQUEtc/p2Ib1I859SXt2ruIBLfhqrswDQYJKoZIhvcNAQEBBQAEggEASo64tT4x
# /rK/f8u9OKiTmNjRNandrE26Xbzuj8ajSUAE4JylkYbeCjTOZl9DiN33zcBBFjoe
# hj/M75/B8kOjOJzoppuz18yG1wYs+Is+TkQ9BnDJKJUwKtbwE79lBZA0+izpRQwt
# lf9Jpumg+gZ2AYlboPd5ryyTbKNPRHWR3i4lB898dGE9X3KmzKiGbPu2fBYEtUpr
# vsv076sxdtBJmbvMWHmnCbEqCwNR+hGJ4CG3bNJMFgNUkkM8GT6EnlSZQdfRB3k/
# wukMdgc6RWywqU4R/uzRav9+HgdieoTsTUofafm/AYZGE4mT+7Dwu9TyT5uTNyRQ
# m5cYawulbeH3paGCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBnMFMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UEAxMiVmVy
# aVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT2bg+9raN
# 7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTIwNjE1MTc1OTU5WjAjBgkqhkiG9w0BCQQxFgQUevRWO6MR6J+8
# idAIdGROiP5p+s4wDQYJKoZIhvcNAQEBBQAEgYCmnMjK1mt67EiCwiuX+y2uP2Fi
# 3bEIMgD4B07E8qIbhMPiQ2K35vMOOMB19rSUhew/JSFElTeg5lL2OK7pUtlSI+52
# J/aDZjUdP67Z5mL/0/G7taUnyKmeLxwC1DxzIN07np2mjuBZtGgArbdWKak1TERn
# jyYJ7cmhC6Y+2TpnwQ==
# SIG # End signature block
