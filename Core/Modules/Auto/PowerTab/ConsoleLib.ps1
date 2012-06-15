# ConsoleLib.ps1
#
# 


Function Out-ConsoleList {
    #[CmdletBinding()]
    param(
        [Parameter(Position = 1)]
        [ValidateNotNull()]
        [String]
        $LastWord = ''
        ,
        [Parameter(Position = 2)]
        [ValidateNotNull()]
        [String]
        $ReturnWord = ''  ## Text to return with filter if list closes without a selected item
        ,
        [Parameter(ValueFromPipeline = $true)]
        [ValidateNotNull()]
        [Object[]]
        $InputObject = @()
        ,
        [Switch]
        $ForceList
    )

    begin {
        [Object[]]$Content = @()
    }

    process {
        $Content += $InputObject
    }

    end {
        if (-not $ReturnWord) {$ReturnWord = $LastWord}

        ## If contents contains less than minimum options, then forward contents without displaying console list
        if (($Content.Length -lt $PowerTabConfig.MinimumListItems) -and (-not $ForceList)) {
            $Content | Select-Object -ExpandProperty Value
            return
        }

        ## Create console list
        $Filter = ''
        $ListHandle = New-ConsoleList $Content $PowerTabConfig.Colors.BorderColor $PowerTabConfig.Colors.BorderBackColor `
            $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor

        ## Preview of current filter, shows up where cursor is at
        $PreviewBuffer =  ConvertTo-BufferCellArray "$Filter " $PowerTabConfig.Colors.FilterColor $Host.UI.RawUI.BackgroundColor
        $Preview = New-Buffer $Host.UI.RawUI.CursorPosition $PreviewBuffer

        Function Add-Status {
            ## Title buffer, shows the last word in header of console list
            $TitleBuffer = ConvertTo-BufferCellArray " $LastWord" $PowerTabConfig.Colors.BorderTextColor $PowerTabConfig.Colors.BorderBackColor
            $TitlePosition = $ListHandle.Position
            $TitlePosition.X += 2
            $TitleHandle = New-Buffer $TitlePosition $TitleBuffer

            ## Filter buffer, shows the current filter after the last word in header of console list
            $FilterBuffer = ConvertTo-BufferCellArray "$Filter " $PowerTabConfig.Colors.FilterColor $PowerTabConfig.Colors.BorderBackColor
            $FilterPosition = $ListHandle.Position
            $FilterPosition.X += (3 + $LastWord.Length)
            $FilterHandle = New-Buffer $FilterPosition $FilterBuffer

            ## Status buffer, shows at footer of console list.  Displays selected item index, index range of currently visible items, and total item count.
            $StatusBuffer = ConvertTo-BufferCellArray "[$($ListHandle.SelectedItem + 1)] $($ListHandle.FirstItem + 1)-$($ListHandle.LastItem + 1) [$($Content.Length)]" $PowerTabConfig.Colors.BorderTextColor $PowerTabConfig.Colors.BorderBackColor
            $StatusPosition = $ListHandle.Position
            $StatusPosition.X += 2
            $StatusPosition.Y += ($listHandle.ListConfig.ListHeight - 1)
            $StatusHandle = New-Buffer $StatusPosition $StatusBuffer

        }
        . Add-Status

        ## Select the first item in the list
        $SelectedItem = 0
        Set-Selection 1 ($SelectedItem + 1) ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.SelectedTextColor $PowerTabConfig.Colors.SelectedBackColor

        ## Listen for first key press
        $Key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')

        ## Process key presses
        $Continue = $true
        while ($Key.VirtualKeyCode -ne 27 -and $Continue -eq $true) {
            if (-not $HasChild) {
                if ($OldFilter -ne $Filter) {
                  $Preview.Clear()
                  $PreviewBuffer = ConvertTo-BufferCellArray "$Filter " $PowerTabConfig.Colors.FilterColor $Host.UI.RawUI.BackgroundColor
                  $Preview = New-Buffer $Preview.Location $PreviewBuffer
                }
                $OldFilter = $Filter
            }
            $Shift = $Key.ControlKeyState.ToString()
            $HasChild = $false
            switch ($Key.VirtualKeyCode) {
                9 { ## Tab
                    ## In Visual Studio, Tab acts like Enter
                    if ($PowerTabConfig.VisualStudioTabBehavior) {
                        ## Expand with currently selected item
                        $ListHandle.Items[$ListHandle.SelectedItem].Value
                        $Continue = $false
                        break
                    } else {
                        if ($Shift -match 'ShiftPressed') {
                            Move-Selection -1  ## Up
                        } else {
                            Move-Selection 1  ## Down
                        }
                        break
                    }
                }
                38 { ## Up Arrow
                    if ($Shift -match 'ShiftPressed') {
                        ## Fast scroll selected
                        if ($PowerTabConfig.FastScrollItemCount -gt ($ListHandle.Items.Count - 1)) {
                            $Count = ($ListHandle.Items.Count - 1)
                        } else {
                            $Count = $PowerTabConfig.FastScrollItemCount
                        }
                        Move-Selection (- $Count)
                    } else {
                        Move-Selection -1
                    }
                    break
                }
                40 { ## Down Arrow
                    if ($Shift -match 'ShiftPressed') {
                        ## Fast scroll selected
                        if ($PowerTabConfig.FastScrollItemCount -gt ($ListHandle.Items.Count - 1)) {
                            $Count = ($ListHandle.Items.Count - 1)
                        } else {
                            $Count = $PowerTabConfig.FastScrollItemCount
                        }
                        Move-Selection $Count
                    } else {
                        Move-Selection 1
                    }
                    break
                }
                33 { ## Page Up
                    $Count = $ListHandle.Items.Count
                    if ($Count -gt $ListHandle.MaxItems) {
                        $Count = $ListHandle.MaxItems
                    }
                    Move-Selection (-($Count - 1))
                    break
                }
                34 { ## Page Down
                    $Count = $ListHandle.Items.Count
                    if ($Count -gt $ListHandle.MaxItems) {
                        $Count = $ListHandle.MaxItems
                    }
                    Move-Selection ($Count - 1)
                    break
                }
                39 { ## Right Arrow
                    ## Add a new character (the one right after the current filter string) from currently selected item
                    $Char = $ListHandle.Items[$ListHandle.SelectedItem].Text[($LastWord.Length + $Filter.Length + 1)]
                    $Filter += $Char
                    
                    $Old = $Items.Length
                    $Items = $Content -match ([Regex]::Escape("$LastWord$Filter") + '.*')
                    $New = $Items.Length
                    if ($New -lt 1) {
                        ## If new filter results in no items, sound error beep and remove character
                        Write-Host "`a" -NoNewline
                        $Filter = $Filter.SubString(0, $Filter.Length - 1)
                    } else {
                        if ($Old -ne $New) {
                            ## Update console list contents
                            $ListHandle.Clear()
                            $ListHandle = New-ConsoleList $Items $PowerTabConfig.Colors.BorderColor $PowerTabConfig.Colors.BorderBackColor `
                                $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor
                            ## Update status buffers
                            . Add-Status
                        }
                        ## Select first item of new list
                        $SelectedItem = 0
                        Set-Selection 1 ($SelectedItem + 1) ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.SelectedTextColor $PowerTabConfig.Colors.SelectedBackColor
                        $Host.UI.Write($PowerTabConfig.Colors.FilterColor, $Host.UI.RawUI.BackgroundColor, $Char)
                    }
                    break
                }
                {(8,37 -contains $_)} { # Backspace or Left Arrow
                    if ($Filter) {
                        ## Remove last character from filter
                        $Filter = $Filter.SubString(0, $Filter.Length - 1)
                        $Host.UI.Write([char]8)
                        Write-Line ($Host.UI.RawUI.CursorPosition.X) ($Host.UI.RawUI.CursorPosition.Y - $Host.UI.RawUI.WindowPosition.Y) " " $PowerTabConfig.Colors.FilterColor $Host.UI.RawUI.BackgroundColor

                        $Old = $Items.Length
                        $Items = @($Content | Where-Object {$_.Text -match ([Regex]::Escape("$LastWord$Filter") + '.*')})
                        $New = $Items.Length
                        if ($Old -ne $New) {
                            ## If the item list changed, update the contents of the console list
                            $ListHandle.Clear()
                            $ListHandle = New-ConsoleList $Items $PowerTabConfig.Colors.BorderColor $PowerTabConfig.Colors.BorderBackColor `
                                $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor
                            ## Update status buffers
                            . Add-Status
                        }
                        ## Select first item of new list
                        $SelectedItem = 0
                        Set-Selection 1 ($SelectedItem + 1) ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.SelectedTextColor $PowerTabConfig.Colors.SelectedBackColor
                    } else {
                        if ($PowerTabConfig.CloseListOnEmptyFilter) {
                            $Key.VirtualKeyCode = 27
                            $Continue = $false
                        } else {
                            Write-Host "`a" -NoNewline
                        }
                    }
                    break
                }
                190 { ## Period
                    if ($PowerTabConfig.DotComplete -and -not $PowerTabFileSystemMode) {
                        if ($PowerTabConfig.AutoExpandOnDot) {
                            ## Expand with currently selected item
                            $Host.UI.Write($Host.UI.RawUI.ForegroundColor, $Host.UI.RawUI.BackgroundColor, ($ListHandle.Items[$ListHandle.SelectedItem].Value.SubString($LastWord.Length + $Filter.Length) + '.'))
                            $ListHandle.Clear()
                            $LinePart = $Line.SubString(0, $Line.Length - $LastWord.Length)

                            ## Remove message handle ([Tab]) because we will be reinvoking tab expansion
                            Remove-TabActivityIndicator

                            ## Recursive tab expansion
                            . TabExpansion ($LinePart + $ListHandle.Items[$ListHandle.SelectedItem].Value + '.') ($ListHandle.Items[$ListHandle.SelectedItem].Value + '.') -ForceList
                            $HasChild = $true
                        } else {
                            $ListHandle.Items[$ListHandle.SelectedItem].Value
                        }
                        $Continue = $false
                        break
                    }
                }
                {'\','/' -contains $Key.Character} { ## Path Separators
                    if ($PowerTabConfig.BackSlashComplete) {
                        if ($PowerTabConfig.AutoExpandOnBackSlash) {
                            ## Expand with currently selected item
                            $Host.UI.Write($Host.UI.RawUI.ForegroundColor, $Host.UI.RawUI.BackgroundColor, ($ListHandle.Items[$ListHandle.SelectedItem].Value.SubString($LastWord.Length + $Filter.Length) + $Key.Character))
                            $ListHandle.Clear()
                            if ($Line.Length -ge $LastWord.Length) {
                                $LinePart = $Line.SubString(0, $Line.Length - $LastWord.Length)
                            }

                            ## Remove message handle ([Tab]) because we will be reinvoking tab expansion
                            Remove-TabActivityIndicator

                            ## Recursive tab expansion
                            . Invoke-TabExpansion ($LinePart + $ListHandle.Items[$ListHandle.SelectedItem].Value + $Key.Character) ($ListHandle.Items[$ListHandle.SelectedItem].Value + $Key.Character) -ForceList
                            $HasChild = $true
                        } else {
                            $ListHandle.Items[$ListHandle.SelectedItem].Value
                        }
                        $Continue = $false
                        break
                    }
                }
                32 { ## Space
                    ## True if "Space" and SpaceComplete is true, or "Ctrl+Space" and SpaceComplete is false
                    if (($PowerTabConfig.SpaceComplete -and -not ($Key.ControlKeyState -match 'CtrlPressed')) -or (-not $PowerTabConfig.SpaceComplete -and ($Key.ControlKeyState -match 'CtrlPressed'))) {
                        ## Expand with currently selected item
                        $Item = $ListHandle.Items[$ListHandle.SelectedItem].Value
                        if ((-not $Item.Contains(' ')) -and ($PowerTabFileSystemMode -ne $true)) {$Item += ' '}
                        $Item
                        $Continue = $false
                        break
                    }
                }
                {($PowerTabConfig.CustomCompletionChars.ToCharArray() -contains $Key.Character) -and $PowerTabConfig.CustomComplete} { ## Extra completions
                    $Item = $ListHandle.Items[$ListHandle.SelectedItem].Value
                    $Item = ($Item + $Key.Character) -replace "\$($Key.Character){2}$",$Key.Character
                    $Item
                    $Continue = $false
                    break
                }
                13 { ## Enter
                    ## Expand with currently selected item
                    $ListHandle.Items[$ListHandle.SelectedItem].Value
                    $Continue = $false
                    break
                }
                {$_ -ge 32 -and $_ -le 190}  { ## Letter or digit or symbol (ASCII)
                    ## Add character to filter
                    $Filter += $Key.Character

                    $Old = $Items.Length
                    $Items = @($Content | Where-Object {$_.Text -match ('^' + [Regex]::Escape("$LastWord$Filter") + '.*')})
                    $New = $Items.Length
                    if ($Items.Length -lt 1) {
                        ## New filter results in no items
                        if ($PowerTabConfig.CloseListOnEmptyFilter) {
                            ## Close console list and return the return word with current filter (includes new character)
                            $ListHandle.Clear()
                            return "$ReturnWord$Filter"
                        } else {
                            ## Sound error beep and remove character
                            Write-Host "`a" -NoNewline
                            $Filter = $Filter.SubString(0, $Filter.Length - 1)
                        }
                    } else {
                        if ($Old -ne $New) {
                            ## If the item list changed, update the contents of the console list
                            $ListHandle.Clear()
                            $ListHandle = New-ConsoleList $Items $PowerTabConfig.Colors.BorderColor $PowerTabConfig.Colors.BorderBackColor `
                                $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor
                            ## Update status buffer
                            . Add-Status
                            ## Select first item of new list
                            $SelectedItem = 0
                            Set-Selection 1 ($SelectedItem + 1) ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.SelectedTextColor $PowerTabConfig.Colors.SelectedBackColor
                        }

                        $Host.UI.Write($PowerTabConfig.Colors.FilterColor, $Host.UI.RawUI.BackgroundColor, $Key.Character)
                    }
                    break
                }
            }

            ## Listen for next key press
            if ($Continue) {$Key = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')}
        }

        $ListHandle.Clear()
        if (-not $HasChild) {
            if ($Key.VirtualKeyCode -eq 27) {
        		#Write-Line ($Host.UI.RawUI.CursorPosition.X - 1) ($Host.UI.RawUI.CursorPosition.Y - $Host.UI.RawUI.WindowPosition.Y) " " $PowerTabConfig.Colors.FilterColor $Host.UI.RawUI.BackgroundColor
                ## No items left and request that console list close, so return the return word with current filter
                return "$ReturnWord$Filter"
            }
        }
    }  ## end of "end" block
}


    Function New-Box {
        param(
            [System.Drawing.Size]
            $Size
            ,
            [System.ConsoleColor]
            $ForegroundColor = $Host.UI.RawUI.ForegroundColor
            ,
            [System.ConsoleColor]
            $BackgroundColor = $Host.UI.RawUI.BackgroundColor
        )

        $Box = New-Object System.Management.Automation.PSObject -Property @{
            'HorizontalDouble' = ([char]9552).ToString()
            'VerticalDouble' = ([char]9553).ToString()
            'TopLeftDouble' = ([char]9556).ToString()
            'TopRightDouble' = ([char]9559).ToString()
            'BottomLeftDouble' = ([char]9562).ToString()
            'BottomRightDouble' = ([char]9565).ToString()
            'Horizontal' = ([char]9472).ToString()
            'Vertical' = ([char]9474).ToString()
            'TopLeft' = ([char]9484).ToString()
            'TopRight' = ([char]9488).ToString()
            'BottomLeft' = ([char]9492).ToString()
            'BottomRight' = ([char]9496).ToString()
            'Cross' = ([char]9532).ToString()
            'HorizontalDoubleSingleUp' = ([char]9575).ToString()
            'HorizontalDoubleSingleDown' = ([char]9572).ToString()
            'VerticalDoubleLeftSingle' = ([char]9570).ToString()
            'VerticalDoubleRightSingle' = ([char]9567).ToString()
            'TopLeftDoubleSingle' = ([char]9554).ToString()
            'TopRightDoubleSingle' = ([char]9557).ToString()
            'BottomLeftDoubleSingle' = ([char]9560).ToString()
            'BottomRightDoubleSingle' = ([char]9563).ToString()
            'TopLeftSingleDouble' = ([char]9555).ToString()
            'TopRightSingleDouble' = ([char]9558).ToString()
            'BottomLeftSingleDouble' = ([char]9561).ToString()
            'BottomRightSingleDouble' = ([char]9564).ToString()
        }

        if ($PowerTabConfig.DoubleBorder) {
            ## Double line box
            $LineTop = $Box.TopLeftDouble `
                + $Box.HorizontalDouble * ($Size.width - 2) `
                + $Box.TopRightDouble
            $LineField = $Box.VerticalDouble `
                + ' ' * ($Size.width - 2) `
                + $Box.VerticalDouble
            $LineBottom = $Box.BottomLeftDouble `
                + $Box.HorizontalDouble * ($Size.width - 2) `
                + $Box.BottomRightDouble
        } elseif ($false) {
            ## Mixed line box, double horizontal, single vertical
            $LineTop = $Box.TopLeftDoubleSingle `
                + $Box.HorizontalDouble * ($Size.width - 2) `
                + $Box.TopRightDoubleSingle
            $LineField = $Box.Vertical `
                + ' ' * ($Size.width - 2) `
                + $Box.Vertical
            $LineBottom = $Box.BottomLeftDoubleSingle `
                + $Box.HorizontalDouble * ($Size.width - 2) `
                + $Box.BottomRightDoubleSingle
        } elseif ($false) {
            ## Mixed line box, single horizontal, double vertical
            $LineTop = $Box.TopLeftDoubleSingle `
                + $Box.HorizontalDouble * ($Size.width - 2) `
                + $Box.TopRightDoubleSingle
            $LineField = $Box.Vertical `
                + ' ' * ($Size.width - 2) `
                + $Box.Vertical
            $LineBottom = $Box.BottomLeftDoubleSingle `
                + $Box.HorizontalDouble * ($Size.width - 2) `
                + $Box.BottomRightDoubleSingle
        } else {  
            ## Single line box
            $LineTop = $Box.TopLeft `
                + $Box.Horizontal * ($Size.width - 2) `
                + $Box.TopRight
            $LineField = $Box.Vertical `
                + ' ' * ($Size.width - 2) `
                + $Box.Vertical
            $LineBottom = $Box.BottomLeft `
                + $Box.Horizontal * ($Size.width - 2) `
                + $Box.BottomRight
        }
        $Box = & {$LineTop; 1..($Size.Height - 2) | ForEach-Object {$LineField}; $LineBottom}
        $BoxBuffer = $Host.UI.RawUI.NewBufferCellArray($Box, $ForegroundColor, $BackgroundColor)
        ,$BoxBuffer
    }


    Function Get-ContentSize {
        param(
            [Object[]]$Content
        )

        $MaxWidth = @($Content | Select-Object -ExpandProperty Text | Sort-Object Length -Descending)[0].Length
        New-Object System.Drawing.Size $MaxWidth, $Content.Length
    }


    Function New-Position {
        param(
            [Int]$X
            ,
            [Int]$Y
        )

        $Position = $Host.UI.RawUI.WindowPosition
        $Position.X += $X
        $Position.Y += $Y
        $Position
    }


    Function New-Buffer {
        param(
            [System.Management.Automation.Host.Coordinates]
            $Position
            ,
            [System.Management.Automation.Host.BufferCell[,]]
            $Buffer
        )

        $BufferBottom = $BufferTop = $Position
        $BufferBottom.X += ($Buffer.GetUpperBound(1))
        $BufferBottom.Y += ($Buffer.GetUpperBound(0))
        $Rectangle = New-Object System.Management.Automation.Host.Rectangle $BufferTop, $BufferBottom
        $OldBuffer = $Host.UI.RawUI.GetBufferContents($Rectangle)
        $Host.UI.RawUI.SetBufferContents($BufferTop, $Buffer)
        $Handle = New-Object System.Management.Automation.PSObject -Property @{
            'Content' = $Buffer
            'OldContent' = $OldBuffer
            'Location' = $BufferTop
        }
        Add-Member -InputObject $Handle -MemberType 'ScriptMethod' -Name 'Clear' -Value {$Host.UI.RawUI.SetBufferContents($This.Location, $This.OldContent)}
        Add-Member -InputObject $Handle -MemberType 'ScriptMethod' -Name 'Show' -Value {$Host.UI.RawUI.SetBufferContents($This.Location, $This.Content)}
        $Handle
    }


    Function ConvertTo-BufferCellArray {
        param(
            [String[]]
            $Content
            ,
            [System.ConsoleColor]
            $ForegroundColor = $Host.UI.RawUI.ForegroundColor
            ,
            [System.ConsoleColor]
            $BackgroundColor = $Host.UI.RawUI.BackgroundColor
        )

        ,$Host.UI.RawUI.NewBufferCellArray($Content, $ForegroundColor, $BackgroundColor)
    }


    Function Parse-List {
        param(
            [System.Drawing.Size]$Size
        )

        $WindowPosition  = $Host.UI.RawUI.WindowPosition
        $WindowSize = $Host.UI.RawUI.WindowSize
        $Cursor = $Host.UI.RawUI.CursorPosition
        $Center = [Math]::Truncate([Float]$WindowSize.Height / 2)
        $CursorOffset = $Cursor.Y - $WindowPosition.Y
        $CursorOffsetBottom = $WindowSize.Height - $CursorOffset

        # Vertical Placement and size
        $ListHeight = $Size.Height + 2

        if (($CursorOffset -gt $Center) -and ($ListHeight -ge $CursorOffsetBottom)) {$Placement = 'Above'}
        else {$Placement =  'Below'}

        switch ($Placement) {
            'Above' {
                $MaxListHeight = $CursorOffset 
                if ($MaxListHeight -lt $ListHeight) {$ListHeight = $MaxListHeight}
                $Y = $CursorOffset - $ListHeight
            }
            'Below' {
                $MaxListHeight = ($CursorOffsetBottom - 1)  
                if ($MaxListHeight -lt $ListHeight) {$ListHeight = $MaxListHeight}
                $Y = $CursorOffSet + 1
            }
        }
        $MaxItems = $MaxListHeight - 2

        # Horizontal
        $ListWidth = $Size.Width + 4
        if ($ListWidth -gt $WindowSize.Width) {$ListWidth = $Windowsize.Width}
        $Max = $ListWidth 
        if (($Cursor.X + $Max) -lt ($WindowSize.Width - 2)) {
            $X = $Cursor.X
        } else {        
            if (($Cursor.X - $Max) -gt 0) {
                $X = $Cursor.X - $Max
            } else {
                $X = $windowSize.Width - $Max
            }
        }

        # Output
        $ListInfo = New-Object System.Management.Automation.PSObject -Property @{
            'Orientation' = $Placement
            'TopX' = $X
            'TopY' = $Y
            'ListHeight' = $ListHeight
            'ListWidth' = $ListWidth
            'MaxItems' = $MaxItems
        }
        $ListInfo
    }


    Function New-ConsoleList {
        param(
            [Object[]]
            $Content
            ,
            [System.ConsoleColor]
            $BorderForegroundColor
            ,
            [System.ConsoleColor]
            $BorderBackgroundColor
            ,
            [System.ConsoleColor]
            $ContentForegroundColor
            ,
            [System.ConsoleColor]
            $ContentBackgroundColor
        )

        $Size = Get-ContentSize $Content
        $MinWidth = ([String]$Content.Count).Length * 4 + 7
        if ($Size.Width -lt $MinWidth) {$Size.Width = $MinWidth}
        $Content = foreach ($Item in $Content) {
            $Item.DisplayText = " $($Item.Text) ".PadRight($Size.Width + 2)
            $Item
        }
        $ListConfig = Parse-List $Size
        $BoxSize = New-Object System.Drawing.Size $ListConfig.ListWidth, $ListConfig.ListHeight
        $Box = New-Box $BoxSize $BorderForegroundColor $BorderBackgroundColor

        $Position = New-Position $ListConfig.TopX $ListConfig.TopY
        $BoxHandle = New-Buffer $Position $Box

        # Place content 
        $Position.X += 1
        $Position.Y += 1
        $ContentBuffer = ConvertTo-BufferCellArray ($Content[0..($ListConfig.ListHeight - 3)] | Select-Object -ExpandProperty DisplayText) $ContentForegroundColor $ContentBackgroundColor
        $ContentHandle = New-Buffer $Position $ContentBuffer
        $Handle = New-Object System.Management.Automation.PSObject -Property @{
            'Position' = (New-Position $ListConfig.TopX $ListConfig.TopY)
            'ListConfig' = $ListConfig
            'ContentSize' = $Size
            'BoxSize' = $BoxSize
            'Box' = $BoxHandle
            'Content' = $ContentHandle
            'SelectedItem' = 0
            'SelectedLine' = 1
            'Items' = $Content
            'FirstItem' = 0
            'LastItem' = ($Listconfig.ListHeight - 3)
            'MaxItems' = $Listconfig.MaxItems
        }
        Add-Member -InputObject $Handle -MemberType 'ScriptMethod' -Name 'Clear' -Value {$This.Box.Clear()}
        Add-Member -InputObject $Handle -MemberType 'ScriptMethod' -Name 'Show' -Value {$This.Box.Show(); $This.Content.Show()}
        $Handle
    }


    Function Write-Line {
        param(
            [Int]$X
            ,
            [Int]$Y
            ,
            [String]$Text
            ,
            [System.ConsoleColor]
            $ForegroundColor
            ,
            [System.ConsoleColor]
            $BackgroundColor
        )

        $Position = $Host.UI.RawUI.WindowPosition
        $Position.X += $X
        $Position.Y += $Y
        if ($Text -eq '') {$Text = '-'}
        $Buffer = $Host.UI.RawUI.NewBufferCellArray([String[]]$Text, $ForegroundColor, $BackgroundColor)
        $Host.UI.RawUI.SetBufferContents($Position, $Buffer)
    }


    Function Move-List {
        param(
            [Int]$X
            ,
            [Int]$Y
            ,
            [Int]$Width
            ,
            [Int]$Height
            ,
            [Int]$Offset
        )

        $Position = $ListHandle.Position
        $Position.X += $X
        $Position.Y += $Y
        $Rectangle = New-Object System.Management.Automation.Host.Rectangle $Position.X, $Position.Y, ($Position.X + $Width), ($Position.Y + $Height - 1)
        $Position.Y += $OffSet
        $BufferCell = New-Object System.Management.Automation.Host.BufferCell
        $BufferCell.BackgroundColor = $PowerTabConfig.Colors.BackColor
        $Host.UI.RawUI.ScrollBufferContents($Rectangle, $Position, $Rectangle, $BufferCell)
    }


    Function Set-Selection {
        param(
            [Int]$X
            ,
            [Int]$Y
            ,
            [Int]$Width
            ,
            [System.ConsoleColor]
            $ForegroundColor
            ,
            [System.ConsoleColor]
            $BackgroundColor
        )

        $Position = $ListHandle.Position
        $Position.X += $X
        $Position.Y += $Y
        $Rectangle = New-Object System.Management.Automation.Host.Rectangle $Position.X, $Position.Y, ($Position.X + $Width), $Position.Y
        $LineBuffer = $Host.UI.RawUI.GetBufferContents($Rectangle)
        $LineBuffer = $Host.UI.RawUI.NewBufferCellArray(@([String]::Join("", ($LineBuffer | ForEach-Object {$_.Character}))),
            $ForegroundColor, $BackgroundColor)
        $Host.UI.RawUI.SetBufferContents($Position, $LineBuffer)
    }


    Function Move-Selection {
        param(
            [Int]$Count
        )

        $SelectedItem = $ListHandle.SelectedItem
        $Line = $ListHandle.SelectedLine
        if ($Count -eq ([Math]::Abs([Int]$Count))) { ## Down in list
            if ($SelectedItem -eq ($ListHandle.Items.Count - 1)) {return}
            $One = 1
            if ($SelectedItem -eq $ListHandle.LastItem) {
                $Move = $true
                if (($ListHandle.Items.Count - $SelectedItem - 1) -lt $Count) {$Count = $ListHandle.Items.Count - $SelectedItem - 1}
            } else {
                $Move = $false
                if (($ListHandle.MaxItems - $Line) -lt $Count) {$Count = $ListHandle.MaxItems - $Line}       
            }
        } else {
            if ($SelectedItem -eq 0) {return}
            $One = -1
            if ($SelectedItem -eq $ListHandle.FirstItem) {
                $Move = $true
                if ($SelectedItem -lt ([Math]::Abs([Int]$Count))) {$Count = (-($SelectedItem))}
            } else {
                $Move = $false
                if ($Line -lt ([Math]::Abs([Int]$Count))) {$Count = (-$Line) + 1}
            }
        }

        if ($Move) {
            Set-Selection 1 $Line ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor
            Move-List 1 1 ($ListHandle.ListConfig.ListWidth - 3) ($ListHandle.ListConfig.ListHeight - 2) (-$Count)
            $SelectedItem += $Count
            $ListHandle.FirstItem += $Count
            $ListHandle.LastItem += $Count

            $LinePosition = $ListHandle.Position
            $LinePosition.X += 1
            if ($One -eq 1) {
                $LinePosition.Y += $Line - ($Count - $One)
                $LineBuffer = ConvertTo-BufferCellArray ($ListHandle.Items[($SelectedItem - ($Count - $One)) .. $SelectedItem] | Select-Object -ExpandProperty Text) $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor
            } else {
                $LinePosition.Y += 1
                $LineBuffer = ConvertTo-BufferCellArray ($ListHandle.Items[($SelectedItem..($SelectedItem - ($Count - $One)))] | Select-Object -ExpandProperty Text) $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor
            }
            $LineHandle = New-Buffer $LinePosition $LineBuffer
            Set-Selection 1 $Line ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.SelectedTextColor $PowerTabConfig.Colors.SelectedBackColor
        } else {
            Set-Selection 1 $Line ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.TextColor $PowerTabConfig.Colors.BackColor
            $SelectedItem += $Count
            $Line += $Count
            Set-Selection 1 $Line ($ListHandle.ListConfig.ListWidth - 3) $PowerTabConfig.Colors.SelectedTextColor $PowerTabConfig.Colors.SelectedBackColor
        }
        $ListHandle.SelectedItem = $SelectedItem
        $ListHandle.SelectedLine = $Line

        ## New status buffer
        $StatusHandle.Clear()
        $StatusBuffer = ConvertTo-BufferCellArray "[$($ListHandle.SelectedItem + 1)] $($ListHandle.FirstItem + 1)-$($ListHandle.LastItem + 1) [$($Content.Length)]" `
            $PowerTabConfig.Colors.BorderTextColor $PowerTabConfig.Colors.BorderBackColor
        $StatusHandle = New-Buffer $StatusHandle.Location $StatusBuffer
    }


# SIG # Begin signature block
# MIIY+QYJKoZIhvcNAQcCoIIY6jCCGOYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUCN/Z4BDIlhFvj1R35kJ2pEPf
# NZagghSrMIIDnzCCAoegAwIBAgIQeaKlhfnRFUIT2bg+9raN7TANBgkqhkiG9w0B
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
# FgQU5bBh8B4CxKrqlMFMgKMFJqKTg3owDQYJKoZIhvcNAQEBBQAEggEAJ2mGulgN
# vs+k/dyNhRdPl+JRxFFIM6FAZGI2ymJWZRVbqliE/hm7n4TofYpu2td7kRnJ9Y8A
# 59quqc3ekwht2fEY+5f8IwrT6TZPdlLV5xIoKv5fa4bLjmoKDPgGgCjQTbc5EQtu
# uYJsLIO8HAm2IfFEZEg7fI7wOg+5ozl4AIi1DVON4LTLxpQmNAogbmKKEYZxZvMQ
# O5VhhQ6zeJ4EK0t8oE+NOIWYXnkl5YWqVXLfxxONaHtbX0z3Emx0bRzk+/o70IJH
# UEewThlDiDSpgSsQ2ZZm+hKfhL6XRVv+GarfgPe0+hz6hEE/HoRpzT226BTAgPfq
# 2BTGC8xUXSmwxaGCAX8wggF7BgkqhkiG9w0BCQYxggFsMIIBaAIBATBnMFMxCzAJ
# BgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjErMCkGA1UEAxMiVmVy
# aVNpZ24gVGltZSBTdGFtcGluZyBTZXJ2aWNlcyBDQQIQeaKlhfnRFUIT2bg+9raN
# 7TAJBgUrDgMCGgUAoF0wGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG
# 9w0BCQUxDxcNMTIwNjE1MTgwNzIyWjAjBgkqhkiG9w0BCQQxFgQUPwI99F/3EqGA
# tayfCCcgFg1Xa7QwDQYJKoZIhvcNAQEBBQAEgYAm1f/zD9fDs+7bad72RZ7lqPY0
# dHzdpm3Krm3d4jbEEpRj0f5MChHcvr9EfN3lyzs2tcMLdHo8oiWxSm64K4XYb79D
# c/CrRcsoRnMA5rxq1gvnLCzQQKGopOwX5rckzdqvjEU/R89Hf1SE40BsMLlouuR/
# Mc/uRHZmX/vU9Hz9ng==
# SIG # End signature block
