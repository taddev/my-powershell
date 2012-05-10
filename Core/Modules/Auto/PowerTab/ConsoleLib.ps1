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
# MIIY1AYJKoZIhvcNAQcCoIIYxTCCGMECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUCN/Z4BDIlhFvj1R35kJ2pEPf
# NZagghSGMIIDejCCAmKgAwIBAgIQOCXX+vhhr570kOcmtdZa1TANBgkqhkiG9w0B
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
# CSqGSIb3DQEJBDEWBBTlsGHwHgLEquqUwUyAowUmopODejANBgkqhkiG9w0BAQEF
# AASCAQAnaYa6WA2+z6T93I2FF0+X4lHEUUgzoUBkYjbKYlZlFVuqWIT+GbufhOh9
# im7a13uRGcn1jwDn2q6pzd6TCG3Z8Rj7l/wjCtPpNk92UtXnEigq/l9rhsuOagoM
# +AaAKNBNtzkRC265gmwsg7wcCbYh8URkSDt8jvA6D7mjOXgAiLUNU43gtMvGlCY0
# CiBuYooRhnFm8xA7lWGFDrN4ngQrS3ygT404hZheeSXlhapVct/HE41oe1tfTPcS
# bHRtHOT7+jvQgkdQR7BOGUOINKmBKxDZlmb6Ep+EvpdFW/4Zqt+A97T6HPqEQT8e
# hGnNPbboFMCA9+rYFMYLzFRdKbDFoYIBfzCCAXsGCSqGSIb3DQEJBjGCAWwwggFo
# AgEBMGcwUzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMSsw
# KQYDVQQDEyJWZXJpU2lnbiBUaW1lIFN0YW1waW5nIFNlcnZpY2VzIENBAhA4Jdf6
# +GGvnvSQ5ya11lrVMAkGBSsOAwIaBQCgXTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcN
# AQcBMBwGCSqGSIb3DQEJBTEPFw0xMTEwMjgxOTUzMzVaMCMGCSqGSIb3DQEJBDEW
# BBQ/Aj30X/cSoYC1rJ8IJyAWDVdrtDANBgkqhkiG9w0BAQEFAASBgExhsB8mcAci
# 8lM+z4A6JFLV3AMYic/dKbryclCZH+sZ5kjAJa3iPP43vE0Ev65DlVWINVSS7ucs
# AhjTvGS9LYhkqWxRL3QrIWfmWDF6O9D+ub0SA/VQdpCmrvJceMgtIyjY2AccmbUu
# q4Wtvqjjd4uMBBrHbzpAkPVqyekaewq7
# SIG # End signature block
