Attribute VB_Name = "module_common_function"
    Public Const colorBlack = 0
    Public Const colorGray = 8355711
    Public Const colorWhite = 16777215
    
    Public Function get_last_row_column()
        get_last_row_column = Array(Application.ActiveSheet.UsedRange.rows.count, Application.ActiveSheet.UsedRange.Columns.count)
    End Function
    
    Public Function id_found(id)
        sheet_size = get_last_row_column
        found = False
        For i = 2 To sheet_size(0) + row_adder
            If Application.Range("A" & i).value = id Then
                found = True
                Exit For
            End If
        Next
        id_found = found
    End Function
    
    Public Sub ApplyBorder(myRange)
        Application.ScreenUpdating = False: disableScreenUpdate = True
        Application.Range(myRange).Select
        Application.Selection.Borders(5).LineStyle = -4142
        Application.Selection.Borders(6).LineStyle = -4142
        With Application.Selection.Borders(7)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(8)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(9)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(10)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(11)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(12)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
    End Sub

    Public Sub ApplyOutsideBorder(myRange)
        Application.ScreenUpdating = False: disableScreenUpdate = True
        Application.Range(myRange).Select
        Application.Selection.Borders(5).LineStyle = -4142
        Application.Selection.Borders(6).LineStyle = -4142
        With Application.Selection.Borders(7)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(8)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(9)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(10)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
    End Sub

    Public Sub ApplyThickBorder(myRange)
        Application.ScreenUpdating = False: disableScreenUpdate = True
        Application.Range(myRange).Select
        Application.Selection.Borders(5).LineStyle = -4142
        Application.Selection.Borders(6).LineStyle = -4142
        With Application.Selection.Borders(7)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = -4138
        End With
        With Application.Selection.Borders(8)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = -4138
        End With
        With Application.Selection.Borders(9)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = -4138
        End With
        With Application.Selection.Borders(10)
            .LineStyle = 1
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = -4138
        End With
    End Sub

    Public Sub ApplyDottedBorder(myRange)
        Application.ScreenUpdating = False: disableScreenUpdate = True
        Application.Range(myRange).Select
        Application.Selection.Borders(5).LineStyle = -4142
        Application.Selection.Borders(6).LineStyle = -4142
        With Application.Selection.Borders(7)
            .LineStyle = -4118
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(8)
            .LineStyle = -4118
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(9)
            .LineStyle = -4118
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(10)
            .LineStyle = -4118
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(11)
            .LineStyle = -4118
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
        With Application.Selection.Borders(12)
            .LineStyle = -4118
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = 2
        End With
    End Sub

    Public Sub ApplyColor(myRange, myColor)
        Application.ScreenUpdating = False: disableScreenUpdate = True
        Application.Range(myRange).Select
        With Application.Selection.Interior
            .Pattern = 1
            .PatternColorIndex = -4105
            .Color = myColor
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End Sub

    Public Sub ApplyAltColor(startRow, endRow, startCol, endCol)
        Dim i, myColor
        Application.ScreenUpdating = False: disableScreenUpdate = True
        For i = startRow To endRow
            Application.Range(startCol & i & ":" & endCol & i).Select
            If i Mod 2 = 0 Then
                myColor = 16777215
            Else
                myColor = 15921906
            End If
            With Application.Selection.Interior
                .Pattern = 1
                .PatternColorIndex = -4105
                .Color = myColor
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next
    End Sub

    Public Sub ApplyFontColor(myRange, myColor)
        Application.Range(myRange).Font.Color = myColor
    End Sub

    Public Sub FreezePane(rows)
        Application.rows(1).Select
        With Application.ActiveWindow
            .SplitColumn = 0
            .SplitRow = rows
        End With
        Application.ActiveWindow.FreezePanes = True
    End Sub

    Public Sub ApplyFilter()
        Dim colLetter
        colLetter = Application.Cells(1, Application.ActiveSheet.UsedRange.Columns.count).Address
        colLetter = Replace(colLetter, "$", "")
        colLetter = Left(colLetter, Len(colLetter) - 1)
        Application.Range("A1:" & colLetter & "1").AutoFilter
    End Sub
    
    Public Sub SheetClear()
        Application.Cells.Select
        Application.Selection.delete Shift:=-4162
        Application.Selection.NumberFormat = "@"
    End Sub
    
    Function columnLetter(ByVal columnNumber)
        Dim n
        Dim C
        Dim s
        s = ""
        n = columnNumber
        Do
            C = ((n - 1) Mod 26)
            s = Chr(C + 65) & s
            n = (n - C) \ 26
        Loop While n > 0
        columnLetter = s
    End Function
    
    Function RandomString(ByVal strLen)
        Dim str, min, max
    
        Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789"
        min = 1
        max = Len(LETTERS)
    
        Randomize
        For i = 1 To strLen
            str = str & Mid(LETTERS, Int((max - min + 1) * Rnd + min), 1)
        Next
        RandomString = str
    End Function
    
    Public Function IsNullOrEmpty(value)
        out = True
        If IsObject(value) Then
            out = False
        Else
            If IsEmpty(value) Then
                out = True
            ElseIf value = "" Then
                out = True
            Else
                out = False
            End If
        End If
        IsNullOrEmpty = out
    End Function

