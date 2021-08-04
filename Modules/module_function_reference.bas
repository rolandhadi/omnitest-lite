Attribute VB_Name = "module_function_reference"
Public function_reference_search_results
Public function_reference_option_search_results

Sub function_reference_new()
On Error GoTo err_handler
    new_fr_name = Trim(InputBox("Enter Function Reference name"))
    If new_fr_name <> "" Then
        Set new_fr = New FunctionReference
        new_fr.name = new_fr_name
        new_fr.save
        function_reference_initial_state
        function_reference_unlock
        Application.Range("A2").value = new_fr.id
        Application.Range("B2").value = new_fr.name
        Application.Range("B4").value = 1
        Application.Range("A5").value = 1
        function_reference_lock
        border_and_color
        function_reference_search_results = Null
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
    function_reference_initial_state
End Sub

Sub function_reference_clear()
    function_reference_clear_sheet
    function_reference_search_results = Null
End Sub

Sub function_reference_clear_sheet()
    function_reference_initial_state
End Sub

Sub function_reference_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init FUNCTION_REFERENCE_TAB, False
    new_search.Show 1
    If Not IsNull(function_reference_search_results) Then function_reference_show
    
Exit Sub
err_handler:
    MsgBox Err.description

End Sub

Sub function_reference_update()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        Set new_fr = FunctionReferenceFactory(CDbl(Application.Range("A2").value))
        If new_fr.count > 0 Then
            Set update_fr = new_fr.first
            update_fr.name = Trim(Application.Range("B2").value)
            update_fr.save
            MsgBox update_fr.name & " updated successfully!"
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub function_reference_delete()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set new_fr = FunctionReferenceFactory(CDbl(Application.Range("A2").value))
        If new_fr.count > 0 Then
            Set update_fr = new_fr.first
            update_fr.delete
            MsgBox update_fr.name & " deleted successfully!"
            function_reference_initial_state
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub function_reference_initial_state()
    function_reference_unlock
    SheetClear
    Application.Range("A1").value = "ID"
    Application.Range("B1").value = "Function Reference Name"
    
    Application.Range("A4").value = "Row"
    
    function_reference_unlock
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyBorder "A1:B2"
    ApplyColor "A1:" & "B" & 1, colorBlack
    ApplyFontColor "A1:" & "B" & 1, colorWhite
    
    ApplyBorder "A4:B5"
    ApplyColor "A4:" & "B" & 4, colorBlack
    ApplyFontColor "A4:" & "B" & 4, colorWhite
    FreezePane 4
    function_reference_lock
    Application.Range("A2").Select
End Sub
    
Public Sub function_reference_show()
    On Error GoTo err_handler
    Application.ScreenUpdating = False
    function_reference_clear_sheet
    function_reference_unlock
    disable_keyboard_check = True
    For Each fr In function_reference_search_results
        Application.Range("A" & 2).value = fr.id
        Application.Range("B" & 2).value = fr.name
        col_adder = 2
        For Each v In Split(fr.struct, "|")
            Application.Range(columnLetter(col_adder) & 4).value = v
            col_adder = col_adder + 1
        Next
        If fr.values.count > 0 Then
            frvs = fr.values.fetch
            col_adder = 2
            For Each frv In frvs
                For Each v In Split(frv.item, "|")
                    Application.Range(columnLetter(col_adder) & 5).value = v
                    col_adder = col_adder + 1
                Next
            Next
        End If
    Next
    border_and_color_to columnLetter(col_adder)
    border_and_color
    function_reference_lock
    Application.ScreenUpdating = True
    disable_keyboard_check = False
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Sub function_reference_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="FunctionReferencesLock", Range:=Range( _
        "B2,B4:Z4,B5:Z50")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Application.Columns("A:C").EntireColumn.AutoFit
End Sub

Public Sub function_reference_unlock()
    Application.ActiveSheet.Unprotect
    Application.Columns("A:G").EntireColumn.AutoFit
End Sub

Public Function function_reference_step_new()
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set new_fr = FunctionReferenceFactory(CDbl(Application.Range("A2").value))
        If new_fr.count > 0 Then
            Set new_frv = New FunctionReferenceValue
            new_frv.fr_id = new_fr.first.id
            new_frv.order_no = new_fr.first.values.count + 1
            new_frv.tc_id = test_case_search_results(0).id
            new_frv.save
            function_reference_unlock
            Application.ActiveCell.value = new_frv.id
            Application.Range("B" & Application.ActiveCell.row).value = new_frv.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_case_search_results(0).name
            border_and_color
            function_reference_lock
        End If
    End If
End Function

Public Function function_reference_step_insert()
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set new_fr = FunctionReferenceFactory(CDbl(Application.Range("A2").value))
        If new_fr.count > 0 Then
            Set new_frv = New FunctionReferenceValue
            new_frv.fr_id = new_fr.first.id
            new_frv.order_no = new_fr.first.values.count + 1
            new_frv.tc_id = test_case_search_results(0).id
            new_frv.save
            function_reference_unlock
            Application.rows(Application.ActiveCell.row + 1).Insert Shift:=xlUp
            cur_selected_row = Application.ActiveCell.row + 1
            Application.Range("A" & Application.ActiveCell.row + 1).Select
            Application.ActiveCell.value = new_frv.id
            Application.Range("B" & Application.ActiveCell.row).value = new_frv.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_case_search_results(0).name
            Application.ScreenUpdating = False
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_frv = get_frv
                If cur_frv.count > 0 Then
                    Set update_frv = cur_frv.first
                    update_frv.order_no = i - 4
                    update_frv.save
                    Application.Range("B" & i).value = update_frv.order_no
                End If
            Next
            border_and_color
            Application.Range("A" & cur_selected_row).Select
            function_reference_lock
        End If
    End If
    
End Function

Public Function function_reference_step_delete()
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_frv = FunctionReferenceValueFactory(CDbl(Application.ActiveCell.value))
        If cur_frv.count > 0 Then
            function_reference_unlock
            cur_frv.first.delete
            Application.rows(Application.ActiveCell.row).delete Shift:=xlUp
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_frv = get_frv
                If cur_frv.count > 0 Then
                    Set update_frv = cur_frv.first
                    update_frv.order_no = i - 4
                    update_frv.save
                    Application.Range("B" & i).value = update_frv.order_no
                End If
            Next
            border_and_color
            function_reference_lock
        End If
    End If
End Function

Public Function function_reference_step_order()
    cur_selected_row = Application.ActiveCell.row
    new_order_no = Val(InputBox("Enter new step order number", "Step Order", cur_selected_row - 4))
    If cur_selected_row = new_order_no Then Exit Function
    If new_order_no <= 0 Then Exit Function
    sheet_size = get_last_row_column
    function_reference_unlock
    Application.ScreenUpdating = False
    If new_order_no >= sheet_size(0) - 4 Then
        new_order_no = sheet_size(0) - 4
        Application.rows(cur_selected_row).Select
        Selection.Cut
        Application.rows(sheet_size(0) + 1).Select
        Selection.Insert Shift:=xlDown
    Else
        Application.rows(cur_selected_row).Select
        Selection.Cut
        Application.rows(new_order_no + 4).Select
        Selection.Insert Shift:=xlDown
    End If
    For i = 5 To get_last_row_column(0)
        Application.Range("A" & i).Select
        Set cur_frv = get_frv
        If cur_frv.count > 0 Then
            Set update_frv = cur_frv.first
            update_frv.order_no = i - 4
            update_frv.save
            Application.Range("B" & i).value = update_frv.order_no
        End If
    Next
    Application.ScreenUpdating = True
    border_and_color
    function_reference_lock
End Function

Public Function function_reference_step_test_case()
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set cur_frv = FunctionReferenceValueFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
        If cur_frv.count > 0 Then
            Set update_frv = cur_frv.first
            update_frv.tc_id = test_case_search_results(0).id
            update_frv.save
            function_reference_unlock
            Application.ActiveCell.value = update_frv.id
            Application.Range("C" & Application.ActiveCell.row).value = test_case_search_results(0).name
            border_and_color
            function_reference_lock
        End If
    End If
End Function

Public Function function_reference_update_column(column_no)
    
    function_reference_unlock
    Application.Columns("A:" & columnLetter(column_no)).EntireColumn.AutoFit
    ApplyBorder "A4:" & columnLetter(column_no) & "5"
    ApplyColor "A4:" & columnLetter(column_no) & 4, colorBlack
    ApplyFontColor "A4:" & columnLetter(column_no) & 4, colorWhite
    border_and_color_to columnLetter(column_no)
    function_reference_lock
    
    
End Function

Private Function get_frv()
    Set get_frv = FunctionReferenceValueFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
End Function

Private Function border_and_color()
    
    function_reference_unlock
    sheet_size = get_last_row_column
    ApplyBorder ("A" & 4 & ":B" & sheet_size(0))
    Application.Columns("A:B").EntireColumn.AutoFit
    ApplyAltColor 5, sheet_size(0), "A", "B"
    function_reference_lock
    
End Function

Private Function border_and_color_to(to_letter)
    
    function_reference_unlock
    sheet_size = get_last_row_column
    ApplyColor "A4:" & to_letter & 4, colorBlack
    ApplyFontColor "A4:" & to_letter & 4, colorWhite
    ApplyBorder ("A" & 4 & ":" & to_letter & sheet_size(0))
    Application.Columns("A:" & to_letter).EntireColumn.AutoFit
    ApplyAltColor 5, sheet_size(0), "A", to_letter
    function_reference_lock
    
End Function






