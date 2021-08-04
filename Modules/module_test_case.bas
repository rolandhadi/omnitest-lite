Attribute VB_Name = "module_test_case"
Public test_case_search_results
Public test_case_option_search_results

Sub test_case_new()
On Error GoTo err_handler
    new_tc_name = Trim(InputBox("Enter Test Case name"))
    If new_tc_name <> "" Then
        Set new_tc = New TestCase
        new_tc.name = new_tc_name
        new_tc.save
        test_case_initial_state
        test_case_unlock
        Application.Range("A2").value = new_tc.id
        Application.Range("B2").value = new_tc.name
        test_case_lock
        border_and_color
        test_case_search_results = Null
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
    test_case_initial_state
End Sub

Sub test_case_clear()
    test_case_clear_sheet
    test_case_search_results = Null
End Sub

Sub test_case_clear_sheet()
    test_case_initial_state
End Sub

Sub test_case_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then test_case_show
    
Exit Sub
err_handler:
    MsgBox Err.description

End Sub

Sub test_case_update()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        Set cur_tc = TestCaseFactory(CDbl(Application.Range("A2").value))
        If cur_tc.count > 0 Then
            Set update_tc = cur_tc.first
            update_tc.name = Trim(Application.Range("B2").value)
            update_tc.save
            MsgBox update_tc.name & " updated successfully!"
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_case_delete()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_tc = TestCaseFactory(CDbl(Application.Range("A2").value))
        If cur_tc.count > 0 Then
            Set update_tc = cur_tc.first
            update_tc.delete
            MsgBox update_tc.name & " deleted successfully!"
            test_case_initial_state
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_case_initial_state()
    test_case_unlock
    SheetClear
    Application.Range("A1").value = "ID"
    Application.Range("B1").value = "Test Case Name"
    
    Application.Range("A4").value = "ID"
    Application.Range("B4").value = "Order"
    Application.Range("C4").value = "Test Procedure"
    
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyBorder "A1:B2"
    ApplyColor "A1:" & "B" & 1, colorBlack
    ApplyFontColor "A1:" & "B" & 1, colorWhite
    
    ApplyBorder "A4:C5"
    ApplyColor "A4:" & "C" & 4, colorBlack
    ApplyFontColor "A4:" & "C" & 4, colorWhite
    FreezePane 4
    test_case_lock
    Application.Range("A2").Select
End Sub
    
Public Sub test_case_show()
    On Error GoTo err_handler
    Application.ScreenUpdating = False
    test_case_clear_sheet
    test_case_unlock
    row_adder = 5
    disable_keyboard_check = True
    For Each tc In test_case_search_results
        Application.Range("A" & 2).value = tc.id
        Application.Range("B" & 2).value = tc.name
        If tc.case_procedures.count > 0 Then
            cps = tc.case_procedures.fetch
            For Each cp In cps
                Application.Range("A" & row_adder).value = cp.id
                Application.Range("B" & row_adder).value = cp.order_no
                Application.Range("C" & row_adder).value = cp.test_procedure.first.name
                row_adder = row_adder + 1
            Next
        End If
    Next
    border_and_color
    test_case_lock
    Application.ScreenUpdating = True
    disable_keyboard_check = False
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Sub test_case_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="TestCasesLock", Range:=Range( _
        "B2")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Application.Columns("A:C").EntireColumn.AutoFit
End Sub

Public Sub test_case_unlock()
    Application.ActiveSheet.Unprotect
    Application.Columns("A:G").EntireColumn.AutoFit
End Sub

Public Function test_case_step_new()
    Set new_search = New form_search
    new_search.init TEST_PROCEDURE_TAB, False
    new_search.Show 1
    If Not IsNull(test_procedure_search_results) Then
         Set cur_tc = TestCaseFactory(CDbl(Application.Range("A2").value))
        If cur_tc.count > 0 Then
            Set new_cp = New CaseProcedure
            new_cp.tc_id = cur_tc.first.id
            new_cp.order_no = cur_tc.first.case_procedures.count + 1
            new_cp.tp_id = test_procedure_search_results(0).id
            new_cp.save
            test_case_unlock
            Application.ActiveCell.value = new_cp.id
            Application.Range("B" & Application.ActiveCell.row).value = new_cp.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_procedure_search_results(0).name
            border_and_color
            test_case_lock
        End If
    End If
End Function

Public Function test_case_step_insert()
    Set new_search = New form_search
    new_search.init TEST_PROCEDURE_TAB, False
    new_search.Show 1
    If Not IsNull(test_procedure_search_results) Then
         Set cur_tc = TestCaseFactory(CDbl(Application.Range("A2").value))
        If cur_tc.count > 0 Then
            Set new_cp = New CaseProcedure
            new_cp.tc_id = cur_tc.first.id
            new_cp.order_no = cur_tc.first.case_procedures.count + 1
            new_cp.tp_id = test_procedure_search_results(0).id
            new_cp.save
            test_case_unlock
            Application.rows(Application.ActiveCell.row + 1).Insert Shift:=xlUp
            cur_selected_row = Application.ActiveCell.row + 1
            Application.Range("A" & Application.ActiveCell.row + 1).Select
            Application.ActiveCell.value = new_cp.id
            Application.Range("B" & Application.ActiveCell.row).value = new_cp.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_procedure_search_results(0).name
            Application.ScreenUpdating = False
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_cp = get_cp
                If cur_cp.count > 0 Then
                    Set update_cp = cur_cp.first
                    update_cp.order_no = i - 4
                    update_cp.save
                    Application.Range("B" & i).value = update_cp.order_no
                End If
            Next
            border_and_color
            Application.Range("A" & cur_selected_row).Select
            test_case_lock
        End If
    End If
    
End Function

Public Function test_case_step_delete()
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_cp = CaseProcedureFactory(CDbl(Application.ActiveCell.value))
        If cur_cp.count > 0 Then
            test_case_unlock
            cur_cp.first.delete
            Application.rows(Application.ActiveCell.row).delete Shift:=xlUp
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_cp = get_cp
                If cur_cp.count > 0 Then
                    Set update_cp = cur_cp.first
                    update_cp.order_no = i - 4
                    update_cp.save
                    Application.Range("B" & i).value = update_cp.order_no
                End If
            Next
            border_and_color
            test_case_lock
        End If
    End If
End Function

Public Function test_case_step_order()
    cur_selected_row = Application.ActiveCell.row
    new_order_no = Val(InputBox("Enter new step order number", "Step Order", cur_selected_row - 4))
    If cur_selected_row = new_order_no Then Exit Function
    If new_order_no <= 0 Then Exit Function
    sheet_size = get_last_row_column
    test_case_unlock
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
        Set cur_cp = get_cp
        If cur_cp.count > 0 Then
            Set update_cp = cur_cp.first
            update_cp.order_no = i - 4
            update_cp.save
            Application.Range("B" & i).value = update_cp.order_no
        End If
    Next
    Application.ScreenUpdating = True
    border_and_color
    test_case_lock
End Function

Public Function test_case_step_test_procedure()
    Set new_search = New form_search
    new_search.init TEST_PROCEDURE_TAB, False
    new_search.Show 1
    If Not IsNull(test_procedure_search_results) Then
         Set cur_cp = CaseProcedureFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
        If cur_cp.count > 0 Then
            Set update_cp = cur_cp.first
            update_cp.tp_id = test_procedure_search_results(0).id
            update_cp.save
            test_case_unlock
            Application.ActiveCell.value = update_cp.id
            Application.Range("C" & Application.ActiveCell.row).value = test_procedure_search_results(0).name
            border_and_color
            test_case_lock
        End If
    End If
End Function

Private Function get_cp()
    Set get_cp = CaseProcedureFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
End Function

Private Function border_and_color()
    
    test_case_unlock
    sheet_size = get_last_row_column
    ApplyBorder ("A" & 4 & ":C" & sheet_size(0))
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyAltColor 5, sheet_size(0), "A", "C"
    test_case_lock
    
End Function


