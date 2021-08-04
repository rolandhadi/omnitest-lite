Attribute VB_Name = "module_test_scenario"
Public test_scenario_search_results
Public test_scenario_option_search_results

Sub test_scenario_new()
On Error GoTo err_handler
    new_ts_name = Trim(InputBox("Enter Test Scenario name"))
    If new_ts_name <> "" Then
        Set new_ts = New TestScenario
        new_ts.name = new_ts_name
        new_ts.save
        test_scenario_initial_state
        test_scenario_unlock
        Application.Range("A2").value = new_ts.id
        Application.Range("B2").value = new_ts.name
        test_scenario_lock
        border_and_color
        test_scenario_search_results = Null
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
    test_scenario_initial_state
End Sub

Sub test_scenario_clear()
    test_scenario_clear_sheet
    test_scenario_search_results = Null
End Sub

Sub test_scenario_clear_sheet()
    test_scenario_initial_state
End Sub

Sub test_scenario_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init TEST_SCENARIO_TAB, False
    new_search.Show 1
    If Not IsNull(test_scenario_search_results) Then test_scenario_show
    
Exit Sub
err_handler:
    MsgBox Err.description

End Sub

Sub test_scenario_update()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        Set new_ts = TestScenarioFactory(CDbl(Application.Range("A2").value))
        If new_ts.count > 0 Then
            Set update_ts = new_ts.first
            update_ts.name = Trim(Application.Range("B2").value)
            update_ts.save
            MsgBox update_ts.name & " updated successfully!"
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_scenario_delete()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set new_ts = TestScenarioFactory(CDbl(Application.Range("A2").value))
        If new_ts.count > 0 Then
            Set update_ts = new_ts.first
            update_ts.delete
            MsgBox update_ts.name & " deleted successfully!"
            test_scenario_initial_state
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_scenario_initial_state()
    test_scenario_unlock
    SheetClear
    Application.Range("A1").value = "ID"
    Application.Range("B1").value = "Test Scenario Name"
    
    Application.Range("A4").value = "ID"
    Application.Range("B4").value = "Order"
    Application.Range("C4").value = "Test Case"
    
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyBorder "A1:B2"
    ApplyColor "A1:" & "B" & 1, colorBlack
    ApplyFontColor "A1:" & "B" & 1, colorWhite
    
    ApplyBorder "A4:C5"
    ApplyColor "A4:" & "C" & 4, colorBlack
    ApplyFontColor "A4:" & "C" & 4, colorWhite
    FreezePane 4
    test_scenario_lock
    Application.Range("A2").Select
End Sub
    
Public Sub test_scenario_show()
    On Error GoTo err_handler
    Application.ScreenUpdating = False
    test_scenario_clear_sheet
    test_scenario_unlock
    row_adder = 5
    disable_keyboard_check = True
    For Each ts In test_scenario_search_results
        Application.Range("A" & 2).value = ts.id
        Application.Range("B" & 2).value = ts.name
        If ts.case_scenarios.count > 0 Then
            css = ts.case_scenarios.fetch
            For Each cs In css
                Application.Range("A" & row_adder).value = cs.id
                Application.Range("B" & row_adder).value = cs.order_no
                Application.Range("C" & row_adder).value = cs.test_case.first.name
                row_adder = row_adder + 1
            Next
        End If
    Next
    border_and_color
    test_scenario_lock
    Application.ScreenUpdating = True
    disable_keyboard_check = False
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Sub test_scenario_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="TestScenariosLock", Range:=Range( _
        "B2")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Application.Columns("A:C").EntireColumn.AutoFit
End Sub

Public Sub test_scenario_unlock()
    Application.ActiveSheet.Unprotect
    Application.Columns("A:G").EntireColumn.AutoFit
End Sub

Public Function test_scenario_step_new()
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set new_ts = TestScenarioFactory(CDbl(Application.Range("A2").value))
        If new_ts.count > 0 Then
            Set new_cs = New CaseScenario
            new_cs.ts_id = new_ts.first.id
            new_cs.order_no = new_ts.first.case_scenarios.count + 1
            new_cs.tc_id = test_case_search_results(0).id
            new_cs.save
            test_scenario_unlock
            Application.ActiveCell.value = new_cs.id
            Application.Range("B" & Application.ActiveCell.row).value = new_cs.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_case_search_results(0).name
            border_and_color
            test_scenario_lock
        End If
    End If
End Function

Public Function test_scenario_step_insert()
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set new_ts = TestScenarioFactory(CDbl(Application.Range("A2").value))
        If new_ts.count > 0 Then
            Set new_cs = New CaseScenario
            new_cs.ts_id = new_ts.first.id
            new_cs.order_no = new_ts.first.case_scenarios.count + 1
            new_cs.tc_id = test_case_search_results(0).id
            new_cs.save
            test_scenario_unlock
            Application.rows(Application.ActiveCell.row + 1).Insert Shift:=xlUp
            cur_selected_row = Application.ActiveCell.row + 1
            Application.Range("A" & Application.ActiveCell.row + 1).Select
            Application.ActiveCell.value = new_cs.id
            Application.Range("B" & Application.ActiveCell.row).value = new_cs.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_case_search_results(0).name
            Application.ScreenUpdating = False
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_cs = get_cs
                If cur_cs.count > 0 Then
                    Set update_cs = cur_cs.first
                    update_cs.order_no = i - 4
                    update_cs.save
                    Application.Range("B" & i).value = update_cs.order_no
                End If
            Next
            border_and_color
            Application.Range("A" & cur_selected_row).Select
            test_scenario_lock
        End If
    End If
    
End Function

Public Function test_scenario_step_delete()
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_cs = CaseScenarioFactory(CDbl(Application.ActiveCell.value))
        If cur_cs.count > 0 Then
            test_scenario_unlock
            cur_cs.first.delete
            Application.rows(Application.ActiveCell.row).delete Shift:=xlUp
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_cs = get_cs
                If cur_cs.count > 0 Then
                    Set update_cs = cur_cs.first
                    update_cs.order_no = i - 4
                    update_cs.save
                    Application.Range("B" & i).value = update_cs.order_no
                End If
            Next
            border_and_color
            test_scenario_lock
        End If
    End If
End Function

Public Function test_scenario_step_order()
    cur_selected_row = Application.ActiveCell.row
    new_order_no = Val(InputBox("Enter new step order number", "Step Order", cur_selected_row - 4))
    If cur_selected_row = new_order_no Then Exit Function
    If new_order_no <= 0 Then Exit Function
    sheet_size = get_last_row_column
    test_scenario_unlock
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
        Set cur_cs = get_cs
        If cur_cs.count > 0 Then
            Set update_cs = cur_cs.first
            update_cs.order_no = i - 4
            update_cs.save
            Application.Range("B" & i).value = update_cs.order_no
        End If
    Next
    Application.ScreenUpdating = True
    border_and_color
    test_scenario_lock
End Function

Public Function test_scenario_step_test_case()
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set cur_cs = CaseScenarioFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
        If cur_cs.count > 0 Then
            Set update_cs = cur_cs.first
            update_cs.tc_id = test_case_search_results(0).id
            update_cs.save
            test_scenario_unlock
            Application.ActiveCell.value = update_cs.id
            Application.Range("C" & Application.ActiveCell.row).value = test_case_search_results(0).name
            border_and_color
            test_scenario_lock
        End If
    End If
End Function

Private Function get_cs()
    Set get_cs = CaseScenarioFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
End Function

Private Function border_and_color()
    
    test_scenario_unlock
    sheet_size = get_last_row_column
    ApplyBorder ("A" & 4 & ":C" & sheet_size(0))
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyAltColor 5, sheet_size(0), "A", "C"
    test_scenario_lock
    
End Function




