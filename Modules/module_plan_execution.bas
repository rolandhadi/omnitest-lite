Attribute VB_Name = "module_plan_execution"
Public plan_execution_search_results
Public plan_execution_option_search_results

Sub plan_execution_new()
On Error GoTo err_handler
    new_pe_name = Trim(InputBox("Enter Plan Execution name"))
    If new_pe_name <> "" Then
        Set new_pe = New PlanExecution
        new_pe.name = new_pe_name
        new_pe.save
        plan_execution_initial_state
        plan_execution_unlock
        Application.Range("A2").value = new_pe.id
        Application.Range("B2").value = new_pe.name
        plan_execution_lock
        border_and_color
        plan_execution_search_results = Null
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
    plan_execution_initial_state
End Sub

Sub plan_execution_clear()
    plan_execution_clear_sheet
    plan_execution_search_results = Null
End Sub

Sub plan_execution_clear_sheet()
    plan_execution_initial_state
End Sub

Sub plan_execution_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init PLAN_EXECUTION_TAB, False
    new_search.Show 1
    If Not IsNull(plan_execution_search_results) Then plan_execution_show
    
Exit Sub
err_handler:
    MsgBox Err.description

End Sub

Sub plan_execution_update()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        Set new_pe = PlanExecutionFactory(CDbl(Application.Range("A2").value))
        If new_pe.count > 0 Then
            Set update_pe = new_pe.first
            update_pe.name = Trim(Application.Range("B2").value)
            If Trim(Application.Range("C2").value) = "Stop Execution" Then
                update_pe.next_action = 1
            Else
                update_pe.next_action = 2
            End If
            update_pe.save
            MsgBox update_pe.name & " updated successfully!"
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub plan_execution_delete()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set new_pe = PlanExecutionFactory(CDbl(Application.Range("A2").value))
        If new_pe.count > 0 Then
            Set update_pe = new_pe.first
            update_pe.delete
            MsgBox update_pe.name & " deleted successfully!"
            plan_execution_initial_state
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub plan_execution_initial_state()
    plan_execution_unlock
    SheetClear
    Application.Range("A1").value = "ID"
    Application.Range("B1").value = "Plan Execution Name"
    Application.Range("C1").value = "Next Action"
    
    Application.Range("A4").value = "ID"
    Application.Range("B4").value = "Order"
    Application.Range("C4").value = "Test Scenario"
    
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyBorder "A1:C2"
    ApplyColor "A1:" & "C" & 1, colorBlack
    ApplyFontColor "A1:" & "C" & 1, colorWhite
    
    ApplyBorder "A4:C5"
    ApplyColor "A4:" & "C" & 4, colorBlack
    ApplyFontColor "A4:" & "C" & 4, colorWhite
    FreezePane 4
    plan_execution_lock
    Application.Range("A2").Select
End Sub
    
Public Sub plan_execution_show()
    On Error GoTo err_handler
    Application.ScreenUpdating = False
    plan_execution_clear_sheet
    plan_execution_unlock
    row_adder = 5
    disable_keyboard_check = True
    For Each pe In plan_execution_search_results
        Application.Range("A" & 2).value = pe.id
        Application.Range("B" & 2).value = pe.name
        If pe.next_action = 1 Then
            Application.Range("C" & 2).value = "Stop Execution"
        Else
            Application.Range("C" & 2).value = "Continue Execution"
        End If
        If pe.execution_scenarios.count > 0 Then
            ess = pe.execution_scenarios.fetch
            For Each es In ess
                Application.Range("A" & row_adder).value = es.id
                Application.Range("B" & row_adder).value = es.order_no
                Application.Range("C" & row_adder).value = es.test_scenario.first.name
                row_adder = row_adder + 1
            Next
        End If
    Next
    border_and_color
    plan_execution_lock
    Application.ScreenUpdating = True
    disable_keyboard_check = False
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Sub plan_execution_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="PlanExecutionsLock", Range:=Range( _
        "B2")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Application.Columns("A:C").EntireColumn.AutoFit
End Sub

Public Sub plan_execution_unlock()
    Application.ActiveSheet.Unprotect
    Application.Columns("A:G").EntireColumn.AutoFit
End Sub

Public Function plan_execution_step_new()
    Set new_search = New form_search
    new_search.init TEST_SCENARIO_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set new_pe = PlanExecutionFactory(CDbl(Application.Range("A2").value))
        If new_pe.count > 0 Then
            Set new_es = New ExecutionScenario
            new_es.pe_id = new_pe.first.id
            new_es.ts_id = test_scenario_search_results(0).id
            new_es.order_no = new_pe.first.execution_scenarios.count + 1
            new_es.save
            plan_execution_unlock
            Application.ActiveCell.value = new_es.id
            Application.Range("B" & Application.ActiveCell.row).value = new_es.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_scenario_search_results(0).name
            border_and_color
            plan_execution_lock
        End If
    End If
End Function

Public Function plan_execution_step_insert()
    Set new_search = New form_search
    new_search.init TEST_CASE_TAB, False
    new_search.Show 1
    If Not IsNull(test_case_search_results) Then
         Set new_pe = PlanExecutionFactory(CDbl(Application.Range("A2").value))
        If new_pe.count > 0 Then
            Set new_es = New ExecutionScenario
            new_es.pe_id = new_pe.first.id
            new_es.ts_id = test_scenario_search_results(0).id
            new_es.order_no = new_pe.first.execution_scenarios.count + 1
            new_es.save
            plan_execution_unlock
            Application.rows(Application.ActiveCell.row + 1).Insert Shift:=xlUp
            cur_selected_row = Application.ActiveCell.row + 1
            Application.Range("A" & Application.ActiveCell.row + 1).Select
            Application.ActiveCell.value = new_es.id
            Application.Range("B" & Application.ActiveCell.row).value = new_es.order_no
            Application.Range("C" & Application.ActiveCell.row).value = test_case_search_results(0).name
            Application.ScreenUpdating = False
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_es = get_es
                If cur_es.count > 0 Then
                    Set update_es = cur_es.first
                    update_es.order_no = i - 4
                    update_es.save
                    Application.Range("B" & i).value = update_es.order_no
                End If
            Next
            border_and_color
            Application.Range("A" & cur_selected_row).Select
            plan_execution_lock
        End If
    End If
    
End Function

Public Function plan_execution_step_delete()
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_es = ExecutionScenarioFactory(CDbl(Application.ActiveCell.value))
        If cur_es.count > 0 Then
            plan_execution_unlock
            cur_es.first.delete
            Application.rows(Application.ActiveCell.row).delete Shift:=xlUp
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_es = get_es
                If cur_es.count > 0 Then
                    Set update_es = cur_es.first
                    update_es.order_no = i - 4
                    update_es.save
                    Application.Range("B" & i).value = update_es.order_no
                End If
            Next
            border_and_color
            plan_execution_lock
        End If
    End If
End Function

Public Function plan_execution_step_order()
    cur_selected_row = Application.ActiveCell.row
    new_order_no = Val(InputBox("Enter new step order number", "Step Order", cur_selected_row - 4))
    If cur_selected_row = new_order_no Then Exit Function
    If new_order_no <= 0 Then Exit Function
    sheet_size = get_last_row_column
    plan_execution_unlock
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
        Set cur_es = get_es
        If cur_es.count > 0 Then
            Set update_es = cur_es.first
            update_es.order_no = i - 4
            update_es.save
            Application.Range("B" & i).value = update_es.order_no
        End If
    Next
    Application.ScreenUpdating = True
    border_and_color
    plan_execution_lock
End Function

Public Function plan_execution_step_test_case()
    Set new_search = New form_search
    new_search.init TEST_SCENARIO_TAB, False
    new_search.Show 1
    If Not IsNull(test_scenario_search_results) Then
         Set cur_es = ExecutionScenarioFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
        If cur_es.count > 0 Then
            Set update_es = cur_es.first
            update_es.ts_id = test_scenario_search_results(0).id
            update_es.save
            plan_execution_unlock
            Application.ActiveCell.value = update_es.id
            Application.Range("C" & Application.ActiveCell.row).value = test_scenario_search_results(0).name
            border_and_color
            plan_execution_lock
        End If
    End If
End Function

Public Function plan_execution_stop_exec()
    plan_execution_unlock
    Application.Range("C2").value = "Stop Execution"
    border_and_color
    Application.Range("C2").Select
    plan_execution_lock
End Function

Public Function plan_execution_continue_exec()
    plan_execution_unlock
    Application.Range("C2").value = "Continue Execution"
    border_and_color
    Application.Range("C2").Select
    plan_execution_lock
End Function

Private Function get_es()
    Set get_es = ExecutionScenarioFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
End Function

Private Function border_and_color()
    
    plan_execution_unlock
    sheet_size = get_last_row_column
    ApplyBorder ("A" & 4 & ":C" & sheet_size(0))
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyAltColor 5, sheet_size(0), "A", "C"
    plan_execution_lock
    
End Function






