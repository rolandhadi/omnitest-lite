Attribute VB_Name = "module_test_scenario_data"
Public test_scenario_data_search_results
Public test_scenario_data_option_search_results

Sub test_scenario_data_clear()
    test_scenario_data_clear_sheet
    test_scenario_data_search_results = Null
End Sub

Sub test_scenario_data_clear_sheet()
    test_scenario_data_initial_state
End Sub

Sub test_scenario_data_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init TEST_SCENARIO_DATA_TAB, False
    new_search.Show 1
    If Not IsNull(test_scenario_data_search_results) Then test_scenario_data_show
    
Exit Sub
err_handler:
    MsgBox Err.description

End Sub

Sub test_scenario_data_update()
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

Sub test_scenario_data_delete()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set new_ts = TestScenarioFactory(CDbl(Application.Range("A2").value))
        If new_ts.count > 0 Then
            Set update_ts = new_ts.first
            update_ts.delete
            MsgBox update_ts.name & " deleted successfully!"
            test_scenario_data_initial_state
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_scenario_data_initial_state()
    test_scenario_data_unlock
    SheetClear
    Application.Range("A1").value = "ID"
    Application.Range("B1").value = "Test Scenario Name"
    
    Application.Range("A4").value = "ID"
    Application.Range("B4").value = "Test Case Order"
    Application.Range("C4").value = "Test Case Name"
    Application.Range("D4").value = "Test Procedure Order"
    Application.Range("E4").value = "Test Procedure Name"
    Application.Range("F4").value = "Step Number"
    Application.Range("G4").value = "Step Keyword"
    Application.Range("H4").value = "Test Object"
    Application.Range("I4").value = "Data Input"
    Application.Range("J4").value = "Data Output"
    
    Application.Columns("A:J").EntireColumn.AutoFit
    ApplyBorder "A1:B2"
    ApplyColor "A1:" & "B" & 1, colorBlack
    ApplyFontColor "A1:" & "B" & 1, colorWhite
    
    ApplyBorder "A4:J5"
    ApplyColor "A4:" & "J" & 4, colorBlack
    ApplyFontColor "A4:" & "J" & 4, colorWhite
    FreezePane 4
    test_scenario_data_lock
    Application.Range("A2").Select
End Sub
    
Public Sub test_scenario_data_show()
    On Error GoTo err_handler
    Application.ScreenUpdating = False
    test_scenario_data_clear_sheet
    test_scenario_data_unlock
    row_adder = 5
    disable_keyboard_check = True
    For Each ts In test_scenario_data_search_results
        Application.Range("A" & 2).value = ts.id
        Application.Range("B" & 2).value = ts.name
        If ts.case_scenarios.count > 0 Then
            css = ts.case_scenarios.fetch
            For Each cs In css
                Application.Range("B" & row_adder).value = cs.order_no
                Application.Range("C" & row_adder).value = cs.test_case.first.name
                If cs.test_case.first.case_procedures.count > 0 Then
                    cps = cs.test_case.first.case_procedures.fetch
                    For Each cp In cps
                        Application.Range("D" & row_adder).value = cp.order_no
                        Application.Range("E" & row_adder).value = cp.test_procedure.first.name
                        If cp.test_procedure.first.procedure_steps.count > 0 Then
                            pss = cp.test_procedure.first.procedure_steps.fetch
                            For Each ps In pss
                                Application.Range("B" & row_adder).value = cs.order_no
                                Application.Range("C" & row_adder).value = cs.test_case.first.name
                                Application.Range("D" & row_adder).value = cp.order_no
                                Application.Range("E" & row_adder).value = cp.test_procedure.first.name
                                Application.Range("F" & row_adder).value = ps.order_no
                                Application.Range("G" & row_adder).value = ps.keyword_name
                                If ps.test_object.count > 0 Then
                                    Application.Range("H" & row_adder).value = ps.test_object.first.name
                                End If
                                Set links = TestScenarioLinkFactory(Array( _
                                                                        Array("ts_id", CDbl(Application.Range("A" & 2).value)), _
                                                                        Array("tc_id", cs.test_case.first.id), _
                                                                        Array("tp_id", cp.test_procedure.first.id), _
                                                                        Array("cs_id", cs.id), _
                                                                        Array("ps_id", ps.id) _
                                                                    ) _
                                                                )
                                If links.count > 0 Then
                                    Set link = links.first
                                    Application.Range("A" & row_adder).value = link.id
                                    If Not IsNullOrEmpty(link.data_value_in) Then
                                        Application.Range("I" & row_adder).value = link.data_value_in
                                    Else
                                        If link.test_data_in.count > 0 Then
                                            Application.Range("I" & row_adder).value = link.test_data_in.first.name
                                        End If
                                    End If
                                    If Not IsNullOrEmpty(link.function_reference_in) Then
                                        If link.function_reference_in.count > 0 Then
                                            Application.Range("I" & row_adder).value = link.function_reference_in.first.values.first.name
                                        End If
                                    End If
                                    If Not IsNullOrEmpty(link.data_value_out) Then
                                        Application.Range("J" & row_adder).value = link.data_value_out
                                    Else
                                        If link.test_data_out.count > 0 Then
                                            Application.Range("J" & row_adder).value = link.test_data_out.first.name
                                        End If
                                    End If
                                Else
                                    Application.Range("A" & row_adder).value = CDbl(Application.Range("A" & 2).value) & "." & _
                                                                               cs.test_case.first.id & "." & _
                                                                               cp.test_procedure.first.id & "." & _
                                                                               cs.id & "." & _
                                                                               cp.id & "." & _
                                                                               ps.id

                                End If
                                row_adder = row_adder + 1
                            Next
                        End If
                    Next
                End If
            Next
        End If
    Next
    border_and_color
    test_scenario_data_lock
    Application.ScreenUpdating = True
    disable_keyboard_check = False
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Sub test_scenario_data_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="TestScenariosLock", Range:=Range( _
        "I:J")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Application.Columns("A:C").EntireColumn.AutoFit
End Sub

Public Sub test_scenario_data_unlock()
    Application.ActiveSheet.Unprotect
    Application.Columns("A:G").EntireColumn.AutoFit
End Sub

Public Sub test_scenario_data_step_clear_data_in()
    test_scenario_data_unlock
    cur_row = Application.ActiveCell.row
    If InStr(1, Application.Range("A" & Application.ActiveCell.row).value, ".") <= 0 Then
        Set cur_link = get_links
        If cur_link.count > 0 Then
            disable_keyboard_check = True
            Application.ActiveCell.value = ""
            disable_keyboard_check = False
            Set update_link = cur_link.first
            update_link.data_value_id_in = Null
            update_link.data_value_in = Null
            update_link.data_value_ref_id = Null
            update_link.save
        End If
    Else
        Set new_link = New TestScenarioLink
        ids = Split(Application.Range("A" & Application.ActiveCell.row).value, ".")
        new_link.ts_id = CDbl(ids(0))
        new_link.tc_id = CDbl(ids(1))
        new_link.tp_id = CDbl(ids(2))
        new_link.cs_id = CDbl(ids(3))
        new_link.cp_id = CDbl(ids(4))
        new_link.ps_id = CDbl(ids(5))
        new_link.data_value_id_in = Null
        new_link.save
        Application.Range("A" & Application.ActiveCell.row).value = new_link.id
    End If
    border_and_color
    test_scenario_data_lock
    Application.Range("I" & cur_row).Select
End Sub

Public Sub test_scenario_data_step_clear_data_out()
    test_scenario_data_unlock
    cur_row = Application.ActiveCell.row
    If InStr(1, Application.Range("A" & Application.ActiveCell.row).value, ".") <= 0 Then
        Set cur_link = get_links
        If cur_link.count > 0 Then
            disable_keyboard_check = True
            Application.ActiveCell.value = ""
            disable_keyboard_check = False
            Set update_link = cur_link.first
            update_link.data_value_id_out = Null
            update_link.data_value_out = Null
            update_link.data_value_ref_id = Null
            update_link.save
        End If
    Else
        Set new_link = New TestScenarioLink
        ids = Split(Application.Range("A" & Application.ActiveCell.row).value, ".")
        new_link.ts_id = CDbl(ids(0))
        new_link.tc_id = CDbl(ids(1))
        new_link.tp_id = CDbl(ids(2))
        new_link.cs_id = CDbl(ids(3))
        new_link.cp_id = CDbl(ids(4))
        new_link.ps_id = CDbl(ids(5))
        new_link.data_value_id_out = Null
        new_link.save
        Application.Range("A" & Application.ActiveCell.row).value = new_link.id
    End If
    border_and_color
    test_scenario_data_lock
    Application.Range("J" & cur_row).Select
End Sub

Public Sub test_scenario_data_step_ref_in()
    Set new_search = New form_search
    new_search.init TEST_DATA_TAB, False
    new_search.Show 1
    If Not IsNull(test_data_search_results) Then
        test_scenario_data_unlock
        disable_keyboard_check = True
        Application.ActiveCell.value = test_data_search_results(0).name
        disable_keyboard_check = False
        cur_row = Application.ActiveCell.row
        If InStr(1, Application.Range("A" & cur_row).value, ".") <= 0 Then
            Set cur_link = get_target_links(Application.ActiveCell.row)
            If cur_link.count > 0 Then
                Set update_link = cur_link.first
                update_link.data_value_id_in = test_data_search_results(0).id
                update_link.data_value_ref_id = Null
                update_link.data_value_in = Null
                update_link.save
            End If
        Else
            Set new_link = New TestScenarioLink
            ids = Split(Application.Range("A" & cur_row).value, ".")
            new_link.ts_id = CDbl(ids(0))
            new_link.tc_id = CDbl(ids(1))
            new_link.tp_id = CDbl(ids(2))
            new_link.cs_id = CDbl(ids(3))
            new_link.cp_id = CDbl(ids(4))
            new_link.ps_id = CDbl(ids(5))
            new_link.data_value_id_in = test_data_search_results(0).id
            new_link.save
            Application.Range("A" & cur_row).value = new_link.id
        End If
        test_scenario_data_lock
        border_and_color
        Application.Range("I" & cur_row).Select
    End If
End Sub

Public Sub test_scenario_data_step_data_in()
    Set new_search = New form_search
    new_search.init TEST_DATA_TAB, False
    new_search.Show 1
    If Not IsNull(test_data_search_results) Then
        test_scenario_data_unlock
        disable_keyboard_check = True
        Application.ActiveCell.value = test_data_search_results(0).name
        disable_keyboard_check = False
        cur_row = Application.ActiveCell.row
        If InStr(1, Application.Range("A" & cur_row).value, ".") <= 0 Then
            Set cur_link = get_target_links(Application.ActiveCell.row)
            If cur_link.count > 0 Then
                Set update_link = cur_link.first
                update_link.data_value_id_in = test_data_search_results(0).id
                update_link.data_value_ref_id = Null
                update_link.data_value_in = Null
                update_link.save
            End If
        Else
            Set new_link = New TestScenarioLink
            ids = Split(Application.Range("A" & cur_row).value, ".")
            new_link.ts_id = CDbl(ids(0))
            new_link.tc_id = CDbl(ids(1))
            new_link.tp_id = CDbl(ids(2))
            new_link.cs_id = CDbl(ids(3))
            new_link.cp_id = CDbl(ids(4))
            new_link.ps_id = CDbl(ids(5))
            new_link.data_value_id_in = test_data_search_results(0).id
            new_link.save
            Application.Range("A" & cur_row).value = new_link.id
        End If
        test_scenario_data_lock
        border_and_color
        Application.Range("I" & cur_row).Select
    End If
End Sub

Public Sub test_scenario_data_step_data_out()
    Set new_search = New form_search
    new_search.init TEST_DATA_TAB, False
    new_search.Show 1
    If Not IsNull(test_data_search_results) Then
        test_scenario_data_unlock
        disable_keyboard_check = True
        Application.ActiveCell.value = test_data_search_results(0).name
        disable_keyboard_check = False
        cur_row = Application.ActiveCell.row
        If InStr(1, Application.Range("A" & cur_row).value, ".") <= 0 Then
            Set cur_link = get_target_links(Application.ActiveCell.row)
            If cur_link.count > 0 Then
                Set update_link = cur_link.first
                update_link.data_value_id_out = test_data_search_results(0).id
                update_link.data_value_out = Null
                update_link.save
            End If
        Else
            Set new_link = New TestScenarioLink
            ids = Split(Application.Range("A" & cur_row).value, ".")
            new_link.ts_id = CDbl(ids(0))
            new_link.tc_id = CDbl(ids(1))
            new_link.tp_id = CDbl(ids(2))
            new_link.cs_id = CDbl(ids(3))
            new_link.cp_id = CDbl(ids(4))
            new_link.ps_id = CDbl(ids(5))
            new_link.data_value_id_out = test_data_search_results(0).id
            new_link.save
            Application.Range("A" & cur_row).value = new_link.id
        End If
        test_scenario_data_lock
        border_and_color
        Application.Range("J" & cur_row).Select
    End If
End Sub

Public Sub test_scenario_data_step_data_in_value(row)
    If Not IsEmpty(Application.Range("A" & row).value) Then
        test_scenario_data_unlock
        cur_row = row
        If InStr(1, Application.Range("A" & cur_row).value, ".") <= 0 Then
            Set cur_link = get_target_links(row)
            Set find_data = TestDataFactory(Array("name", "=", "'" & Application.Range("I" & row).value & "'"))
            If cur_link.count > 0 Then
                Set update_link = cur_link.first
                If find_data.count > 0 Then
                    update_link.data_value_id_in = find_data.first.id
                    update_link.data_value_in = Null
                Else
                    update_link.data_value_id_in = Null
                    update_link.data_value_in = Application.Range("I" & row).value
                End If
                update_link.data_value_ref_id = Null
                update_link.save
            End If
        Else
            Set new_link = New TestScenarioLink
            ids = Split(Application.Range("A" & cur_row).value, ".")
            new_link.ts_id = CDbl(ids(0))
            new_link.tc_id = CDbl(ids(1))
            new_link.tp_id = CDbl(ids(2))
            new_link.cs_id = CDbl(ids(3))
            new_link.cp_id = CDbl(ids(4))
            new_link.ps_id = CDbl(ids(5))
            new_link.save
            Application.Range("A" & cur_row).value = new_link.id
        End If
        test_scenario_data_lock
        border_and_color
        Application.Range("I" & cur_row).Select
    End If
End Sub

Public Sub test_scenario_data_step_data_out_value(row)
    If Not IsEmpty(Application.Range("A" & row).value) Then
        test_scenario_data_unlock
        cur_row = row
        If InStr(1, Application.Range("A" & cur_row).value, ".") <= 0 Then
            Set cur_link = get_target_links(row)
            Set find_data = TestDataFactory(Array("name", "=", "'" & Application.Range("J" & row).value & "'"))
            If cur_link.count > 0 Then
                Set update_link = cur_link.first
                If find_data.count > 0 Then
                    update_link.data_value_id_out = find_data.first.id
                    update_link.data_value_out = Null
                Else
                    update_link.data_value_id_out = Null
                    update_link.data_value_out = Application.Range("J" & row).value
                End If
                update_link.save
            End If
        Else
            Set new_link = New TestScenarioLink
            ids = Split(Application.Range("A" & cur_row).value, ".")
            new_link.ts_id = CDbl(ids(0))
            new_link.tc_id = CDbl(ids(1))
            new_link.tp_id = CDbl(ids(2))
            new_link.cs_id = CDbl(ids(3))
            new_link.cp_id = CDbl(ids(4))
            new_link.ps_id = CDbl(ids(5))
            new_link.save
            Application.Range("A" & cur_row).value = new_link.id
        End If
        test_scenario_data_lock
        border_and_color
        Application.Range("J" & cur_row).Select
    End If
End Sub

Private Function get_cs()
    Set get_cs = CaseScenarioFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
End Function

Private Function get_links()
    Set get_links = TestScenarioLinkFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
End Function

Private Function get_target_links(rows)
    Set get_target_links = TestScenarioLinkFactory(CDbl(Application.Range("A" & rows).value))
End Function

Private Function border_and_color()
    
    test_scenario_data_unlock
    sheet_size = get_last_row_column
    ApplyBorder ("A" & 4 & ":J" & sheet_size(0))
    Application.Columns("A:J").EntireColumn.AutoFit
    ApplyAltColor 5, sheet_size(0), "A", "J"
    test_scenario_data_lock
    
End Function

Sub merge_rows(column_no)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For i = 5 To get_last_row_column(0)
        
    Next
    Application.Range("B7:B11").Select
    With Application.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Application.Selection.Merge
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
