Attribute VB_Name = "module_test_procedure"
Public test_procedure_search_results
Public test_procedure_option_search_results

Sub test_procedure_new()
On Error GoTo err_handler
    new_tp_name = Trim(InputBox("Enter new Test Procedure name"))
    If new_tp_name <> "" Then
        Set new_tp = New TestProcedure
        new_tp.name = new_tp_name
        new_tp.save
        test_procedure_initial_state
        test_procedure_unlock
        Application.Range("A2").value = new_tp.id
        Application.Range("B2").value = new_tp.name
        test_procedure_lock
        border_and_color
        test_procedure_search_results = Null
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
    test_procedure_initial_state
End Sub

Sub test_procedure_clear()
    test_procedure_clear_sheet
    test_procedure_search_results = Null
End Sub

Sub test_procedure_clear_sheet()
    test_procedure_initial_state
End Sub

Sub test_procedure_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init TEST_PROCEDURE_TAB, False
    new_search.Show 1
    If Not IsNull(test_procedure_search_results) Then test_procedure_show
    
Exit Sub
err_handler:
    MsgBox Err.description

End Sub

Sub test_procedure_update()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        Set cur_tp = TestProcedureFactory(CDbl(Application.Range("A2").value))
        If cur_tp.count > 0 Then
            Set update_tp = cur_tp.first
            update_tp.name = Trim(Application.Range("B2").value)
            update_tp.save
            MsgBox update_tp.name & " updated successfully!"
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_procedure_delete()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_tp = TestProcedureFactory(CDbl(Application.Range("A2").value))
        If cur_tp.count > 0 Then
            Set update_tp = cur_tp.first
            update_tp.delete
            MsgBox update_tp.name & " deleted successfully!"
            test_procedure_initial_state
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_procedure_initial_state()
    test_procedure_unlock
    SheetClear
    Application.Range("A1").value = "ID"
    Application.Range("B1").value = "Test Procedure Name"
    
    Application.Range("A4").value = "ID"
    Application.Range("B4").value = "Step Number"
    Application.Range("C4").value = "Step Keyword"
    Application.Range("D4").value = "Test Object"
    Application.Range("E4").value = "Data Input"
    Application.Range("F4").value = "Data Output"
    Application.Range("G4").value = "Step Option"
    
    Application.Columns("A:G").EntireColumn.AutoFit
    ApplyBorder "A1:B2"
    ApplyColor "A1:" & "B" & 1, colorBlack
    ApplyFontColor "A1:" & "B" & 1, colorWhite
    
    ApplyBorder "A4:G5"
    ApplyColor "A4:" & "G" & 4, colorBlack
    ApplyFontColor "A4:" & "G" & 4, colorWhite
    FreezePane 4
    test_procedure_lock
    Application.Range("A2").Select
End Sub
    
Public Sub test_procedure_show()
    'On Error GoTo err_handler
    Application.ScreenUpdating = False
    test_procedure_clear_sheet
    test_procedure_unlock
    row_adder = 5
    disable_keyboard_check = True
    For Each o In test_procedure_search_results
        Application.Range("A" & 2).value = o.id
        Application.Range("B" & 2).value = o.name
        If o.procedure_steps.count > 0 Then
            steps = o.procedure_steps.fetch
            For Each step In steps
                Application.Range("A" & row_adder).value = step.id
                Application.Range("B" & row_adder).value = step.order_no
                Application.Range("C" & row_adder).value = step.keyword_name
                If step.test_object.count > 0 Then
                    Application.Range("D" & row_adder).value = step.test_object.first.name
                End If
                If step.links.count > 0 Then
                    If Not IsNullOrEmpty(step.links.first.data_value_in) Then
                        Application.Range("E" & row_adder).value = step.links.first.data_value_in
                    Else
                        If step.links.first.test_data_in.count > 0 Then
                            Application.Range("E" & row_adder).value = step.links.first.test_data_in.first.name
                        End If
                    End If
                    If Not IsNullOrEmpty(step.links.first.data_value_out) Then
                        Application.Range("F" & row_adder).value = step.links.first.data_value_out
                    Else
                        If step.links.first.test_data_out.count > 0 Then
                            Application.Range("F" & row_adder).value = step.links.first.test_data_out.first.name
                        End If
                    End If
                End If
                Columns("G:G").ColumnWidth = 100
                If step.test_option_links.count > 0 Then
                    For Each opt In step.test_option_links.fetch
                        Application.Range("G" & row_adder).value = Application.Range("G" & row_adder).value & opt.test_option.first.name & " := " & opt.item & vbCrLf
                    Next
                End If
                If Len(Application.Range("G" & row_adder).value) > 0 Then
                    Application.Range("G" & row_adder).value = Left(Application.Range("G" & row_adder).value, Len(Application.Range("G" & row_adder).value) - 1)
                End If
                row_adder = row_adder + 1
            Next
        End If
    Next
    border_and_color
    test_procedure_lock
    Application.ScreenUpdating = True
    disable_keyboard_check = False
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Sub test_procedure_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="TestProceduresLock", Range:=Range( _
        "B2,E:F")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Application.Columns("A:G").EntireColumn.AutoFit
End Sub

Public Sub test_procedure_unlock()
    Application.ActiveSheet.Unprotect
    Application.Columns("A:G").EntireColumn.AutoFit
End Sub

Public Sub test_procedure_step_clear_object()
    test_procedure_unlock
    Set cur_ps = get_ps
    If cur_ps.count > 0 Then
        Application.ActiveCell.value = ""
        Set update_ps = cur_ps.first
        update_ps.to_id = Null
        update_ps.save
    End If
    border_and_color
    test_procedure_lock
End Sub

Public Sub test_procedure_step_object()
    Set new_search = New form_search
    new_search.init TEST_OBJECT_TAB, False
    new_search.Show 1
    If Not IsNull(test_object_search_results) Then
        test_procedure_unlock
        Application.ActiveCell.value = test_object_search_results(0).name
        test_procedure_lock
        Set update_ps = get_ps.first
        update_ps.to_id = test_object_search_results(0).id
        update_ps.save
    End If
    border_and_color
End Sub

Public Sub test_procedure_step_clear_data_in()
    test_procedure_unlock
    Set cur_link = get_links
    If cur_link.count > 0 Then
        disable_keyboard_check = True
        Application.ActiveCell.value = ""
        disable_keyboard_check = False
        Set update_link = cur_link.first
        update_link.data_value_id_in = Null
        update_link.data_value_in = Null
        update_link.save
    Else
        Set new_link = New TestScenarioLink
        new_link.tp_id = CDbl(Application.Range("A" & 2).value)
        new_link.ps_id = CDbl(Application.Range("A" & Application.ActiveCell.row).value)
        new_link.data_value_id_in = Null
        new_link.data_value_in = Null
        new_link.save
    End If
    border_and_color
    test_procedure_lock
End Sub

Public Sub test_procedure_step_clear_data_out()
    test_procedure_unlock
    Set cur_link = get_links
    If cur_link.count > 0 Then
        disable_keyboard_check = True
        Application.ActiveCell.value = ""
        disable_keyboard_check = False
        Set update_link = cur_link.first
        update_link.data_value_id_out = Null
        update_link.data_value_out = Null
        update_link.save
    Else
        Set new_link = New TestScenarioLink
        new_link.tp_id = CDbl(Application.Range("A" & 2).value)
        new_link.ps_id = CDbl(Application.Range("A" & Application.ActiveCell.row).value)
        new_link.data_value_id_out = Null
        new_link.data_value_out = Null
        new_link.save
    End If
    border_and_color
    test_procedure_lock
End Sub

Public Sub test_procedure_step_data_in()
    Set new_search = New form_search
    new_search.init TEST_DATA_TAB, False
    new_search.Show 1
    If Not IsNull(test_data_search_results) Then
        test_procedure_unlock
        disable_keyboard_check = True
        Application.ActiveCell.value = test_data_search_results(0).name
        disable_keyboard_check = False
        test_procedure_lock
        Set cur_link = get_links
        If cur_link.count > 0 Then
            Set update_link = cur_link.first
            update_link.data_value_id_in = test_data_search_results(0).id
            update_link.data_value_in = Null
            update_link.save
        Else
            Set new_link = New TestScenarioLink
            new_link.tp_id = CDbl(Application.Range("A" & 2).value)
            new_link.ps_id = CDbl(Application.Range("A" & Application.ActiveCell.row).value)
            new_link.data_value_id_in = test_data_search_results(0).id
            new_link.save
        End If
        border_and_color
    End If
End Sub

Public Sub test_procedure_step_data_out()
    Set new_search = New form_search
    new_search.init TEST_DATA_TAB, False
    new_search.Show 1
    If Not IsNull(test_data_search_results) Then
        test_procedure_unlock
        disable_keyboard_check = True
        Application.ActiveCell.value = test_data_search_results(0).name
        disable_keyboard_check = False
        test_procedure_lock
        Set cur_link = get_links
        If cur_link.count > 0 Then
            Set update_link = cur_link.first
            update_link.data_value_id_out = test_data_search_results(0).id
            update_link.data_value_out = Null
            update_link.save
        Else
            Set new_link = New TestScenarioLink
            new_link.tp_id = CDbl(Application.Range("A" & 2).value)
            new_link.ps_id = CDbl(Application.Range("A" & Application.ActiveCell.row).value)
            new_link.data_value_id_out = test_data_search_results(0).id
            new_link.save
        End If
        border_and_color
    End If
End Sub

Public Sub test_procedure_step_data_in_value(row)
    If Not IsEmpty(Application.Range("A" & row).value) Then
        Set cur_link = get_target_links(row)
        Set find_data = TestDataFactory(Array("name", "=", "'" & Application.Range("E" & row).value & "'"))
        If cur_link.count > 0 Then
            Set update_link = cur_link.first
            If find_data.count > 0 Then
                update_link.data_value_id_in = find_data.first.id
                update_link.data_value_in = Null
            Else
                update_link.data_value_id_in = Null
                update_link.data_value_in = Application.Range("E" & row).value
            End If
            update_link.save
        Else
            Set new_link = New TestScenarioLink
            new_link.tp_id = CDbl(Application.Range("A" & 2).value)
            new_link.ps_id = CDbl(Application.Range("A" & Application.ActiveCell.row).value)
            If find_data.count > 0 Then
                new_link.data_value_id_in = find_data.first.id
                new_link.data_value_in = Null
            Else
                new_link.data_value_id_in = Null
                new_link.data_value_in = Application.Range("E" & row).value
            End If
            new_link.save
        End If
        border_and_color
        Application.Range("E" & row).Select
    End If
End Sub

Public Sub test_procedure_step_data_out_value(row)
    If Not IsEmpty(Application.Range("A" & row).value) Then
        Set cur_link = get_target_links(row)
        Set find_data = TestDataFactory(Array("name", "=", "'" & Application.Range("F" & row).value & "'"))
        If cur_link.count > 0 Then
            Set update_link = cur_link.first
            If find_data.count > 0 Then
                update_link.data_value_id_out = find_data.first.id
                update_link.data_value_out = Null
            Else
                update_link.data_value_id_out = Null
                update_link.data_value_out = Application.Range("F" & row).value
            End If
            update_link.save
        Else
            Set new_link = New TestScenarioLink
            new_link.tp_id = CDbl(Application.Range("A" & 2).value)
            new_link.ps_id = CDbl(Application.Range("A" & Application.ActiveCell.row).value)
            If find_data.count > 0 Then
                new_link.data_value_id_out = find_data.first.id
                new_link.data_value_out = Null
            Else
                new_link.data_value_id_out = Null
                new_link.data_value_out = Application.Range("F" & row).value
            End If
            new_link.save
        End If
        border_and_color
        Application.Range("F" & row).Select
    End If
End Sub

Public Sub test_procedure_step_option()
    On Error GoTo err_handler
    test_procedure_option_search_results = Null
    Set new_options = New form_options
    new_options.init CDbl(Application.Range("A" & Application.ActiveCell.row).value)
    new_options.Show 1
    If Not IsNull(test_procedure_option_search_results) Then
        test_option_unlock
        Application.ActiveCell.value = Join(test_procedure_option_search_results, vbCrLf)
        test_option_lock
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Function test_procedure_step_new()
    new_keyword = InputBox("Enter keyword name")
    If Trim(new_keyword) <> "" Then
        Set cur_tp = TestProcedureFactory(CDbl(Application.Range("A2").value))
        If cur_tp.count > 0 Then
            Set new_ps = New ProcedureStep
            new_ps.tp_id = cur_tp.first.id
            new_ps.order_no = cur_tp.first.procedure_steps.count + 1
            new_ps.keyword_name = new_keyword
            new_ps.save
            test_procedure_unlock
            Application.ActiveCell.value = new_ps.id
            Application.Range("B" & Application.ActiveCell.row).value = new_ps.order_no
            Application.Range("C" & Application.ActiveCell.row).value = new_ps.keyword_name
            border_and_color
            test_procedure_lock
        End If
    End If
End Function

Public Function test_procedure_step_insert()
    new_keyword = Trim(InputBox("Enter keyword name"))
    If Trim(new_keyword) <> "" Then
        Set cur_tp = TestProcedureFactory(CDbl(Application.Range("A2").value))
        If cur_tp.count > 0 Then
            Set new_ps = New ProcedureStep
            new_ps.tp_id = cur_tp.first.id
            new_ps.order_no = cur_tp.first.procedure_steps.count + 1
            new_ps.keyword_name = new_keyword
            new_ps.save
            test_procedure_unlock
            Application.rows(Application.ActiveCell.row + 1).Insert Shift:=xlUp
            cur_selected_row = Application.ActiveCell.row + 1
            Application.Range("A" & Application.ActiveCell.row + 1).Select
            Application.ActiveCell.value = new_ps.id
            Application.Range("B" & Application.ActiveCell.row).value = new_ps.order_no
            Application.Range("C" & Application.ActiveCell.row).value = new_ps.keyword_name
            Application.ScreenUpdating = False
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_ps = get_ps
                If cur_ps.count > 0 Then
                    Set update_ps = cur_ps.first
                    update_ps.order_no = i - 4
                    update_ps.save
                    Application.Range("B" & i).value = update_ps.order_no
                End If
            Next
            Application.ScreenUpdating = True
            border_and_color
            Application.Range("A" & cur_selected_row).Select
            test_procedure_lock
        End If
    End If
End Function

Public Function test_procedure_step_delete()
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_ps = ProcedureStepFactory(CDbl(Application.ActiveCell.value))
        If cur_ps.count > 0 Then
            test_procedure_unlock
            cur_ps.first.delete
            Application.rows(Application.ActiveCell.row).delete Shift:=xlUp
            For i = 5 To get_last_row_column(0)
                Application.Range("A" & i).Select
                Set cur_ps = get_ps
                If cur_ps.count > 0 Then
                    Set update_ps = cur_ps.first
                    update_ps.order_no = i - 4
                    update_ps.save
                    Application.Range("B" & i).value = update_ps.order_no
                End If
            Next
            border_and_color
            test_procedure_lock
        End If
    End If
End Function

Public Function test_procedure_step_order()
    cur_selected_row = Application.ActiveCell.row
    new_order_no = Val(InputBox("Enter new step order number", "Step Order", cur_selected_row - 4))
    If cur_selected_row = new_order_no Then Exit Function
    If new_order_no <= 0 Then Exit Function
    sheet_size = get_last_row_column
    test_procedure_unlock
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
        Set cur_ps = get_ps
        If cur_ps.count > 0 Then
            Set update_ps = cur_ps.first
            update_ps.order_no = i - 4
            update_ps.save
            Application.Range("B" & i).value = update_ps.order_no
        End If
    Next
    Application.ScreenUpdating = True
    border_and_color
    test_procedure_lock
End Function

Private Function get_ps()
    Set get_ps = ProcedureStepFactory(CDbl(Application.Range("A" & Application.ActiveCell.row).value))
End Function

Private Function get_options()
    Set get_options = TestOptionLinkFactory(Array("ps_id", CDbl(Application.Range("A" & Application.ActiveCell.row).value)))
End Function

Private Function get_links()
    Set get_links = TestScenarioLinkFactory(Array( _
                                            Array("ps_id", CDbl(Application.Range("A" & Application.ActiveCell.row).value)), _
                                            Array("tp_id", CDbl(Application.Range("A" & 2).value)), _
                                            Array("ts_id", Null), _
                                            Array("tc_id", Null), _
                                            Array("cs_id", Null), _
                                            Array("cp_id", Null) _
                                            ))
End Function

Private Function get_target_links(row)
    Set get_target_links = TestScenarioLinkFactory(Array( _
                                            Array("ps_id", CDbl(Application.Range("A" & row).value)), _
                                            Array("tp_id", CDbl(Application.Range("A" & 2).value)), _
                                            Array("ts_id", Null), _
                                            Array("tc_id", Null), _
                                            Array("cs_id", Null), _
                                            Array("cp_id", Null) _
                                            ))
End Function

Private Function border_and_color()
    
    test_procedure_unlock
    sheet_size = get_last_row_column
    ApplyBorder ("A" & 4 & ":G" & sheet_size(0))
    Application.Columns("A:G").EntireColumn.AutoFit
    ApplyAltColor 5, sheet_size(0), "A", "G"
    test_procedure_lock
    
End Function
