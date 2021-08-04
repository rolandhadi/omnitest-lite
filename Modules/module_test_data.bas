Attribute VB_Name = "module_test_data"
Public test_data_search_results

Sub test_data_new()
    
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to add new test data?", vbYesNo) = vbYes Then
        If TestDataFactory(Array("name", "'" & Application.Range("C" & Application.ActiveCell.row).value & "'")).count = 0 Then
            Set new_object = New TestData
            new_object.parent_or_new Trim(Application.Range("B" & Application.ActiveCell.row).value)
            new_object.name = Trim(Application.Range("C" & Application.ActiveCell.row).value)
            new_object.save
            test_data_unlock
            Application.Range("A" & Application.ActiveCell.row).value = "TD" & new_object.id
            draw_placeholder get_last_row_column(0) + 2
            test_data_lock
            MsgBox new_object.name & " was updated succesfully!"
        Else
            MsgBox "test data name already exist"
        End If
    End If
Exit Sub
err_handler:
    MsgBox Err.description
    
End Sub

Sub test_data_clear()
    test_data_clear_sheet
    test_data_search_results = Null
End Sub

Sub test_data_clear_sheet()
    test_data_initial_state
End Sub

Sub test_data_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init TEST_DATA_TAB, True
    new_search.Show 1
    If Not IsNull(test_data_search_results) Then test_data_show
    
Exit Sub
err_handler:
    MsgBox Err.description
    
End Sub

Sub test_data_update()

    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        Set cur_object = TestDataFactory(CDbl(Replace(Application.Range("A" & Application.ActiveCell.row).value, "TD", ""))).first
        cur_object.parent_or_new Trim(Application.Range("B" & Application.ActiveCell.row).value)
        cur_object.name = Trim(Application.Range("C" & Application.ActiveCell.row).value)
        cur_object.save
        MsgBox cur_object.name & " was updated succesfully!"
    End If
Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_data_value_update()

    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        Set cur_object_value = TestDataValueFactory(CDbl(Replace(Application.Range("A" & Application.ActiveCell.row).value, "DV", ""))).first
        cur_object_value.iteration = Trim(Application.Range("B" & Application.ActiveCell.row).value)
        cur_object_value.item = Trim(Application.Range("C" & Application.ActiveCell.row).value)
        cur_object_value.save
        MsgBox cur_object_value.item & " value was updated succesfully!"
    End If
Exit Sub
err_handler:
    MsgBox Err.description
    
End Sub

Sub test_data_delete()

    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_object = TestDataFactory(CDbl(Replace(Application.Range("A" & Application.ActiveCell.row).value, "TD", ""))).first
        cur_object.delete
        start_row = Application.ActiveCell.row - 1
        end_row = start_row
        For i = start_row + 1 To 100
            end_row = end_row + 1
            If Trim(Application.Range("A" & i).value) = "" Then
                Exit For
            End If
        Next
        test_data_unlock
        Application.rows(start_row & ":" & end_row - 1).delete Shift:=xlUp
        test_data_lock
        MsgBox cur_object.name & " was updated succesfully!"
    End If

Exit Sub
err_handler:
    MsgBox Err.description
    
End Sub

Sub test_data_value_delete()

    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        Set cur_object_value = TestDataValueFactory(CDbl(Replace(Application.Range("A" & Application.ActiveCell.row).value, "DV", ""))).first
        cur_object_value.delete
        test_data_unlock
        Application.rows(Application.ActiveCell.row).delete Shift:=xlUp
        test_data_lock
        MsgBox cur_object_value.item & " was updated succesfully!"
    End If

Exit Sub
err_handler:
    MsgBox Err.description
    
End Sub

Sub test_data_initial_state()
    test_data_unlock
    SheetClear
    draw_placeholder 1
    Application.Range("A2").Select
End Sub

Public Sub test_data_show()
    
    On Error GoTo err_handler
    Application.ScreenUpdating = False
    test_data_clear_sheet
    test_data_unlock
    row_adder = 0
    
    For Each o In test_data_search_results
        
        sheet_size = get_last_row_column
        draw_placeholder sheet_size(0) + 2
        test_data_unlock
        Application.Range("A" & sheet_size(0) - 2).value = "TD" & o.id
        If o.parent.count > 0 Then
            Application.Range("B" & sheet_size(0) - 2).value = o.parent.first.name
        End If
        Application.Range("C" & sheet_size(0) - 2).value = o.name
        Application.Range("C" & sheet_size(0) - 2).Font.Bold = True
        
        If o.values.count > 0 Then
            values = o.values.fetch
            row_adder = 0
            If IsArray(values) Then
                For Each v In values
                    Application.rows(sheet_size(0) + 1).Insert Shift:=xlUp
                    Application.Range("A" & sheet_size(0) + row_adder).value = "DV" & v.id
                    Application.Range("B" & sheet_size(0) + row_adder).value = v.iteration
                    Application.Range("C" & sheet_size(0) + row_adder).value = v.item
                    ApplyBorder "A" & sheet_size(0) + row_adder & ":" & "C" & sheet_size(0) + row_adder
                    Application.Range("B" & sheet_size(0) + row_adder).Font.Size = 10
                    row_adder = row_adder + 1
                    Application.Columns("A:D").EntireColumn.AutoFit
                Next
            End If
        End If
            
    Next
    test_data_lock
    Application.ScreenUpdating = True
Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Public Sub test_data_value_add()

    On Error GoTo err_handler
    iteration = InputBox("Enter data iteration number", "Add Data Iteration", "1")
    If Trim(iteration) = "" Then Exit Sub
    Set new_data_value = New TestDataValue
    Set find_data_value = TestDataFactory(CDbl(Replace(Application.Range("A" & Application.ActiveCell.row - 1).value, "TD", ""))).first
    new_data_value.td_id = find_data_value.id
    new_data_value.iteration = iteration
    new_data_value.save
    test_data_unlock
    Application.rows(Application.ActiveCell.row + 1).Insert Shift:=xlDown
    Application.Range("A" & Application.ActiveCell.row + 1 & ":" & "C" & Application.ActiveCell.row + 1).Interior.Color = vbWhite
    Application.Range("A" & Application.ActiveCell.row + 1 & ":" & "C" & Application.ActiveCell.row + 1).Font.Color = vbBlack
    Application.Range("A" & Application.ActiveCell.row + 1).value = "DV" & new_data_value.id
    Application.Range("B" & Application.ActiveCell.row + 1).value = new_data_value.iteration
    Application.Range("C" & Application.ActiveCell.row + 1).value = new_data_value.item
    test_data_lock
Exit Sub
err_handler:
    MsgBox Err.description
End Sub
    
Public Sub test_data_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="TestDatasLock", Range _
    :=Columns("B:F")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
End Sub
    
Public Sub test_data_unlock()
    Application.ActiveSheet.Unprotect
End Sub

Public Sub draw_placeholder(row)
    test_data_unlock
        
    Application.Range("A" & row).value = "Data ID"
    Application.Range("A" & row + 1).value = "NEW"
    Application.Range("B" & row).value = "Parent Name"
    Application.Range("C" & row).value = "Test Data Name"
    
    Application.Range("A" & row + 2).value = "Value ID"
    Application.Range("B" & row + 2).value = "Test Data Type"
    Application.Range("C" & row + 2).value = "Test Data Value"
    
    Application.Columns("A:C").EntireColumn.AutoFit
    ApplyBorder "A" & row & ":C" & row + 1
    ApplyColor "A" & row & ":C" & row, colorBlack
    ApplyFontColor "A" & row & ":C" & row, colorWhite
    
    Application.Columns("A:D").EntireColumn.AutoFit
    ApplyBorder "A" & row + 2 & ":C" & row + 3
    ApplyColor "A" & row + 2 & ":C" & row + 2, colorGray
    ApplyFontColor "A" & row + 2 & ":C" & row + 2, colorWhite
    
    Application.rows(row + 1 & ":" & row + 2).Group
    
    test_data_lock
    Application.Range("A2").Select
End Sub


