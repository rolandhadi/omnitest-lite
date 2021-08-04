Attribute VB_Name = "module_test_option"
Public test_option_search_results

Sub test_option_clear()
    test_option_clear_sheet
    test_option_search_results = Null
End Sub

Sub test_option_clear_sheet()
    test_option_initial_state
End Sub

Sub test_option_find()

On Error GoTo err_handler
    Set new_search = New form_search
    new_search.init TEST_OPTION_TAB, True
    new_search.Show 1
    If Not IsNull(test_option_search_results) Then test_option_show
    
Exit Sub
err_handler:
    MsgBox Err.description
    
End Sub

Sub test_option_update()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to save your changes?", vbYesNo) = vbYes Then
        updated_count = 0
        For i = 2 To get_last_row_column(0)
            If UCase(Trim(Application.Range("D" & i).value)) = "Y" And Trim(Application.Range("B" & i).value) <> "" Then
                If Trim(Application.Range("A" & i).value) <> "" Then
                    Set cur_option = TestOptionFactory(CDbl(Application.Range("A" & i).value)).first
                    cur_option.name = Trim(Application.Range("B" & i).value)
                    cur_option.description = Trim(Application.Range("C" & i).value)
                    cur_option.save
                Else
                    Set new_option = New TestOption
                    test_option_unlock
                    new_option.name = Trim(Application.Range("B" & i).value)
                    new_option.description = Trim(Application.Range("C" & i).value)
                    Application.Range("A" & i).value = new_option.save
                    test_option_lock
                End If
                updated_count = updated_count + 1
            End If
        Next
        If updated_count Then
            MsgBox updated_count & " Test Objects updated succesfully!"
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_option_delete()
    On Error GoTo err_handler
    If MsgBox("Are you sure you want to delete record(s) permanently?", vbYesNo) = vbYes Then
        deleted_count = 0
        For i = 2 To get_last_row_column(0)
            If UCase(Trim(Application.Range("D" & i).value)) = "Y" Then
                test_option_unlock
                Set cur_option = TestOptionFactory(CDbl(Application.Range("A" & i).value)).first
                cur_option.delete
                deleted_count = deleted_count + 1
                rows(i).delete Shift:=xlUp
                i = i - 1
                test_option_lock
            End If
        Next
        If deleted_count Then
            MsgBox deleted_count & " Test Objects deleted succesfully!"
        End If
    End If
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub

Sub test_option_initial_state()
    test_option_unlock
    SheetClear
    Application.Range("A1").value = "ID"
    Application.Range("B1").value = "Name"
    Application.Range("C1").value = "Description"
    Application.Range("D1").value = "Update"
    Application.Columns("A:D").EntireColumn.AutoFit
    ApplyBorder "A1:D1"
    ApplyColor "A1:" & "D" & 1, colorBlack
    ApplyFontColor "A1:" & "D" & 1, colorWhite
    
    FreezePane 1
    test_option_lock
    Application.Range("A2").Select
End Sub
    
Public Sub test_option_show()
    On Error GoTo err_handler
   
    row_adder = 0
    test_option_clear_sheet
    sheet_size = get_last_row_column
    test_option_unlock
    For Each o In test_option_search_results
        Application.Range("A" & sheet_size(0) + 1 + row_adder).value = o.id
        Application.Range("B" & sheet_size(0) + 1 + row_adder).value = o.name
        Application.Range("C" & sheet_size(0) + 1 + row_adder).value = o.description
        row_adder = row_adder + 1
    Next
    sheet_size = get_last_row_column
    Application.Range("A2").Select
    ApplyBorder ("A" & 2 & ":D" & sheet_size(0))
    Application.Columns("A:D").EntireColumn.AutoFit
    ApplyAltColor 2, sheet_size(0), "A", "D"
    test_option_lock
    Application.ScreenUpdating = True
    Exit Sub
err_handler:
    MsgBox Err.description
End Sub


Public Sub test_option_lock()
    On Error Resume Next
    Application.ActiveSheet.Protection.AllowEditRanges.Add Title:="TestOptionsLock", Range _
    :=Columns("B:D")
    Application.ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
End Sub

Public Sub test_option_unlock()
    Application.ActiveSheet.Unprotect
End Sub


