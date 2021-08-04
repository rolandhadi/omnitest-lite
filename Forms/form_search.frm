VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_search 
   Caption         =   "Search"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12120
   OleObjectBlob   =   "form_search.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private factory
Private data_selected
Private module
Private multi_search

Private Sub button_edit_Click()
    Select Case module
        Case TEST_OBJECT_TAB
            test_object_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist test_object_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case TEST_OPTION_TAB
            test_option_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist test_option_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case TEST_DATA_TAB
            test_data_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist test_data_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case TEST_PROCEDURE_TAB
            test_procedure_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist test_procedure_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case TEST_CASE_TAB
            test_case_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist test_case_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case TEST_SCENARIO_TAB
            test_scenario_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist test_scenario_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case TEST_SCENARIO_DATA_TAB
            test_scenario_data_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist test_scenario_data_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case PLAN_EXECUTION_TAB
            plan_execution_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist plan_execution_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
        Case FUNCTION_REFERENCE_TAB
            function_reference_search_results = Null
            For i = 0 To list_search.ListCount - 1
                If list_search.List(i, 0) = LIST_CHECKED Then
                    array_push_not_exist function_reference_search_results, factory.find(list_search.List(i, 1))
                End If
            Next
    End Select
    data_selected = True
    Unload Me
End Sub

Private Sub button_refresh_Click()
    text_search.text = ""
    checkbox_select_all.value = False
End Sub

Private Sub checkbox_select_all_Click()
    If checkbox_select_all.value = True Then
        For i = 0 To list_search.ListCount - 1
            list_search.List(i, 0) = LIST_CHECKED
        Next
    Else
        For i = 0 To list_search.ListCount - 1
            list_search.List(i, 0) = LIST_UNCHECKED
        Next
    End If
End Sub

Private Sub list_search_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    select_item list_search.ListIndex
End Sub

Private Sub list_search_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc(" ") Then
        select_item list_search.ListIndex
    End If
End Sub

Public Sub select_item(list_index)
    If list_index <> -1 Then
        If list_search.List(list_index, 0) = LIST_UNCHECKED Then
            If multi_search = False Then
                For i = 0 To list_search.ListCount - 1
                    list_search.List(i, 0) = LIST_UNCHECKED
                Next
            End If
            list_search.List(list_index, 0) = LIST_CHECKED
        Else
            list_search.List(list_index, 0) = LIST_UNCHECKED
        End If
    End If
End Sub

Public Sub init(module_, multi_search_)
    module = module_
    multi_search = multi_search_
    If multi_search = True Then
        checkbox_select_all.Visible = True
    Else
        checkbox_select_all.Visible = False
    End If
    list_search.Clear
End Sub

Private Sub text_search_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        search_string = Trim(text_search.text)
        If Trim(search_string) = "" Then Exit Sub
        If Trim(search_string) = "*" Then search_string = ""
        list_search.Clear
        Select Case module
            Case TEST_OBJECT_TAB
                If Left(search_string, 1) = "[" And Right(search_string, 1) = "]" Then
                    cur_folder = "'%" & Replace(Replace(search_string, "]", ""), "[", "") & "%'"
                    Set cur_parent = FolderPathFactory(Array("name", "LIKE", cur_folder))
                    If cur_parent.count > 0 Then
                        p = Array("parent", cur_parent.first.id)
                    Else
                        p = Array("parent", -1)
                    End If
                Else
                    p = Array("name", "LIKE", "'%" & search_string & "%'")
                End If
                Set factory = TestObjectFactory(p)
                list_search.ColumnCount = 4
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        Set parent_folder = item.parent.first
                        If Not parent_folder Is Nothing Then
                            list_search.List(list_search.ListCount - 1, 2) = parent_folder.name
                        End If
                        list_search.List(list_search.ListCount - 1, 3) = item.name
                    Next
                End If
            Case TEST_OPTION_TAB
                p = Array("name", "LIKE", "'%" & search_string & "%'")
                list_search.ColumnCount = 4
                Set factory = TestOptionFactory(p)
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        list_search.List(list_search.ListCount - 1, 2) = item.name
                        list_search.List(list_search.ListCount - 1, 3) = item.description
                    Next
                End If
            Case TEST_DATA_TAB
                list_search.ColumnCount = 4
                If Left(search_string, 1) = "[" And Right(search_string, 1) = "]" Then
                    cur_folder = "'%" & Replace(Replace(search_string, "]", ""), "[", "") & "%'"
                    Set cur_parent = FolderPathFactory(Array("name", "LIKE", cur_folder))
                    If cur_parent.count > 0 Then
                        p = Array("parent", cur_parent.first.id)
                    Else
                        p = Array("parent", -1)
                    End If
                Else
                    p = Array("name", "LIKE", "'%" & search_string & "%'")
                End If
                Set factory = TestDataFactory(p)
                list_search.ColumnCount = 4
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        Set parent_folder = item.parent.first
                        If Not parent_folder Is Nothing Then
                            list_search.List(list_search.ListCount - 1, 2) = parent_folder.name
                        End If
                        list_search.List(list_search.ListCount - 1, 3) = item.name
                    Next
                End If
            Case TEST_PROCEDURE_TAB
                p = Array("name", "LIKE", "'%" & search_string & "%'")
                Set factory = TestProcedureFactory(p)
                list_search.ColumnCount = 3
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        list_search.List(list_search.ListCount - 1, 2) = item.name
                    Next
                End If
            Case TEST_CASE_TAB
                p = Array("name", "LIKE", "'%" & search_string & "%'")
                Set factory = TestCaseFactory(p)
                list_search.ColumnCount = 3
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        list_search.List(list_search.ListCount - 1, 2) = item.name
                    Next
                End If
            Case TEST_SCENARIO_TAB, TEST_SCENARIO_DATA_TAB
                p = Array("name", "LIKE", "'%" & search_string & "%'")
                Set factory = TestScenarioFactory(p)
                list_search.ColumnCount = 3
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        list_search.List(list_search.ListCount - 1, 2) = item.name
                    Next
                End If
            Case PLAN_EXECUTION_TAB
                p = Array("name", "LIKE", "'%" & search_string & "%'")
                Set factory = PlanExecutionFactory(p)
                list_search.ColumnCount = 3
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        list_search.List(list_search.ListCount - 1, 2) = item.name
                    Next
                End If
            Case FUNCTION_REFERENCE_TAB
                p = Array("name", "LIKE", "'%" & search_string & "%'")
                Set factory = FunctionReferenceFactory(p)
                list_search.ColumnCount = 3
                If factory.count > 0 Then
                    For Each item In factory.fetch
                        list_search.AddItem
                        list_search.List(list_search.ListCount - 1, 0) = LIST_UNCHECKED
                        list_search.List(list_search.ListCount - 1, 1) = item.id
                        list_search.List(list_search.ListCount - 1, 2) = item.name
                    Next
                End If
            End Select
    End If
End Sub

Private Sub UserForm_Activate()
    text_search.SetFocus
End Sub

Private Sub UserForm_Terminate()
    Select Case module
        Case TEST_OBJECT_TAB
            If data_selected <> True Then test_object_search_results = Null
        Case TEST_OPTION_TAB
            If data_selected <> True Then test_option_search_results = Null
        Case TEST_DATA_TAB
            If data_selected <> True Then test_data_search_results = Null
        Case TEST_PROCEDURE_TAB
            If data_selected <> True Then test_procedure_search_results = Null
        Case TEST_CASE_TAB
            If data_selected <> True Then test_case_search_results = Null
        Case TEST_SCENARIO_TAB
            If data_selected <> True Then test_scenario_search_results = Null
        Case TEST_SCENARIO_DATA_TAB
            If data_selected <> True Then test_scenario_data_search_results = Null
        Case PLAN_EXECUTION_TAB
            If data_selected <> True Then plan_execution_search_results = Null
        Case FUNCTION_REFERENCE_TAB
            If data_selected <> True Then function_reference_search_results = Null
    End Select
End Sub
