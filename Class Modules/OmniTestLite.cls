VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OmniTestLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Initialize()
    Application.DisplayAlerts = False
    For Each sheet_ In ActiveWorkbook.Sheets
        If sheet_.name <> "Read_Me" Then
            sheet_.delete
        End If
    Next
    ActiveWorkbook.Sheets.Add.name = TEST_OPTION_TAB
    test_option_initial_state
    ActiveWorkbook.Sheets.Add.name = TEST_OBJECT_TAB
    test_object_initial_state
    ActiveWorkbook.Sheets.Add.name = TEST_DATA_TAB
    test_data_initial_state
    ActiveWorkbook.Sheets.Add.name = FUNCTION_REFERENCE_TAB
    function_reference_initial_state
    ActiveWorkbook.Sheets.Add.name = TEST_PROCEDURE_TAB
    test_procedure_initial_state
    ActiveWorkbook.Sheets.Add.name = TEST_CASE_TAB
    test_case_initial_state
    ActiveWorkbook.Sheets.Add.name = TEST_SCENARIO_TAB
    test_scenario_initial_state
    ActiveWorkbook.Sheets.Add.name = TEST_SCENARIO_DATA_TAB
    test_scenario_data_initial_state
    ActiveWorkbook.Sheets.Add.name = PLAN_EXECUTION_TAB
    plan_execution_initial_state
    Application.DisplayAlerts = True
    test_objects = Null
End Sub
