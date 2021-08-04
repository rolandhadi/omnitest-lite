VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestScenarioLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_ts_id
Public m_tc_id
Public m_tp_id
Public m_cs_id
Public m_cp_id
Public m_ps_id
Public m_execute
Public m_screenshot
Public m_data_value_ref_id
Public m_data_value_id_in
Public m_data_value_in
Public m_data_value_id_out
Public m_data_value_out
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get ts_id()
    ts_id = m_ts_id
End Property

Public Property Let ts_id(value)
    m_update("ts_id") = value
    m_ts_id = value
End Property

Public Property Get tc_id()
    tc_id = m_tc_id
End Property

Public Property Let tc_id(value)
    m_update("tc_id") = value
    m_tc_id = value
End Property

Public Property Get tp_id()
    tp_id = m_tp_id
End Property

Public Property Let tp_id(value)
    m_update("tp_id") = value
    m_tp_id = value
End Property

Public Property Get cs_id()
    cs_id = m_cs_id
End Property

Public Property Let cs_id(value)
    m_update("cs_id") = value
    m_cs_id = value
End Property

Public Property Get cp_id()
    cp_id = m_cp_id
End Property

Public Property Let cp_id(value)
    m_update("cp_id") = value
    m_cp_id = value
End Property

Public Property Get ps_id()
    ps_id = m_ps_id
End Property

Public Property Let ps_id(value)
    m_update("ps_id") = value
    m_ps_id = value
End Property

Public Property Get execute()
    execute = m_execute
End Property

Public Property Let execute(value)
    m_update("execute") = value
    m_execute = value
End Property

Public Property Get screenshot()
    screenshot = m_screenshot
End Property

Public Property Let screenshot(value)
    m_update("screenshot") = value
    m_screenshot = value
End Property

Public Property Get data_value_ref_id()
    data_value_ref_id = m_data_value_ref_id
End Property

Public Property Let data_value_ref_id(value)
    m_update("data_value_ref_id") = value
    m_data_value_ref_id = value
End Property

Public Property Get data_value_id_in()
    data_value_id_in = m_data_value_id_in
End Property

Public Property Let data_value_id_in(value)
    m_update("data_value_id_in") = value
    m_data_value_id_in = value
End Property

Public Property Get data_value_in()
    data_value_in = m_data_value_in
End Property

Public Property Let data_value_in(value)
    m_update("data_value_in") = "'" & value & "'"
    m_data_value_in = value
End Property

Public Property Get data_value_id_out()
    data_value_id_out = m_data_value_id_out
End Property

Public Property Let data_value_id_out(value)
    m_update("data_value_id_out") = value
    m_data_value_id_out = value
End Property

Public Property Get data_value_out()
    data_value_out = m_data_value_out
End Property

Public Property Let data_value_out(value)
    m_update("data_value_out") = "'" & value & "'"
    m_data_value_out = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, ts_id, tc_id, tp_id, cs_id, cp_id, ps_id, execute, screenshot, data_value_ref_id, data_value_id_in, data_value_in, data_value_id_out, data_value_out)
    m_id = id
    Me.ts_id ts_id
    Me.tc_id tc_id
    Me.tp_id tp_id
    Me.cs_id cs_id
    Me.cp_id cp_id
    Me.ps_id ps_id
    Me.execute execute
    Me.screenshot screenshot
    Me.data_value_ref_id data_value_ref_id
    Me.data_value_id_in data_value_id_in
    Me.data_value_in data_value_in
    Me.data_value_id_out data_value_id_out
    Me.data_value_out data_value_out
End Sub

Public Function function_reference_in()
    Set function_reference_in = FunctionReferenceFactory(Array("id", m_data_value_ref_id))
End Function

Public Function test_data_in()
    Set test_data_in = TestDataFactory(Array("id", m_data_value_id_in))
End Function

Public Function test_data_out()
    Set test_data_out = TestDataFactory(Array("id", m_data_value_id_out))
End Function

Public Function test_procedure()
    Set test_procedure = TestProcedureFactory(Array("id", m_tp_id))
End Function

Public Function test_case()
    Set test_case = TestCaseFactory(Array("id", m_tc_id))
End Function

Public Function test_scenario()
    Set test_scenario = TestScenarioFactory(Array("id", m_ts_id))
End Function

Public Function procedure_step()
    Set procedure_step = ProcedureStepFactory(Array("id", m_ps_id))
End Function

Public Function case_procedure()
    Set case_procedure = CaseProcedureFactory(Array("id", m_cp_id))
End Function

Public Function case_scenario()
    Set case_scenario = CaseScenarioFactory(Array("id", m_cs_id))
End Function

Public Function save()
    new_id = db_save_record(m_id, m_update, "test_scenario_links")
    If new_id Then
        m_id = new_id
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "test_scenario_links")
End Function


