VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExecutionScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_pe_id
Public m_ts_id
Public m_order_no
Public m_iteration
Public m_execute
Public m_dependency
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get pe_id()
    pe_id = m_pe_id
End Property

Public Property Let pe_id(value)
    m_update("pe_id") = "'" & value & "'"
    m_pe_id = value
End Property

Public Property Get ts_id()
    ts_id = m_ts_id
End Property

Public Property Let ts_id(value)
    m_update("ts_id") = "'" & value & "'"
    m_ts_id = value
End Property

Public Property Get order_no()
    order_no = m_order_no
End Property

Public Property Let order_no(value)
    m_update("order_no") = "'" & value & "'"
    m_order_no = value
End Property

Public Property Get iteration()
    iteration = m_iteration
End Property

Public Property Let iteration(value)
    m_update("iteration") = value
    m_iteration = value
End Property

Public Property Get execute()
    execute = m_execute
End Property

Public Property Let execute(value)
    m_update("execute") = "'" & value & "'"
    m_execute = value
End Property

Public Property Get dependency()
    dependency = m_dependency
End Property

Public Property Let dependency(value)
    m_update("dependency") = "'" & value & "'"
    m_dependency = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, pe_id, ts_id, order_no, iteration, execute, dependency)
    m_id = id
    Me.pe_id pe_id
    Me.ts_id ts_id
    Me.order_no order_no
    Me.iteration iteration
    Me.execute execute
    Me.dependency dependency
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "execution_scenarios")
    If new_id Then
        m_id = new_id
    End If
End Function

Public Function delete()
    delete = db_delete_record(m_id, "execution_scenarios")
End Function

Public Function test_case()
    Set test_case = TestCaseFactory(Array("id", m_tc_id))
End Function

Public Function test_scenario()
    Set test_scenario = TestScenarioFactory(Array("id", m_ts_id))
End Function

Public Function plan_execution()
    Set plan_execution = PlanExecutionFactory(Array("id", m_pe_id))
End Function

