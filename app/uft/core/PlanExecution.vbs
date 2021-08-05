'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "PlanExecution"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class PlanExecution
Public m_id
Public m_name
Public m_next_action
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get name()
    name = m_name
End Property

Public Property Let name(value)
    m_update("name") = "'" & db_qoutes(value) & "'"
    m_name = value
End Property

Public Property Get next_action()
    next_action = m_next_action
End Property

Public Property Let next_action(value)
    m_update("next_action") = "'" & db_qoutes(value) & "'"
    m_next_action = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, name, next_action)
    m_id = id
    Me.name name
    Me.next_action next_action
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "plan_executions")
    If new_id Then
        m_id = new_id
    End If
    save = new_id
End Function

Public Function batch_save()
    batch_save = db_get_sql(m_id, m_update, "plan_executions")
End Function

Public Function delete()
    delete = db_delete_record(m_id, "plan_executions")
End Function

Public Function execution_scenarios()
    Set execution_scenarios = ExecutionScenarioFactory(Array("pe_id", m_id, Array("order_no", "asc")))
End Function



End Class