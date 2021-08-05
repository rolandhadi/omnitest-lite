'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "TestOptionLink"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class TestOptionLink
Public m_id
Public m_tp_id
Public m_ps_id
Public m_triggered
Public m_option_id
Public m_item
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get tp_id()
    tp_id = m_tp_id
End Property

Public Property Let tp_id(value)
    m_update("tp_id") = value
    m_tp_id = value
End Property

Public Property Get ps_id()
    ps_id = m_ps_id
End Property

Public Property Let ps_id(value)
    m_update("ps_id") = value
    m_ps_id = value
End Property

Public Property Get triggered()
    triggered = m_triggered
End Property

Public Property Let triggered(value)
    m_update("triggered") = "'" & db_qoutes(value) & "'"
    m_type = value
End Property

Public Property Get option_id()
    option_id = m_option_id
End Property

Public Property Let option_id(value)
    m_update("option_id") = "'" & db_qoutes(value) & "'"
    m_option_id = value
End Property

Public Property Get item()
    item = m_item
End Property

Public Property Let item(item_)
    m_update("item") = "'" & db_qoutes(item_) & "'"
    m_item = item_
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, tp_id, ps_id, triggered, option_id, item)
    m_id = id
    Me.tp_id tp_id
    Me.ps_id ps_id
    Me.triggered triggered
    Me.option_id option_id
    Me.item item
End Sub

Public Function test_option()
    Set test_option = TestOptionFactory(Array("id", m_option_id))
End Function

Public Function procedure_steps()
    Set procedure_steps = ProcedureStepFactory(Array("id", m_ps_id))
End Function

Public Function case_procedures()
    Set case_procedures = CaseProcedureFactory(Array("id", m_cp_id))
End Function

Public Function case_scenarios()
    Set case_scenarios = CaseScenarioFactory(Array("id", m_cs_id))
End Function

Public Function save()
    new_id = db_save_record(m_id, m_update, "test_option_links")
    If new_id Then
        m_id = new_id
    Else
        new_id = ""
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "test_option_links")
End Function

End Class