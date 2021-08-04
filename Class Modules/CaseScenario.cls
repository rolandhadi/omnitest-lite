VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_ts_id
Public m_tc_id
Public m_order_no
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
    m_update("ts_id") = "'" & value & "'"
    m_ts_id = value
End Property

Public Property Get tc_id()
    tc_id = m_tc_id
End Property

Public Property Let tc_id(value)
    m_update("tc_id") = "'" & value & "'"
    m_tc_id = value
End Property

Public Property Get order_no()
    order_no = m_order_no
End Property

Public Property Let order_no(value)
    m_update("order_no") = "'" & value & "'"
    m_order_no = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, ts_id, tc_id, order_no)
    m_id = id
    Me.ts_id ts_id
    Me.tc_id tc_id
    Me.order_no order_no
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "case_scenarios")
    If new_id Then
        m_id = new_id
    End If
End Function

Public Function delete()
    delete = db_delete_record(m_id, "case_scenarios")
End Function

Public Function test_case()
    Set test_case = TestCaseFactory(Array("id", m_tc_id))
End Function

Public Function test_scenario()
    Set test_scenario = TestScenarioFactory(Array("ts_id", m_ts_id))
End Function

