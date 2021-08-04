VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_tc_id
Public m_tp_id
Public m_order_no
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get tc_id()
    tc_id = m_tc_id
End Property

Public Property Let tc_id(value)
    m_update("tc_id") = "'" & value & "'"
    m_tc_id = value
End Property

Public Property Get tp_id()
    tp_id = m_tp_id
End Property

Public Property Let tp_id(value)
    m_update("tp_id") = "'" & value & "'"
    m_tp_id = value
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

Public Sub init(id, tc_id, tp_id, order_no)
    m_id = id
    Me.tc_id tc_id
    Me.tp_id tp_id
    Me.order_no order_no
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "case_procedures")
    If new_id Then
        m_id = new_id
    End If
End Function

Public Function delete()
    delete = db_delete_record(m_id, "case_procedures")
End Function

Public Function test_procedure()
    Set test_procedure = TestProcedureFactory(Array("id", m_tp_id))
End Function

Public Function test_case()
    Set test_case = TestCaseFactory(Array("tc_id", m_tc_id))
End Function
