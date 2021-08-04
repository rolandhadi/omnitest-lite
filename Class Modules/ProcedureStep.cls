VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcedureStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_tp_id
Public m_to_id
Public m_keyword_name
Public m_order_no
Public m_description
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
    If IsNull(value) Then value = "NULL"
    m_update("tp_id") = value
    m_tp_id = value
End Property

Public Property Get to_id()
    to_id = m_to_id
End Property

Public Property Let to_id(value)
    If IsNull(value) Then value = "NULL"
    m_update("to_id") = value
    m_to_id = value
End Property

Public Property Get keyword_name()
    keyword_name = m_keyword_name
End Property

Public Property Let keyword_name(value)
    m_update("keyword_name") = "'" & value & "'"
    m_keyword_name = value
End Property

Public Property Get order_no()
    order_no = m_order_no
End Property

Public Property Let order_no(value)
    If IsNull(value) Then value = "NULL"
    m_update("order_no") = value
    m_order_no = value
End Property

Public Property Get description()
    description = m_description
End Property

Public Property Let description(value)
    m_update("description") = "'" & value & "'"
    m_description = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, tp_id, to_id, keyword_name, order_no, description)
    m_id = id
    Me.tp_id tp_id
    Me.to_id to_id
    Me.keyword_name keyword_name
    Me.order_no order_no
    Me.description description
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "procedure_steps")
    If new_id Then
        m_id = new_id
    End If
End Function

Public Function delete()
    delete = db_delete_record(m_id, "procedure_steps")
End Function

Public Function test_procedure()
    Set test_procedure = TestProcedureFactory(Array("tp_id", m_tp_id))
End Function

Public Function test_object()
    If IsNull(m_to_id) Then m_to_id = -1
    Set test_object = TestObjectFactory(Array("id", m_to_id))
End Function

Public Function test_option_links()
    Set test_option_links = TestOptionLinkFactory(Array("ps_id", m_id))
End Function

Public Function links()
    Set links = TestScenarioLinkFactory(Array( _
                                            Array("ps_id", m_id), _
                                            Array("tp_id", m_tp_id), _
                                            Array("ts_id", Null), _
                                            Array("tc_id", Null), _
                                            Array("cs_id", Null), _
                                            Array("cp_id", Null) _
                                            ))
End Function
