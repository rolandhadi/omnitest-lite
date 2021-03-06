'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "TestProcedure"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class TestProcedure
Public m_id
Public m_name
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

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, name)
    m_id = id
    Me.name name
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "test_procedures")
    If new_id Then
        m_id = new_id
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "test_procedures")
End Function

Public Function procedure_steps()
    Set procedure_steps = ProcedureStepFactory(Array("tp_id", m_id, Array("order_no", "asc")))
End Function
End Class