'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "FunctionReferenceValue"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class FunctionReferenceValue
Public m_id
Public m_fr_id
Public m_name
Public m_item
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get fr_id()
    fr_id = m_fr_id
End Property

Public Property Let fr_id(value)
    m_update("fr_id") = "'" & db_qoutes(value) & "'"
    m_fr_id = value
End Property

Public Property Get name()
    name = m_name
End Property

Public Property Let name(value)
    m_update("name") = "'" & db_qoutes(value) & "'"
    m_name = value
End Property

Public Property Get item()
    item = m_item
End Property

Public Property Let item(value)
    m_update("item") = "'" & db_qoutes(value) & "'"
    m_item = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, fr_id, name, item)
    m_id = id
    Me.fr_id fr_id
    Me.name name
    Me.item item
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "function_reference_values")
    If new_id Then
        m_id = new_id
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "function_reference_values")
End Function

Public Function function_reference()
    Set function_reference = FunctionReferenceFactory(Array("id", m_fr_id))
End Function

End Class