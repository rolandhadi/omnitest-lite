'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "FunctionReference"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class FunctionReference
Public m_id
Public m_name
Public m_type_code
Public m_struct
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

Public Property Get type_code()
    type_code = m_type_code
End Property

Public Property Let type_code(value)
    m_update("type_code") = "'" & db_qoutes(value) & "'"
    m_type_code = value
End Property

Public Property Get struct()
    struct = m_struct
End Property

Public Property Let struct(value)
    m_update("struct") = "'" & db_qoutes(value) & "'"
    m_struct = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, name, type_code, struct)
    m_id = id
    Me.name name
    Me.type_code type_code
    Me.struct struct
End Sub

Public Function save()
    new_id = db_save_record(m_id, m_update, "function_references")
    If new_id Then
        m_id = new_id
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "function_references")
End Function

Public Function values()
    Set values = FunctionReferenceValueFactory(Array("fr_id", m_id))
End Function
End Class