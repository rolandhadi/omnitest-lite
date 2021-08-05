'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "TestOption"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class TestOption
Public m_id
Public m_name
Public m_description
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

Public Property Get description()
    description = m_description
End Property

Public Property Let description(value)
    m_update("description") = "'" & db_qoutes(value) & "'"
    m_description = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, name, description)
    m_id = id
    Me.name name
    Me.description description
End Sub

Public Function parent()
    If IsEmpty(m_description) Or IsNull(m_description) Then
        pid = -1
    Else
        pid = m_description
    End If
    Set parent = FolderPathFactory(Array("id", pid))
End Function

Public Function save()
    new_id = db_save_record(m_id, m_update, "test_options")
    If new_id Then
        m_id = new_id
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "test_options")
End Function

Public Function links()
    Set links = TestOptionLinkFactory(Array("option_id", m_id))
End Function


End Class