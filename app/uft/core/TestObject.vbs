'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "TestObject"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class TestObject
Public m_id
Public m_parent_id
Public m_name
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get parent_id()
    parent_id = m_parent_id
End Property

Public Property Let parent_id(value)
    m_update("parent") = value
    m_parent_id = value
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

Public Sub init(id, parent_id, name)
    m_id = id
    Me.parent_id parent_id
    Me.name name
End Sub

Public Function parent()
    If IsEmpty(m_parent_id) Or IsNull(m_parent_id) Then
        pid = -1
    Else
        pid = m_parent_id
    End If
    Set parent = FolderPathFactory(Array("id", pid))
End Function

Public Function parent_or_new(name)
    Set find_parent = FolderPathFactory(Array("name", "'" & name & "'"))
    If find_parent.count = 0 Then
        Set new_parent = New FolderPath
        new_parent.name = name
        new_parent.module_code = 1
        parent_id = new_parent.save
        parent_or_new = m_parent_id
    Else
        parent_id = find_parent.first.id
        parent_or_new = m_parent_id
    End If
End Function

Public Function save()
    new_id = db_save_record(m_id, m_update, "test_objects")
    If new_id Then
        m_id = new_id
    Else
        new_id = ""
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "test_objects")
End Function

Public Function values()
    Set values = TestObjectValueFactory(Array("to_id", m_id))
End Function

Public Function test_object_value(type_code)
    If type_code = "" Then type_code = "WEB"
    If IsNull(m_to_id) Then
        test_object_value = ""
    Else
        Set value = TestObjectValueFactory(Array( _
                                                    Array("to_id", m_id), _
                                                    Array("type_code", "'" & type_code & "'") _
                                                ) _
                                            )
        If value.count > 0 Then
            out = value.first.item
            Set matches = reg_match(out, "OBJ\[(\w+)\]")
            If matches.count > 0 Then
                For Each matching In matches
                    If InStr(1, out, matching) <= 0 Then Exit For
                    Set o = TestObjectFactory(Array("name", "'" & matching.SubMatches.item(0) & "'"))
                    If o.count > 0 Then
                        out = Replace(out, matching, o.first.values.first.item)
                    End If
                Next
                test_object_value = out
            Else
                test_object_value = value.first.item
            End If
        Else
            test_object_value = ""
        End If
    End If
End Function

End Class