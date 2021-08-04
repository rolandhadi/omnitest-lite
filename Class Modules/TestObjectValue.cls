VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestObjectValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_to_id
Public m_type_code
Public m_item
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get to_id()
    to_id = m_to_id
End Property

Public Property Let to_id(value)
    m_update("to_id") = value
    m_to_id = value
End Property

Public Property Get type_code()
    type_code = m_type_code
End Property

Public Property Let type_code(value)
    m_update("type_code") = "'" & value & "'"
    m_type_code = value
End Property

Public Property Get item()
    item = m_item
End Property

Public Property Let item(item_)
    m_update("item") = "'" & item_ & "'"
    m_item = item_
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, to_id, type_code, item)
    m_id = id
    Me.to_id to_id
    Me.type_code type_code
    Me.item item
End Sub

Public Function test_object()
    Set test_object = TestObjectFactory(Array("id", m_to_id))
End Function

Public Function save()
    new_id = db_save_record(m_id, m_update, "test_object_values")
    If new_id Then
        m_id = new_id
    Else
        new_id = ""
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "test_object_values")
End Function
