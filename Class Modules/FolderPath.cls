VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_module_code
Public m_name
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get module_code()
    module_code = m_module_code
End Property

Public Property Let module_code(value)
    m_update("module_code") = value
    m_module_code = value
End Property

Public Property Get name()
    name = m_name
End Property

Public Property Let name(value)
    m_update("name") = "'" & value & "'"
    m_name = value
End Property

Private Sub Class_Initialize()
    Set m_update = CreateObject("Scripting.Dictionary")
End Sub

Public Sub init(id, module_code, name)
    m_id = id
    Me.module_code module_code
    Me.name name
End Sub

Public Function find_or_new(name)
    find_or_new = db_find_or_new_record(name, "folder_paths")
End Function

Public Function save()
    save = db_save_record(m_id, m_update, "folder_paths")
End Function

Public Function delete()
    delete = db_delete_record(m_id, "folder_paths")
End Function
