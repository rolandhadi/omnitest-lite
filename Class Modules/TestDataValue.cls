VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestDataValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_id
Public m_td_id
Public m_iteration
Public m_item
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get td_id()
    td_id = m_td_id
End Property

Public Property Let td_id(value)
    m_update("td_id") = value
    m_td_id = value
End Property

Public Property Get iteration()
    iteration = m_iteration
End Property

Public Property Let iteration(value)
    m_update("iteration") = "'" & value & "'"
    m_iteration = value
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

Public Sub init(id, td_id, iteration, item)
    m_id = id
    Me.td_id td_id
    Me.iteration iteration
    Me.item item
End Sub

Public Function test_data()
    Set test_data = TestDataFactory(Array("id", m_td_id))
End Function

Public Function save()
    new_id = db_save_record(m_id, m_update, "test_data_values")
    If new_id Then
        m_id = new_id
    Else
        new_id = ""
    End If
    save = new_id
End Function

Public Function delete()
    delete = db_delete_record(m_id, "test_data_values")
End Function

