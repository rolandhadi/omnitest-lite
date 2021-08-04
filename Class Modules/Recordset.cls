VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Recordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_records
Public m_cursor
Public m_eof

Public Property Get eof()
    eof = m_eof
End Property

Private Sub Class_Initialize()
    m_cursor = 0
    m_eof = True
End Sub

Public Function init(records)
    Set m_records = records
    If m_records("rows") > 0 Then
        m_eof = False
    Else
        m_eof = True
    End If
End Function

Public Function Recordset(column_name)
    Set Recordset = data(column_name)
End Function

Public Function loaded()
    If m_records("rows") > 0 Then
        loaded = True
    Else
        loaded = False
    End If
End Function

Public Function count()
    count = m_records("rows")
End Function

Public Function first()
    data = m_records("data")(column_name)(0, 0)
End Function

Public Function fetch()
    
End Function

Public Function data(column_name)
    If Not m_eof Then
        data = m_records("data")(column_name)(0, m_cursor)
    End If
End Function

Public Function move_next()
    If m_cursor < m_records("rows") Then
        m_cursor = m_cursor + 1
        If m_cursor >= m_records("rows") Then
            m_eof = True
        Else
            m_eof = False
        End If
    Else
        m_eof = True
    End If
End Function

Public Function move_first()
    m_cursor = 0
End Function

