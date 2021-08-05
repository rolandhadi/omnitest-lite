'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "SqliteRecordset"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class SqliteRecordset
Public m_first_item_value
Public m_record_raw
Public m_records

Public Property Get record_raw()
    record_raw = m_record_raw
End Property

Public Property Let record_raw(value)
    m_record_raw = value
    Set out = CreateObject("Scripting.Dictionary")
    Set out("fields") = CreateObject("Scripting.Dictionary")
    Set out("item") = CreateObject("Scripting.Dictionary")
    row_counter = 1
    For Each row In Split(m_record_raw, Chr(10))
        If Trim(row) <> "" Then
            col_counter = 1
            For Each col In Split(row, "|")
                If row_counter = 1 Then
                    If Trim(col) <> "" Then
                        out("fields")(col_counter) = col
                        Set out("item")(out("fields")(col_counter)) = CreateObject("Scripting.Dictionary")
                    End If
                Else
                    If Trim(out("fields")(col_counter)) <> "" Then
                        out("item")(out("fields")(col_counter))(row_counter - 1) = db_restore_delimeter(col)
                    End If
                End If
                col_counter = col_counter + 1
            Next
            row_counter = row_counter + 1
        End If
    Next
    Set m_records = out
End Property

Public Property Get records()
    records = m_records
End Property

Public Property Let records(value)
    m_records = value
End Property

Public Property Get first_item_value()
    first_item_value = m_first_item_value
End Property

Public Property Let first_item_value(value)
    m_first_item_value = value
End Property

Public Function init(records)
    Me.records = records
End Function

Public Function getrows(field_name)
    Dim out()
    If Trim(field_name) <> "" Then
    ReDim out(0, m_records("item")(field_name).count - 1)
        cnt = 0
        For Each i In m_records("item")(field_name)
            out(0, cnt) = m_records("item")(field_name)(i)
            cnt = cnt + 1
        Next
    Else
        ReDim out(0, 0)
    End If
    getrows = out
End Function

Public Function fields_count()
    fields_count = m_records("fields").count
End Function

Public Function fields()
    out = ""
    For Each f In m_records("fields")
        array_push out, m_records("fields")(f)
    Next
    fields = out
End Function


End Class