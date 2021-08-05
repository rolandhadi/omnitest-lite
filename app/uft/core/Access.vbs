'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "Access"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class Access
Private connection
Private connection_string
Public last_id

Private Sub Class_Initialize()
    Set connection = CreateObject("ADODB.Connection")
End Sub

Public Sub init()
    connection_string = CURRENT_CONNECTION
End Sub

Public Function get_records(sql)
    If connection.state = 0 Then
        connection.ConnectionTimeout = 100
        connection.Open (connection_string)
    End If
    Set rs = connection.execute(sql, , 1)
    Set out = CreateObject("Scripting.Dictionary")
    out("columns") = rs.fields.count
    out("rows") = 0
    Set out("data") = CreateObject("Scripting.Dictionary")
    For Each field In rs.fields
        If Not rs.eof Then
            out("data")(field.name) = rs.getrows(, , field.name)
            out("rows") = UBound(out("data")(field.name), 2) + 1
            rs.movefirst
        End If
    Next
    connection.Close
    Set records = New Recordset
    records.init (out)
    Set get_records = records
End Function

Public Function get_batch_records(sqls)
    If connection.state = 0 Then
        connection.Open connection_string
    End If
    For Each sql In sqls
        array_push rss, connection.execute(sql, , 1)
    Next
    rs_cnt = 0
    For Each rs In rss
        array_push out, CreateObject("Scripting.Dictionary")
        If rs.eof Then
            out(rs_cnt)("columns") = 0
            out(rs_cnt)("rows") = 0
            Set out(rs_cnt)("data") = CreateObject("Scripting.Dictionary")
        Else
            out(rs_cnt)("columns") = rs.fields.count
            out(rs_cnt)("rows") = 0
            Set out(rs_cnt)("data") = CreateObject("Scripting.Dictionary")
            For Each field In rs.fields
                If Not rs.eof Then
                    out(rs_cnt)("data")(field.name) = rs.getrows(, , field.name)
                    out(rs_cnt)("rows") = UBound(out(rs_cnt)("data")(field.name), 2) + 1
                    rs.movefirst
                End If
            Next
        End If
        array_push records, New Recordset
        records(rs_cnt).init (out(rs_cnt))
        rs_cnt = rs_cnt + 1
    Next
    get_batch_records = records
End Function

Public Function execute(sql)
    connection.Open (connection_string)
    If IsArray(sql) Then
        For Each s In sql
            connection.execute s, , 1
        Next
    Else
        connection.execute sql, , 1
    End If
    get_last_id
    connection.Close
End Function

Private Function get_last_id()
    last_id = connection.execute("SELECT @@Identity", , 1).fields.item(0).value
End Function
End Class