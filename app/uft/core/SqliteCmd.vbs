''VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "SqliteCmd"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class SqliteCmd
Private connection
Private connection_string
Public last_id

Private Sub Class_Initialize()
    Set connection = New SqliteConnection
End Sub

Public Sub init()
    connection_string = CURRENT_CONNECTION
End Sub

Public Function get_records(sql)
    If connection.state = 0 Then
        connection.open_ connection_string
    End If
    Set rs = connection.execute(sql, Array( _
                                    ".headers on", _
                                    ".timeout 10000", _
                                    ".output " & "output.txt", _
                                    "PRAGMA foreign_keys = ON;" _
                                ))
    Set out = CreateObject("Scripting.Dictionary")
    out("columns") = rs.fields_count
    out("rows") = 0
    Set out("data") = CreateObject("Scripting.Dictionary")
    If IsArray(rs.fields) Then
        For Each field In rs.fields
                out("data")(field) = rs.getrows(field)
                out("rows") = UBound(out("data")(field), 2) + 1
        Next
    End If
    Set records = New Recordset
    records.init (out)
    Set get_records = records
End Function

Public Function get_batch_records(sqls)
    If connection.state = 0 Then
        connection.open_ connection_string
    End If
    rss = connection.execute_batch(sqls, Array( _
                                    ".headers on", _
                                    ".timeout 10000", _
                                    ".output " & "output.txt", _
                                    "PRAGMA foreign_keys = ON;" _
                                ))
    rs_cnt = 0
    For Each rs In rss
        array_push out, CreateObject("Scripting.Dictionary")
        If IsNullOrEmpty(rs.record_raw) Then
            out(rs_cnt)("columns") = 0
            out(rs_cnt)("rows") = 0
            Set out(rs_cnt)("data") = CreateObject("Scripting.Dictionary")
        Else
            out(rs_cnt)("columns") = rs.fields_count
            out(rs_cnt)("rows") = 0
            Set out(rs_cnt)("data") = CreateObject("Scripting.Dictionary")
            If IsArray(rs.fields) Then
                For Each field In rs.fields
                        out(rs_cnt)("data")(field) = rs.getrows(field)
                        out(rs_cnt)("rows") = UBound(out(rs_cnt)("data")(field), 2) + 1
                Next
            End If
        End If
        array_push records, New Recordset
        records(rs_cnt).init (out(rs_cnt))
        rs_cnt = rs_cnt + 1
    Next
    get_batch_records = records
End Function

Public Function execute(sql)
    connection.open_ (connection_string)
    If IsArray(sql) Then
        connection.execute Join(sql, ";"), Array( _
                                        ".timeout 10000", _
                                        ".output " & "output.txt", _
                                        "PRAGMA foreign_keys = ON;" _
                                    )
    Else
        If InStr(1, sql, "INSERT INTO") > 0 Then
            sql = sql & "; select last_insert_rowid();"
            connection.execute sql, Array( _
                                        ".timeout 10000", _
                                        ".output " & "output.txt", _
                                        "PRAGMA foreign_keys = ON;" _
                                    )
            last_id = CDbl(connection.last_output)
        Else
            connection.execute sql, Array( _
                                    ".headers on", _
                                    ".timeout 10000", _
                                    ".output " & "output.txt", _
                                    "PRAGMA foreign_keys = ON;" _
                                )
        End If
    End If
End Function
End Class