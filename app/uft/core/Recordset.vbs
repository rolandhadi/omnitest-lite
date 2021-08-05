'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "Recordset"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class Recordset
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
    first = m_records("data")(column_name)(0, 0)
End Function

Public Function fetch()

End Function

Public Function data(column_name)
    If Not m_eof Then
        data = m_records("data")(column_name)(0, m_cursor)
    End If
End Function

Public Function find(column_name, exp, first_value)
    move_first
    out = ""
    If IsArray(column_name) Then
      If UBound(column_name) = 1 Then
        out = Array("", "")
      ElseIf UBound(column_name) = 2 Then
        out = Array("", "", "")
      ElseIf UBound(column_name) = 3 Then
        out = Array("", "", "", "")
      ElseIf UBound(column_name) = 4 Then
        out = Array("", "", "", "", "")
      ElseIf UBound(column_name) = 5 Then
        out = Array("", "", "", "", "", "")
      ElseIf UBound(column_name) = 6 Then
        out = Array("", "", "", "", "", "", "")
      End If
    End If
    If Not IsArray(exp(0)) Then
        While Not m_eof
            If CStr(m_records("data")(exp(0))(0, m_cursor)) = CStr(exp(1)) Then
                If IsArray(column_name) Then
                    If UBound(column_name) = 1 Then
                      array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                      array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                    ElseIf UBound(column_name) = 2 Then
                      array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                      array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                      array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                    ElseIf UBound(column_name) = 3 Then
                      array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                      array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                      array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                      array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                    ElseIf UBound(column_name) = 4 Then
                      array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                      array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                      array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                      array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                      array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                    ElseIf UBound(column_name) = 5 Then
                      array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                      array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                      array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                      array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                      array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                      array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                    ElseIf UBound(column_name) = 6 Then
                      array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                      array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                      array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                      array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                      array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                      array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                      array_push out(6), m_records("data")(column_name(6))(0, m_cursor)
                    End If
                Else
                    array_push out, m_records("data")(column_name)(0, m_cursor)
                End If
                If first_value Then
                    find = out
                    Exit Function
                End If
            End If
            move_next
        Wend
    Else
        While Not m_eof
            If UBound(exp) = 1 Then
                If CStr(m_records("data")(exp(0)(0))(0, m_cursor)) = CStr(exp(0)(1)) And _
                   CStr(m_records("data")(exp(1)(0))(0, m_cursor)) = CStr(exp(1)(1)) Then
                   If IsArray(column_name) Then
                     If UBound(column_name) = 1 Then
                       array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                       array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                     ElseIf UBound(column_name) = 2 Then
                       array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                       array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                       array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                     ElseIf UBound(column_name) = 3 Then
                       array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                       array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                       array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                       array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                     ElseIf UBound(column_name) = 4 Then
                       array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                       array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                       array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                       array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                       array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                     ElseIf UBound(column_name) = 5 Then
                       array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                       array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                       array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                       array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                       array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                       array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                     ElseIf UBound(column_name) = 6 Then
                       array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                       array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                       array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                       array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                       array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                       array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                       array_push out(6), m_records("data")(column_name(6))(0, m_cursor)
                     End If
                    Else
                        array_push out, m_records("data")(column_name)(0, m_cursor)
                    End If
                    If first_value Then
                        find = out
                        Exit Function
                    End If
                End If
            ElseIf UBound(exp) = 2 Then
                If CStr(m_records("data")(exp(0)(0))(0, m_cursor)) = CStr(exp(0)(1)) And _
                   CStr(m_records("data")(exp(1)(0))(0, m_cursor)) = CStr(exp(1)(1)) And _
                   CStr(m_records("data")(exp(2)(0))(0, m_cursor)) = CStr(exp(2)(1)) Then
                    If IsArray(column_name) Then
                      If UBound(column_name) = 1 Then
                        array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                        array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                      ElseIf UBound(column_name) = 2 Then
                        array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                        array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                        array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                      ElseIf UBound(column_name) = 3 Then
                        array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                        array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                        array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                        array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                      ElseIf UBound(column_name) = 4 Then
                        array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                        array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                        array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                        array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                        array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                      ElseIf UBound(column_name) = 5 Then
                        array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                        array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                        array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                        array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                        array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                        array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                      ElseIf UBound(column_name) = 6 Then
                        array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                        array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                        array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                        array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                        array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                        array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                        array_push out(6), m_records("data")(column_name(6))(0, m_cursor)
                      End If
                    Else
                        array_push out, m_records("data")(column_name)(0, m_cursor)
                    End If
                    If first_value Then
                        find = out
                        Exit Function
                    End If
                End If
            ElseIf UBound(exp) = 3 Then
                If CStr(m_records("data")(exp(0)(0))(0, m_cursor)) = CStr(exp(0)(1)) And _
                   CStr(m_records("data")(exp(1)(0))(0, m_cursor)) = CStr(exp(1)(1)) And _
                   CStr(m_records("data")(exp(2)(0))(0, m_cursor)) = CStr(exp(2)(1)) And _
                   CStr(m_records("data")(exp(3)(0))(0, m_cursor)) = CStr(exp(3)(1)) Then
                      If IsArray(column_name) Then
                        If UBound(column_name) = 1 Then
                          array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                          array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                        ElseIf UBound(column_name) = 2 Then
                          array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                          array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                          array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                        ElseIf UBound(column_name) = 3 Then
                          array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                          array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                          array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                          array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                        ElseIf UBound(column_name) = 4 Then
                          array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                          array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                          array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                          array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                          array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                        ElseIf UBound(column_name) = 5 Then
                          array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                          array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                          array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                          array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                          array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                          array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                        ElseIf UBound(column_name) = 6 Then
                          array_push out(0), m_records("data")(column_name(0))(0, m_cursor)
                          array_push out(1), m_records("data")(column_name(1))(0, m_cursor)
                          array_push out(2), m_records("data")(column_name(2))(0, m_cursor)
                          array_push out(3), m_records("data")(column_name(3))(0, m_cursor)
                          array_push out(4), m_records("data")(column_name(4))(0, m_cursor)
                          array_push out(5), m_records("data")(column_name(5))(0, m_cursor)
                          array_push out(6), m_records("data")(column_name(6))(0, m_cursor)
                        End If
                      Else
                          array_push out, m_records("data")(column_name)(0, m_cursor)
                      End If
                      If first_value Then
                          find = out
                          Exit Function
                      End If
                End If
            End If
            move_next
        Wend
    End If
    If IsArray(out) Then
        find = out
    Else
        find = Array("")
    End If
End Function

Public Function column(column_name, exp)
    move_first
    out = ""
    If IsArray(column_name) Then
        out = Array("", "")
    End If
    If Not IsArray(exp(0)) Then
        While Not m_eof
            If CStr(m_records("data")(exp(0))(0, m_cursor)) = CStr(exp(1)) Or exp(1) = "*" Then
                array_push_not_exist out, m_records("data")(column_name)(0, m_cursor)
            End If
            move_next
        Wend
    Else
        While Not m_eof
            If UBound(exp) = 1 Then
                If CStr(m_records("data")(exp(0)(0))(0, m_cursor)) = CStr(exp(0)(1)) And _
                    CStr(m_records("data")(exp(1)(0))(0, m_cursor)) = CStr(exp(1)(1)) Then
                    If IsArray(column_name) Then
                      array_push_not_exist out(0), m_records("data")(column_name(0))(0, m_cursor)
                      array_push_not_exist out(1), m_records("data")(column_name(1))(0, m_cursor)
                    Else
                      array_push_not_exist out, m_records("data")(column_name)(0, m_cursor)
                    End If
                End If
            ElseIf UBound(exp) = 2 Then
                If CStr(m_records("data")(exp(0)(0))(0, m_cursor)) = CStr(exp(0)(1)) And _
                   CStr(m_records("data")(exp(1)(0))(0, m_cursor)) = CStr(exp(1)(1)) And _
                   CStr(m_records("data")(exp(2)(0))(0, m_cursor)) = CStr(exp(2)(1)) Then
                   If IsArray(column_name) Then
                     array_push_not_exist out(0), m_records("data")(column_name(0))(0, m_cursor)
                     array_push_not_exist out(1), m_records("data")(column_name(1))(0, m_cursor)
                   Else
                     array_push_not_exist out, m_records("data")(column_name)(0, m_cursor)
                   End If
                End If
            ElseIf UBound(exp) = 3 Then
                If CStr(m_records("data")(exp(0)(0))(0, m_cursor)) = CStr(exp(0)(1)) And _
                   CStr(m_records("data")(exp(1)(0))(0, m_cursor)) = CStr(exp(1)(1)) And _
                   CStr(m_records("data")(exp(2)(0))(0, m_cursor)) = CStr(exp(2)(1)) And _
                   CStr(m_records("data")(exp(3)(0))(0, m_cursor)) = CStr(exp(3)(1)) Then
                   If IsArray(column_name) Then
                     array_push_not_exist out(0), m_records("data")(column_name(0))(0, m_cursor)
                     array_push_not_exist out(1), m_records("data")(column_name(1))(0, m_cursor)
                   Else
                     array_push_not_exist out, m_records("data")(column_name)(0, m_cursor)
                   End If
                End If
            End If
            move_next
        Wend
    End If
    If IsArray(out) Then
        column = out
    Else
        column = Array("")
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
    m_eof = False
    If m_records("rows") > 0 Then
        m_eof = False
    Else
        m_eof = True
    End If
End Function

End Class
