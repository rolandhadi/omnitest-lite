'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "TestFactory"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class TestFactory
Public content

Public Sub init(content)
    Set Me.content = content
End Sub

Public Function count()
    count = content("count")
End Function

Public Function keys()
    keys = content("data").keys
End Function

Public Function data(key)
    Set data = content("data")(key)
End Function

Public Function first()
    If content("count") > 0 Then
        Set first = content("data")(content("data").keys()(0))
    Else
        Set first = Nothing
    End If
End Function

Public Function fetch()
    If content("count") > 0 Then
        For Each item In content("data")
            array_push out, content("data")(item)
        Next
        fetch = out
    Else
        Set fetch = Nothing
    End If
End Function

Public Function find(exp)
    If IsNumeric(exp) Then
        If IsObject(content("data")((exp))) Then
            Set find = content("data")((exp))
        Else
            Set find = content("data")(CDbl(exp))
        End If
    Else
        For Each v In content("data")
            If Trim(UCase(content("data")(v).name)) = Trim(UCase(exp)) Then
                Set find = content("data")(v)
                Exit Function
            End If
        Next
    End If
End Function
End Class