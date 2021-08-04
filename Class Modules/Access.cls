VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Access"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Function execute(sql)
    connection.Open (connection_string)
    connection.execute sql, , 1
    get_last_id
    connection.Close
End Function

Private Function get_last_id()
    last_id = connection.execute("SELECT @@Identity", , 1).fields.item(0).value
End Function
