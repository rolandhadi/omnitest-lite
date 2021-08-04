VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public connection

Private Sub Class_Initialize()
    get_connection
End Sub

Public Function get_connection()
    If DATABASE_TYPE = "ACCESS" Then
        CURRENT_CONNECTION = DATABASE_PROVIDER & "Data Source=" & DATABASE_PATH
        Set connection = New Access
    ElseIf DATABASE_TYPE = "SQLITE" Then
        CURRENT_CONNECTION = DATABASE_PATH
        Set connection = New SqliteCmd
    End If
    connection.init
End Function
