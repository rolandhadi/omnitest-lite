'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "SqliteConnection"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class SqliteConnection
Public m_state
Public m_connection_string
Public m_last_output

Public Property Get state()
    state = m_state
End Property

Public Property Let state(value)
    m_state = value
End Property

Public Property Get last_output()
    last_output = m_last_output
End Property

Public Property Let last_output(value)
    m_last_output = value
End Property

Private Sub Class_Initialize()
    m_state = 0
End Sub

Public Function open_(connection_string)

    m_connection_string = connection_string
    
End Function


Public Function execute(sql, headers)
    sql_file = create_temp_text
    head = Replace(Join(headers, vbCrLf), "output.txt", SESSION_NAME)
    write_text sql_file, head & vbCrLf & sql & ";" & vbCrLf & ".exit", 2
    Set objShell = CreateObject("WScript.Shell")
    comspec = objShell.ExpandEnvironmentStrings("%comspec%")
    DontWaitUntilFinished = False
    ShowWindow = 1
    DontShowWindow = 0
    WaitUntilFinished = True
    cmd = comspec & " /c c:\sqlite3\exec.bat " & LCase(Left(DATABASE_PATH, 2)) & " """ & Replace(DATABASE_PATH, "database.sqlite", "") & " "" " & sql_file
    objShell.Run cmd, DontShowWindow, WaitUntilFinished
    Set new_sqlite_rs = New SqliteRecordset
    On Error Resume Next
    m_last_output = CreateObject("Scripting.FileSystemObject").openTextFile(Left(DATABASE_PATH, Len(DATABASE_PATH) - 15) & SESSION_NAME).readAll()
    Err.Clear: On Error GoTo 0
    new_sqlite_rs.record_raw = m_last_output
    On Error Resume Next
    Kill sql_file
    Kill Left(DATABASE_PATH, Len(DATABASE_PATH) - 15) & SESSION_NAME
    Err.Clear: On Error GoTo 0
    Set execute = new_sqlite_rs
End Function

Public Function execute_batch(sqls, headers)
    sql_cnt = 1
    sql_batch = ""
    For Each sql In sqls
        sql_file = create_temp_text
        new_session_name = SESSION_NAME & "_" & sql_cnt
        head = Replace(Join(headers, vbCrLf), "output.txt", new_session_name)
        sql_batch = sql_batch & head & vbCrLf & sql & ";" & vbCrLf
        sql_cnt = sql_cnt + 1
    Next
    write_text sql_file, sql_batch & ".exit", 2
    Set objShell = CreateObject("WScript.Shell")
    comspec = objShell.ExpandEnvironmentStrings("%comspec%")
    DontWaitUntilFinished = False
    ShowWindow = 1
    DontShowWindow = 0
    WaitUntilFinished = True
    cmd = comspec & " /c c:\sqlite3\exec.bat " & LCase(Left(DATABASE_PATH, 2)) & " """ & Replace(DATABASE_PATH, "database.sqlite", "") & " "" " & sql_file
    objShell.Run cmd, DontShowWindow, WaitUntilFinished
    On Error Resume Next
    sql_cnt = 1
    For Each sql In sqls
        new_session_name = SESSION_NAME & "_" & sql_cnt
        array_push new_sqlite_rs, New SqliteRecordset
        array_push m_last_output, CreateObject("Scripting.FileSystemObject").openTextFile(Left(DATABASE_PATH, Len(DATABASE_PATH) - 15) & new_session_name).readAll()
        Kill Left(DATABASE_PATH, Len(DATABASE_PATH) - 15) & new_session_name
        new_sqlite_rs(sql_cnt - 1).record_raw = m_last_output(sql_cnt - 1)
        sql_cnt = sql_cnt + 1
    Next
    Err.Clear: On Error GoTo 0
    On Error Resume Next
    Kill sql_file
    Err.Clear: On Error GoTo 0
    execute_batch = new_sqlite_rs
End Function

Public Function create_temp_text()
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set tfolder = fso.GetSpecialFolder(2)
   tname = fso.GetTempName
   Set tfile = tfolder.CreateTextFile(tname)
   create_temp_text = tfolder & "\" & tname
End Function

Public Function create_temp_filename()
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set tfolder = fso.GetSpecialFolder(2)
   tname = fso.GetTempName
   Set tfile = tfolder.CreateTextFile(tname)
   create_temp_filename = tname
End Function

Public Function write_text(filename, text, iomode)
    Set object_file = CreateObject("Scripting.FileSystemObject").openTextFile(filename, iomode, True)
    object_file.WriteLine text
    object_file.Close
    Set object_file = Nothing
End Function
End Class