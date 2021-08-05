'VERSION 1.0 CLASS
'BEGIN
  'MultiUse = -1  'True
'END
'Attribute VB_Name = "TestData"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
Class TestData
Public m_id
Public m_parent_id
Public m_name
Public m_update

Public Property Get id()
    id = m_id
End Property

Public Property Let id(value)
    m_id = value
End Property

Public Property Get parent_id()
    parent_id = m_parent_id
End Property

Public Property Let parent_id(value)
    m_update("parent") = value
    m_parent_id = value
End Property

Public Property Get name()
    name = m_name
End Property

Public Property Let name(value)
    m_update("name") = "'" & db_qoutes(value) & "'"
    m_name = value
End Property

Private Sub Class_Initialize()

    Set m_update = CreateObject("Scripting.Dictionary")
				
End Sub

Public Sub init(id, parent_id, name)

    m_id = id
    Me.parent_id parent_id
    Me.name name
				
End Sub

Public Function parent()

    If IsEmpty(m_parent_id) Or IsNull(m_parent_id) Then
        pid = -1
    Else
        pid = m_parent_id
    End If
    Set parent = FolderPathFactory(Array("id", pid))
				
End Function

Public Function parent_or_new(name)

    Set find_parent = FolderPathFactory(Array("name", "'" & name & "'"))
    If find_parent.count = 0 Then
        Set new_parent = New FolderPath
        new_parent.name = name
        new_parent.module_code = 1
        parent_id = new_parent.save
        parent_or_new = m_parent_id
    Else
        parent_id = find_parent.first.id
        parent_or_new = m_parent_id
    End If
				
End Function

Public Function save()

    new_id = db_save_record(m_id, m_update, "test_datas")
    If new_id Then
        m_id = new_id
    End If
    save = new_id
				
End Function

Public Function delete()

    delete = db_delete_record(m_id, "test_datas")
				
End Function

Public Function values()

    Set values = TestDataValueFactory(Array("td_id", m_id))
				
End Function

Public Function iteration(iteration_number)

    Set cur_iteration = TestDataValueFactory(Array( _
																																																				Array("td_id", m_id), _
																																																				Array("iteration", iteration_number) _
																																												) _
																				)
															If cur_iteration.count <= 0 Then
																				Set cur_iteration = TestDataValueFactory(Array( _
																																																																				Array("td_id", m_id) _
																																																												) _
																																				)
                End If

                If cur_iteration.count > 0 Then
                    iteration = cur_iteration.first.item
                Else
                    iteration = ""
                End If

        Set data_input_matches = reg_match(iteration, "INPUT\[(\w+)\]")
        For i = 0 To data_input_matches.count - 1
          data_identifier_name = data_input_matches(i).SubMatches(0)
          Set data_identifier = TestDataFactory(Array("name", "'" & data_identifier_name & "'"))
          If data_identifier.count > 0 Then
              iteration = reg_replace(iteration, Replace(Replace(data_input_matches(i).value, "[", "\["), "]", "\]"), data_identifier.first.iteration(iteration_number))
          End If
        Next

        Set data_input_matches = reg_match(iteration, "VAR\[(\w+)\]")
        For i = 0 To data_input_matches.count - 1
          data_identifier_name = data_input_matches(i).SubMatches(0)
          iteration = reg_replace(iteration, Replace(Replace(data_input_matches(i).value, "[", "\["), "]", "\]"), Eval(data_identifier_name))
        Next
        
End Function

Public Function set_iteration(iteration_number, value)

    Set cur_iteration = TestDataValueFactory(Array( _
																																																Array("td_id", m_id), _
																																																Array("iteration", iteration_number) _
																																														) _
																																								)
				If cur_iteration.count <= 0 Then
					 Set new_iteration = New TestDataValue
						new_iteration.td_id = m_id
						new_iteration.iteration = iteration_number
						new_iteration.item = value
						new_iteration.save
				Else
						Set update_iteration = cur_iteration.first
						update_iteration.item = value
						update_iteration.save
				End If
				
End Function


End Class
