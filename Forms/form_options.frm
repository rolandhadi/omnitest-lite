VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_options 
   Caption         =   "Test Options"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12120
   OleObjectBlob   =   "form_options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private factory
Private data_selected
Private cur_ps_id

Private Sub button_edit_Click()
    For i = 0 To list_search.ListCount - 1
        If IsNull(list_search.List(i, 3)) Or list_search.List(i, 3) = "" Then
            If list_search.List(i, 1) = "Y" Then
                Set cur_options = TestOptionLinkFactory(Array( _
                                                                Array("ps_id", cur_ps_id), _
                                                                Array("option_id", CDbl(list_search.List(i, 0))) _
                                                      ))
                If cur_options.count > 0 Then
                    For Each o In cur_options.fetch
                        o.delete
                    Next
                End If
            End If
        Else
            If list_search.List(i, 1) = "Y" Then
                Set cur_options = TestOptionLinkFactory(Array( _
                                                                Array("ps_id", cur_ps_id), _
                                                                Array("option_id", CDbl(list_search.List(i, 0))) _
                                                      ))
                If cur_options.count > 0 Then
                    Set o = cur_options.first
                    o.ps_id = cur_ps_id
                    o.option_id = CDbl(list_search.List(i, 0))
                    o.item = Trim(list_search.List(i, 3))
                    o.save
                Else
                    Set o = New TestOptionLink
                    o.ps_id = cur_ps_id
                    o.option_id = CDbl(list_search.List(i, 0))
                    o.item = Trim(list_search.List(i, 3))
                    o.save
                End If
            End If
            array_push_not_exist test_procedure_option_search_results, list_search.List(i, 2) & ":=" & list_search.List(i, 3)
        End If
    Next
    data_selected = True
    Unload Me
End Sub

Private Sub button_refresh_Click()
    text_search.text = ""
    init cur_ps_id
End Sub

Private Sub list_search_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    select_item list_search.ListIndex
End Sub

Private Sub list_search_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = Asc(" ") Then
        select_item list_search.ListIndex
    End If
End Sub

Public Sub select_item(list_index)
   If IsNull(list_search.List(list_index, 3)) Then
        new_value = ""
   Else
        new_value = list_search.List(list_index, 3)
   End If
   new_option = InputBox("Enter new option value", "Test Option", new_value)
   list_search.List(list_index, 1) = "Y"
   list_search.List(list_index, 3) = new_option
End Sub

Public Sub init(ps_id)
    cur_ps_id = CDbl(ps_id)
    list_search.Clear
    Set factory = TestOptionFactory(Array("name", "LIKE", "'%%'"))
    If factory.count > 0 Then
        For Each item In factory.fetch
            list_search.AddItem
            list_search.List(list_search.ListCount - 1, 0) = item.id
            list_search.List(list_search.ListCount - 1, 2) = item.name
            Set option_values = TestOptionLinkFactory(Array( _
                                                            Array("ps_id", cur_ps_id), _
                                                            Array("option_id", CDbl(item.id)) _
                                                      ))
            If option_values.count > 0 Then
                list_search.List(list_search.ListCount - 1, 3) = option_values.first.item
            End If
            list_search.List(list_search.ListCount - 1, 4) = item.description
        Next
    End If
End Sub

Private Sub text_search_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        search_string = Trim(text_search.text)
        If Trim(search_string) = "" Then Exit Sub
        If list_search.ListCount > 0 Then
            If list_search.ListIndex = -1 Then list_search.ListIndex = 0
        End If
        For i = list_search.ListIndex To list_search.ListCount - 1
            If InStr(1, list_search.List(i, 2), text_search.text, vbTextCompare) <> 0 And i <> list_search.ListIndex Then
                list_search.ListIndex = i
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub UserForm_Activate()
    text_search.SetFocus
End Sub

Private Sub UserForm_Terminate()
    If data_selected <> True Then test_procedure_option_search_results = Null
End Sub
