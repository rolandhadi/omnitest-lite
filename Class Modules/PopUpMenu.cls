VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PopUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Public Function create_menu(menu_items)
    Const menu_prefix = "Omni_"
    
    custom_menu_name = menu_prefix & menu_items("name")
    
    Call delete_menu(Left(custom_menu_name, 50))
    
    Set command_bar = command_bar_add(Left(custom_menu_name, 50))
    
    For Each menu_item In menu_items("menu")
        Set command_bar_control = command_bar.Controls.Add
        With command_bar_control
            .Caption = Trim(menu_items("menu")(menu_item)("caption"))
            .OnAction = Trim(menu_items("menu")(menu_item)("action"))
        End With
    Next
    
    Set create_menu = command_bar
    Set command_bar_control = Nothing
    Set command_bar = Nothing
End Function

Public Sub delete_menu(command_button_name)
    On Error Resume Next
    Application.CommandBars(command_button_name).delete
End Sub

Function command_bar_add(custom_menu_name)
Set command_bar_add = CommandBars.Add(name:=custom_menu_name, _
    Position:=5, _
    menuBar:=False, _
    Temporary:=False)
End Function

Function control_add(menu_bar, menu_type)
    Set control_add = menu_bar.Controls.Add(menu_type)
End Function

