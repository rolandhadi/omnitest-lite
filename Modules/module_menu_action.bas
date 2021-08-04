Attribute VB_Name = "module_menu_action"
Public disable_keyboard_check


Public Function evaluate_right_click(Target)
    
    If disable_keyboard_check = True Then Exit Function
    Set new_popup = New PopUpMenu
    Set new_menu_content = CreateObject("Scripting.Dictionary")
    Set new_menu_content("menu") = CreateObject("Scripting.Dictionary")
    
    valid_target = False
    
    If Target(1).Worksheet.name = TEST_OBJECT_TAB Then
    
        new_menu_content("name") = TEST_OBJECT_TAB
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "test_object_clear"), _
                                        Array("2", "Find", "test_object_find") _
                                )
            valid_target = True
        
        ElseIf Target(1).Column = 1 Then
        
            If Target(1).value = "Value ID" Then
            
                new_popup_content = Array( _
                                            Array("1", "Add Object Value", "test_object_value_add") _
                                    )
                valid_target = True
                
            ElseIf Trim(Target(1).value) <> "" Then
            
                If Left(Target(1).value, 2) = "TO" Then
                    
                    new_popup_content = Array( _
                                            Array("1", "Update Test Object", "test_object_update"), _
                                            Array("2", "Delete Test Object", "test_object_delete") _
                                    )
                    valid_target = True
                    
                ElseIf Left(Target(1).value, 2) = "OV" Then
                
                    new_popup_content = Array( _
                                            Array("1", "Update Value", "test_object_value_update"), _
                                            Array("2", "Delete Value", "test_object_value_delete") _
                                    )
                    valid_target = True
                    
                ElseIf Target(1).value = "NEW" Then
                
                    new_popup_content = Array( _
                                            Array("1", "Save", "test_object_new") _
                                    )
                    valid_target = True
                
                End If
                
            End If
            
        End If
        
    ElseIf Target(1).Worksheet.name = TEST_DATA_TAB Then
    
        new_menu_content("name") = TEST_DATA_TAB
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "test_data_clear"), _
                                        Array("2", "Find", "test_data_find") _
                                )
            valid_target = True
        
        ElseIf Target(1).Column = 1 Then
        
            If Target(1).value = "Value ID" Then
            
                new_popup_content = Array( _
                                            Array("1", "Add Iteration", "test_data_value_add") _
                                    )
                valid_target = True
                
            ElseIf Trim(Target(1).value) <> "" Then
            
                If Left(Target(1).value, 2) = "TD" Then
                    
                    new_popup_content = Array( _
                                            Array("1", "Update Test Data", "test_data_update"), _
                                            Array("2", "Delete Test Data", "test_data_delete") _
                                    )
                    valid_target = True
                    
                ElseIf Left(Target(1).value, 2) = "DV" Then
                
                    new_popup_content = Array( _
                                            Array("1", "Update Value", "test_data_value_update"), _
                                            Array("2", "Delete Value", "test_data_value_delete") _
                                    )
                    valid_target = True
                    
                ElseIf Target(1).value = "NEW" Then
                
                    new_popup_content = Array( _
                                            Array("1", "Save", "test_data_new") _
                                    )
                    valid_target = True
                
                End If
                
            End If
            
        End If
        
    ElseIf Target(1).Worksheet.name = TEST_OPTION_TAB Then
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "test_option_clear"), _
                                        Array("2", "Find", "test_option_find"), _
                                        Array("3", "Delete", "test_option_delete"), _
                                        Array("4", "Update", "test_option_update") _
                                )
            valid_target = True
            
        End If
        
    ElseIf Target(1).Worksheet.name = TEST_PROCEDURE_TAB Then
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "test_procedure_clear"), _
                                        Array("2", "New", "test_procedure_new"), _
                                        Array("3", "Find", "test_procedure_find") _
                                )
            valid_target = True
            
        ElseIf Target(1).row = 2 And Target(1).Column = 1 And Not IsEmpty(Target(1).value) Then
            
            new_popup_content = Array( _
                                        Array("1", "Update Test Procedure", "test_procedure_update"), _
                                        Array("2", "Delete Test Procedure", "test_procedure_delete") _
                                )
            valid_target = True
            
        ElseIf Target(1).row > 4 And Target(1).Column = 1 And Not IsEmpty(Application.Range("A" & 2).value) Then
        
            If Not IsEmpty(Target(1).value) Then
                new_popup_content = Array( _
                                            Array("1", "Insert New Step", "test_procedure_step_insert"), _
                                            Array("2", "Delete Step", "test_procedure_step_delete") _
                                    )
            Else
                new_popup_content = Array( _
                                            Array("1", "Insert New Step", "test_procedure_step_new") _
                                    )
            End If
            
            valid_target = True
            
        
        ElseIf Target(1).row > 4 And Target(1).Column = 2 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Order", "test_procedure_step_order") _
                                )
            
            valid_target = True
            
        ElseIf Target(1).row > 4 And Target(1).Column = 3 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Keyword", "test_procedure_step_keyword") _
                                )
            
            valid_target = True
            
        ElseIf Target(1).row > 4 And Target(1).Column = 4 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Object", "test_procedure_step_object"), _
                                            Array("2", "Clear Test Object", "test_procedure_step_clear_object") _
                                )
            
            valid_target = True
            
        ElseIf Target(1).row > 4 And (Target(1).Column = 5) And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Data", "test_procedure_step_data_in"), _
                                            Array("2", "Clear", "test_procedure_step_clear_data_in") _
                                )
            
            valid_target = True
        
        ElseIf Target(1).row > 4 And (Target(1).Column = 6) And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Data", "test_procedure_step_data_out"), _
                                            Array("2", "Clear", "test_procedure_step_clear_data_out") _
                                )
            
            valid_target = True
            
        ElseIf Target(1).row > 4 And Target(1).Column = 7 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Option", "test_procedure_step_option") _
                                )
            
            valid_target = True
        
        End If
        
    
    ElseIf Target(1).Worksheet.name = TEST_CASE_TAB Then
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "test_case_clear"), _
                                        Array("2", "New", "test_case_new"), _
                                        Array("3", "Find", "test_case_find") _
                                )
            valid_target = True
            
        ElseIf Target(1).row = 2 And Target(1).Column = 1 And Not IsEmpty(Target(1).value) Then
            
            new_popup_content = Array( _
                                        Array("1", "Update Test Case", "test_case_update"), _
                                        Array("2", "Delete Test Case", "test_case_delete") _
                                )
            valid_target = True
            
        ElseIf Target(1).row > 4 And Target(1).Column = 1 And Not IsEmpty(Application.Range("A" & 2).value) Then
        
            If Not IsEmpty(Target(1).value) Then
                new_popup_content = Array( _
                                            Array("1", "Insert Test Procedure", "test_case_step_insert"), _
                                            Array("2", "Delete Test Procedure", "test_case_step_delete") _
                                    )
            Else
                new_popup_content = Array( _
                                            Array("1", "Insert Test Procedure", "test_case_step_new") _
                                    )
            End If
            
            valid_target = True
            
        
        ElseIf Target(1).row > 4 And Target(1).Column = 2 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Order", "test_case_step_order") _
                                )
            
            valid_target = True
            
        
        ElseIf Target(1).row > 4 And Target(1).Column = 3 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Procedure", "test_case_step_test_procedure") _
                                )
            
            valid_target = True
            
        
        End If
        
    
    ElseIf Target(1).Worksheet.name = TEST_SCENARIO_TAB Then
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "test_scenario_clear"), _
                                        Array("2", "New", "test_scenario_new"), _
                                        Array("3", "Find", "test_scenario_find") _
                                )
            valid_target = True
            
        ElseIf Target(1).row = 2 And Target(1).Column = 1 And Not IsEmpty(Target(1).value) Then
            
            new_popup_content = Array( _
                                        Array("1", "Update Test Scenario", "test_scenario_update"), _
                                        Array("2", "Delete Test Scenario", "test_scenario_delete") _
                                )
            valid_target = True
            
        ElseIf Target(1).row > 4 And Target(1).Column = 1 And Not IsEmpty(Application.Range("A" & 2).value) Then
        
            If Not IsEmpty(Target(1).value) Then
                new_popup_content = Array( _
                                            Array("1", "Insert Test Case", "test_scenario_step_insert"), _
                                            Array("2", "Delete Test Case", "test_scenario_step_delete") _
                                    )
            Else
                new_popup_content = Array( _
                                            Array("1", "Insert Test Case", "test_scenario_step_new") _
                                    )
            End If
            
            valid_target = True
            
        
        ElseIf Target(1).row > 4 And Target(1).Column = 2 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Order", "test_scenario_step_order") _
                                )
            
            valid_target = True
            
            
        ElseIf Target(1).row > 4 And Target(1).Column = 3 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Case", "test_scenario_step_test_case") _
                                )
            
            valid_target = True
            
        
        End If
        
    
    ElseIf Target(1).Worksheet.name = TEST_SCENARIO_DATA_TAB Then
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "test_scenario_data_clear"), _
                                        Array("2", "Find", "test_scenario_data_find") _
                                )
            valid_target = True
            
            
        ElseIf Target(1).row > 4 And (Target(1).Column = 9) And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Data", "test_scenario_data_step_data_in"), _
                                            Array("2", "Update Reference Data", "test_scenario_data_step_ref_in"), _
                                            Array("3", "Clear", "test_scenario_data_step_clear_data_in") _
                                )
            
            valid_target = True
        
        ElseIf Target(1).row > 4 And (Target(1).Column = 10) And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Data", "test_scenario_data_step_data_out"), _
                                            Array("2", "Clear", "test_scenario_data_step_clear_data_out") _
                                )
            
            valid_target = True
        
        End If
        
        
    ElseIf Target(1).Worksheet.name = PLAN_EXECUTION_TAB Then
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "plan_execution_clear"), _
                                        Array("2", "New", "plan_execution_new"), _
                                        Array("3", "Find", "plan_execution_find") _
                                )
            valid_target = True
            
        ElseIf Target(1).row = 2 And Target(1).Column = 1 And Not IsEmpty(Target(1).value) Then
            
            new_popup_content = Array( _
                                        Array("1", "Update Plan Execution", "plan_execution_update"), _
                                        Array("2", "Delete Plan Execution", "plan_execution_delete"), _
                                        Array("3", "Execute Plan Execution", "plan_execution_execute") _
                                )
            valid_target = True
            
        ElseIf Target(1).row = 2 And Target(1).Column = 3 And Not IsEmpty(Target(1).value) Then
            
            new_popup_content = Array( _
                                        Array("1", "Stop Execution", "plan_execution_stop_exec"), _
                                        Array("2", "Continue Execution", "plan_execution_continue_exec") _
                                )
            valid_target = True
            
        ElseIf Target(1).row > 4 And Target(1).Column = 1 And Not IsEmpty(Application.Range("A" & 2).value) Then
        
            If Not IsEmpty(Target(1).value) Then
                new_popup_content = Array( _
                                            Array("1", "Insert Test Scenario", "plan_execution_step_insert"), _
                                            Array("2", "Delete Test Scenario", "plan_execution_step_delete") _
                                    )
            Else
                new_popup_content = Array( _
                                            Array("1", "Insert Test Scenario", "plan_execution_step_new") _
                                    )
            End If
            
            valid_target = True
            
        
        ElseIf Target(1).row > 4 And Target(1).Column = 2 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Order", "plan_execution_step_order") _
                                )
            
            valid_target = True
            
            
        ElseIf Target(1).row > 4 And Target(1).Column = 3 And Not IsEmpty(Application.Range("A" & Target(1).row).value) Then
        
            new_popup_content = Array( _
                                            Array("1", "Update Test Scenario", "plan_execution_step_test_case") _
                                )
            
            valid_target = True
            
        
        End If
        
    ElseIf Target(1).Worksheet.name = FUNCTION_REFERENCE_TAB Then
        
        If Target(1).row = 1 Then
        
            new_popup_content = Array( _
                                        Array("1", "Clear", "function_reference_clear"), _
                                        Array("2", "New", "function_reference_new"), _
                                        Array("3", "Find", "function_reference_find") _
                                )
            valid_target = True
            
        ElseIf Target(1).row = 2 And Target(1).Column = 1 And Not IsEmpty(Target(1).value) Then
            
            new_popup_content = Array( _
                                        Array("1", "Update Function Reference", "function_reference_update"), _
                                        Array("2", "Delete Function Reference", "function_reference_delete") _
                                )
            valid_target = True
            
            
        ElseIf Target(1).row > 4 And Target(1).Column = 1 And Not IsEmpty(Application.Range("A" & 2).value) Then
        
            If Not IsEmpty(Target(1).value) Then
                new_popup_content = Array( _
                                            Array("1", "Insert Row", "function_reference_step_insert"), _
                                            Array("2", "Delete Row", "function_reference_step_delete") _
                                    )
            Else
                new_popup_content = Array( _
                                            Array("1", "Insert Row", "function_reference_step_new") _
                                    )
            End If
            
            valid_target = True
            
        
        End If
        
    End If
    
    If valid_target Then
        For Each content In new_popup_content
            Set new_menu_content("menu")(content(0)) = CreateObject("Scripting.Dictionary")
            new_menu_content("menu")(content(0))("caption") = Trim(content(1))
            new_menu_content("menu")(content(0))("action") = Trim(content(2))
        Next
        Set new_menu = new_popup.create_menu(new_menu_content)
        new_menu.ShowPopup
    End If
    
    evaluate_right_click = valid_target
    
End Function


Public Sub evaluate_value_change(Target)

    If disable_keyboard_check = True Then Exit Sub
    
    If Target(1).Worksheet.name = TEST_OBJECT_TAB Then
    
        Select Case Target(1).Column
        
            
        
        End Select
    
    ElseIf Target(1).Worksheet.name = TEST_OPTION_TAB Then
        
        
        
    ElseIf Target(1).Worksheet.name = TEST_PROCEDURE_TAB Then
    
        If Target.Column = 5 And IsNumeric(Application.Range("A" & Target(1).row).value) And Target(1).row > 4 Then
            
            test_procedure_step_data_in_value Target(1).row
            
        ElseIf Target.Column = 6 And IsNumeric(Application.Range("A" & Target(1).row).value) And Target(1).row > 4 Then
            
            test_procedure_step_data_out_value Target(1).row
            
        End If
        
    ElseIf Target(1).Worksheet.name = TEST_CASE_TAB Then
    
    ElseIf Target(1).Worksheet.name = TEST_SCENARIO_TAB Then
    
    ElseIf Target(1).Worksheet.name = TEST_SCENARIO_DATA_TAB Then
    
        If Target.Column = 9 And Not IsEmpty(Application.Range("A" & Target(1).row).value) And Target(1).row > 4 Then
            
            test_scenario_data_step_data_in_value Target(1).row
            
        ElseIf Target.Column = 10 And Not IsEmpty(Application.Range("A" & Target(1).row).value) And Target(1).row > 4 Then
            
            test_scenario_data_step_data_out_value Target(1).row
            
        End If
        
    ElseIf Target(1).Worksheet.name = FUNCTION_REFERENCE_TAB Then
    
        If Target.row = 4 Then
            
            function_reference_update_column Target(1).Column
            
        End If
        
    End If

End Sub
