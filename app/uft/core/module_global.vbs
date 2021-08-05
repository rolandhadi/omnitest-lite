'Attribute VB_Name = "module_global"
Public Const TEST_DATA_TAB = "Test_Data"
Public Const FUNCTION_REFERENCE_TAB = "Function_References"
Public Const FUNCTION_REFERENCE_VALUE_TAB = "Function_Reference_Values"
Public Const PLAN_EXECUTION_TAB = "Plan_Executions"
Public Const TEST_SCENARIO_DATA_TAB = "Scenario_Data"
Public Const TEST_SCENARIO_TAB = "Test_Scenarios"
Public Const TEST_CASE_TAB = "Test_Cases"
Public Const TEST_PROCEDURE_TAB = "Test_Procedures"
Public Const TEST_OPTION_TAB = "Test_Options"
Public Const TEST_OBJECT_TAB = "Test_Objects"
Public CURRENT_CONNECTION
Public LIST_CHECKED
Public LIST_UNCHECKED
Public session
Public SESSION_NAME
Public DATABASE_TYPE
Public DATABASE_PROVIDER
Public DATABASE_PATH

Public Function omnilite_init()
    LIST_CHECKED = ChrW(&H2611)
    LIST_UNCHECKED = ChrW(&H2610)
    DATABASE_TYPE = UCase(Trim(config("DATABASE_TYPE")))
    DATABASE_PROVIDER = config("DATABASE_PROVIDER")
    DATABASE_PATH = config("DATABASE_PATH")
    If DATABASE_TYPE = "ACCESS" Then
        CURRENT_CONNECTION = DATABASE_PROVIDER & "Data Source=" & DATABASE_PATH
    ElseIf DATABASE_TYPE = "SQLITE" Then
        CURRENT_CONNECTION = DATABASE_PATH
    End If
    disable_keyboard_check = False
    SESSION_NAME = RandomString(27)
End Function

Public Function config(env_name)
    config = get_conf_env(ROOT_FOLDER & "\config\app.config", env_name)
End Function

Public Function env(env_name)
    config = get_conf_env(ROOT_FOLDER & "\.env", env_name)
End Function

Public Function get_conf_env(file_path, env_name)

    env_text = CreateObject("Scripting.FileSystemObject").openTextFile(file_path).readAll()
    Set regex = CreateObject("VBScript.RegExp")
    regex_pattern = env_name & "\s*=\S*"
    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = regex_pattern
    If regex.Test(env_text) Then
      Set matches = regex.execute(env_text)
      For Each Match In matches
        value_count = 1
        For Each value In Split(Match, "=")
            If value_count > 1 Then
                function_output = function_output & value & "="
            End If
            value_count = value_count + 1
        Next
        get_conf_env = Trim(Left(function_output, Len(function_output) - 1))
        Exit Function
      Next
    Else
      get_conf_env = ""
    End If

  End Function

Public Function is_enter(k)
    If Asc(k) = 13 Then
        is_enter = True
    Else
        is_enter = False
    End If
End Function

Public Function evaluate_where_params(exp)
    order_by_params = ""
    If IsArray(exp) Then
        If IsArray(exp(LBound(exp))) Then
            out = ""
            For Each E In exp
                If UBound(E) = 2 Then
                    out = out & E(0) & " " & E(1) & " " & E(2) & "" & " AND "
                ElseIf UBound(E) = 1 Then
                    If IsNull(E(1)) Then
                        out = out & E(0) & " IS " & "NULL" & "" & " AND "
                    Else
                        out = out & E(0) & " = " & E(1) & "" & " AND "
                    End If
                ElseIf UBound(E) = 0 Then
                    If IsNull(E(0)) Then
                        out = out & "id IS " & "NULL" & "" & " AND "
                    Else
                        out = out & "id = " & E(0) & "" & " AND "
                    End If
                End If
            Next
            If Len(out) > 0 Then
                out = Left(out, Len(out) - 5)
            End If
            evaluate_where_params = out
        Else
            If UBound(exp) = 2 Then
                If IsArray(exp(2)) Then
                    evaluate_where_params = exp(0) & " " & " = " & " " & exp(1) & ""
                    If UBound(exp(2)) = 1 Then
                        order_by_params = " ORDER BY " & exp(2)(0) & " " & exp(2)(1)
                    Else
                        order_by_params = " ORDER BY " & exp(2)(0)
                    End If
                Else
                    evaluate_where_params = exp(0) & " " & exp(1) & " " & exp(2) & ""
                End If
            ElseIf UBound(exp) = 1 Then
                If IsNull(exp(1)) Then
                    evaluate_where_params = exp(0) & " IS " & "NULL" & ""
                Else
                    evaluate_where_params = exp(0) & " = " & exp(1) & ""
                End If
            ElseIf UBound(exp) = 0 Then
                If IsNull(exp(0)) Then
                    evaluate_where_params = "id IS " & "NULL" & ""
                Else
                    evaluate_where_params = "id = " & exp(0) & ""
                End If
            End If
        End If
    Else
        evaluate_where_params = "id = " & exp
    End If
    evaluate_where_params = evaluate_where_params & order_by_params
End Function

Public Sub array_push(ByRef array_target, value)
    If IsArray(array_target) Then
        ReDim Preserve array_target(UBound(array_target) + 1)
        If IsObject(value) Then
            Set array_target(UBound(array_target)) = value
        Else
            array_target(UBound(array_target)) = value
        End If
    Else
        ReDim array_target(0)
        If IsObject(value) Then
            Set array_target(0) = value
        Else
            array_target(0) = value
        End If
    End If
End Sub

Public Sub array_push_not_exist(ByRef array_target, value)
    If IsArray(array_target) Then
        If IsObject(value) Then
            found = False
            For Each obj In array_target
                If obj.id = value.id Then
                    found = False
                    Exit For
                End If
            Next
            If Not found Then
                ReDim Preserve array_target(UBound(array_target) + 1)
                Set array_target(UBound(array_target)) = value
            End If
        Else
            found = False
            For Each obj In array_target
                If obj = value Then
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                ReDim Preserve array_target(UBound(array_target) + 1)
                array_target(UBound(array_target)) = value
            End If
        End If
    Else
        ReDim array_target(0)
        If IsObject(value) Then
            Set array_target(0) = value
        Else
            array_target(0) = value
        End If
    End If
End Sub

Public Function evaluate_insert_params(m_update)
    sql_set = Array("", "")
    For Each update_field In m_update
        If Not IsEmpty(m_update(update_field)) Then
            If Not IsNull(m_update(update_field)) Then
                sql_set(0) = sql_set(0) & update_field & ","
                sql_set(1) = sql_set(1) & m_update(update_field) & ","
            End If
        End If
    Next
    If Not IsEmpty(sql_set(0)) Then
        sql_set(0) = Left(sql_set(0), Len(sql_set(0)) - 1)
    End If
    If Not IsEmpty(sql_set(1)) Then
        sql_set(1) = Left(sql_set(1), Len(sql_set(1)) - 1)
    End If
    evaluate_insert_params = sql_set
End Function

Public Function evaluate_set_params(m_update)
    sql_set = ""
    For Each update_field In m_update
        If Not IsEmpty(m_update(update_field)) Then
            If IsNull(m_update(update_field)) Then
                sql_set = sql_set & update_field & " = " & "NULL" & ","
            Else
                sql_set = sql_set & update_field & " = " & m_update(update_field) & ","
            End If
        End If
    Next
    If Not IsEmpty(sql_set) Then
        sql_set = Left(sql_set, Len(sql_set) - 1)
    End If
    evaluate_set_params = sql_set
End Function

Public Function db_find_or_new_record(name, table_name)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT id FROM " & table_name & " " & _
                            "WHERE name = '" & name & "'")
    If IsEmpty(m_id) Then
        sql_set = evaluate_insert_params(m_update)
        cn.connection.execute "INSERT INTO " & table_name & " " & _
            "(" & sql_set(0) & ") VALUES " & _
            "(" & sql_set(1) & ")"
    Else
        sql_set = evaluate_set_params(m_update)
        cn.connection.execute "UPDATE " & table_name & " " & _
            "SET " & sql_set & " " & _
            "WHERE id = " & m_id
    End If
    db_find_or_new_record = cn.connection.last_id
End Function

Public Function db_get_sql(m_id, m_update, table_name)
    Set cn = New connection
    If IsEmpty(m_id) Then
        sql_set = evaluate_insert_params(m_update)
        db_get_sql = "INSERT INTO " & table_name & " " & _
            "(" & sql_set(0) & ") VALUES " & _
            "(" & sql_set(1) & ")"
    Else
        sql_set = evaluate_set_params(m_update)
        db_get_sql = "UPDATE " & table_name & " " & _
            "SET " & sql_set & " " & _
            "WHERE id = " & m_id
    End If
End Function

Public Function db_save_batch(sqls)
    Set cn = New connection
    cn.connection.execute sqls
End Function

Public Function db_save_record(m_id, m_update, table_name)
    Set cn = New connection
    If IsEmpty(m_id) Then
        sql_set = evaluate_insert_params(m_update)
        cn.connection.execute "INSERT INTO " & table_name & " " & _
            "(" & sql_set(0) & ") VALUES " & _
            "(" & sql_set(1) & ")"
    Else
        sql_set = evaluate_set_params(m_update)
        cn.connection.execute "UPDATE " & table_name & " " & _
            "SET " & sql_set & " " & _
            "WHERE id = " & m_id
    End If
    db_save_record = cn.connection.last_id
End Function

Public Function db_delete_record(m_id, table_name)
    Set cn = New connection
    If Not IsEmpty(m_id) Then
        cn.connection.execute "DELETE FROM " & table_name & " WHERE id = " & m_id
    End If
    db_delete_record = m_id
End Function

Public Function CaseProcedureFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM case_procedures WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New CaseProcedure
        out("data")(id).id = id
        out("data")(id).m_tc_id = rs.data("tc_id")
        out("data")(id).m_tp_id = rs.data("tp_id")
        out("data")(id).m_order_no = rs.data("order_no")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set CaseProcedureFactory = factory
End Function

Public Function CaseScenarioFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM case_scenarios WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New CaseScenario
        out("data")(id).m_id = id
        out("data")(id).m_ts_id = rs.data("ts_id")
        out("data")(id).m_tc_id = rs.data("tc_id")
        out("data")(id).m_order_no = rs.data("order_no")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set CaseScenarioFactory = factory
End Function

Public Function FolderPathFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM folder_paths WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New FolderPath
        out("data")(id).m_id = id
        out("data")(id).m_module_code = rs.data("module_code")
        out("data")(id).m_name = rs.data("name")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set FolderPathFactory = factory
End Function

Public Function FunctionReferenceFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM function_references WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New FunctionReference
        out("data")(id).id = id
        out("data")(id).m_name = rs.data("name")
        out("data")(id).m_type_code = rs.data("type_code")
        out("data")(id).m_struct = rs.data("struct")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set FunctionReferenceFactory = factory
End Function

Public Function FunctionReferenceValueFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM function_reference_values WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New FunctionReferenceValue
        out("data")(id).id = id
        out("data")(id).m_fr_id = rs.data("fr_id")
        out("data")(id).m_name = rs.data("name")
        out("data")(id).m_item = rs.data("item")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set FunctionReferenceValueFactory = factory
End Function

Public Function ProcedureStepFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM procedure_steps WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New ProcedureStep
        out("data")(id).id = id
        out("data")(id).m_tp_id = rs.data("tp_id")
        out("data")(id).m_to_id = rs.data("to_id")
        out("data")(id).m_keyword_name = rs.data("keyword_name")
        out("data")(id).m_order_no = rs.data("order_no")
        out("data")(id).m_description = rs.data("description")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set ProcedureStepFactory = factory
End Function

Public Function TestCaseFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_cases WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestCase
        out("data")(id).id = id
        out("data")(id).m_name = rs.data("name")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestCaseFactory = factory
End Function

Public Function TestDataFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_datas WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestData
        out("data")(id).id = id
        out("data")(id).m_parent_id = rs.data("parent")
        out("data")(id).m_name = rs.data("name")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestDataFactory = factory
End Function

Public Function TestDataValueFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_data_values WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestDataValue
        out("data")(id).id = id
        out("data")(id).m_td_id = rs.data("td_id")
        out("data")(id).m_iteration = rs.data("iteration")
        out("data")(id).m_item = rs.data("item")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestDataValueFactory = factory
End Function

Public Function TestObjectFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_objects WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestObject
        out("data")(id).id = id
        out("data")(id).m_parent_id = rs.data("parent")
        out("data")(id).m_name = rs.data("name")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestObjectFactory = factory
End Function

Public Function TestObjectValueFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_object_values WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestObjectValue
        out("data")(id).id = id
        out("data")(id).m_to_id = rs.data("to_id")
        out("data")(id).m_type_code = rs.data("type_code")
        out("data")(id).m_item = rs.data("item")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestObjectValueFactory = factory
End Function

Public Function TestProcedureOptionFactory(exp)
    Set cn = New connection
    rs = cn.connection.get_batch_records(Array( _
                                                "SELECT * FROM test_options;", _
                                                "SELECT test_option_links.ps_id AS ps_id, test_option_links.option_id AS option_id, test_options.name AS name, test_option_links.item AS item " & _
                                                "     FROM test_option_links INNER JOIN test_options ON test_option_links.option_id = test_options.id " & _
                                                "     WHERE (((test_option_links.tp_id)=" & exp & "));" _
                                                ) _
                                        )
    TestProcedureOptionFactory = rs
End Function

Public Function TestOptionFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_options WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestOption
        out("data")(id).id = id
        out("data")(id).m_name = rs.data("name")
        out("data")(id).m_description = rs.data("description")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestOptionFactory = factory
End Function

Public Function TestOptionLinkFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_option_links WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestOptionLink
        out("data")(id).id = id
        out("data")(id).m_ps_id = rs.data("ps_id")
        out("data")(id).m_triggered = rs.data("triggered")
        out("data")(id).m_option_id = rs.data("option_id")
        out("data")(id).m_item = rs.data("item")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestOptionLinkFactory = factory
End Function

Public Function TestProcedureFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_procedures WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestProcedure
        out("data")(id).id = id
        out("data")(id).m_name = rs.data("name")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestProcedureFactory = factory
End Function

Public Function TestScenarioFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_scenarios WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestScenario
        out("data")(id).id = id
        out("data")(id).m_name = rs.data("name")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestScenarioFactory = factory
End Function

Public Function TestScenarioLinkFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM test_scenario_links WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New TestScenarioLink
        out("data")(id).id = id
        out("data")(id).m_ts_id = rs.data("ts_id")
        out("data")(id).m_tc_id = rs.data("tc_id")
        out("data")(id).m_tp_id = rs.data("tp_id")
        out("data")(id).m_cs_id = rs.data("cs_id")
        out("data")(id).m_cp_id = rs.data("cp_id")
        out("data")(id).m_ps_id = rs.data("ps_id")
        out("data")(id).m_execute = rs.data("execute")
        out("data")(id).m_screenshot = rs.data("screenshot")
        out("data")(id).m_data_value_ref_id = rs.data("data_value_ref_id")
        out("data")(id).m_data_value_id_in = rs.data("data_value_id_in")
        out("data")(id).m_data_value_in = rs.data("data_value_in")
        out("data")(id).m_data_value_id_out = rs.data("data_value_id_out")
        out("data")(id).m_data_value_out = rs.data("data_value_out")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set TestScenarioLinkFactory = factory
End Function

Public Function PlanExecutionFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM plan_executions WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New PlanExecution
        out("data")(id).id = id
        out("data")(id).m_name = rs.data("name")
        out("data")(id).m_next_action = rs.data("next_action")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set PlanExecutionFactory = factory
End Function

Public Function ExecutionScenarioFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records("SELECT * FROM execution_scenarios WHERE " & evaluate_where_params(exp))
    Set out = CreateObject("Scripting.Dictionary")
    out("count") = rs.count
    Set out("data") = CreateObject("Scripting.Dictionary")
    While Not rs.eof
        id = rs.data("id")
        Set out("data")(id) = New ExecutionScenario
        out("data")(id).m_id = id
        out("data")(id).m_pe_id = rs.data("pe_id")
        out("data")(id).m_ts_id = rs.data("ts_id")
        out("data")(id).m_order_no = rs.data("order_no")
        out("data")(id).m_iteration = rs.data("iteration")
        out("data")(id).m_execute = rs.data("execute")
        out("data")(id).m_dependency = rs.data("dependency")
        rs.move_next
    Wend
    Set factory = New TestFactory
    factory.init out
    Set ExecutionScenarioFactory = factory
End Function

Public Function ExecutionScenarioDataFactory(exp)
    Set cn = New connection
    rs = cn.connection.get_batch_records(Array("SELECT execution_scenarios.id AS es_id, plan_executions.name AS pe_name, execution_scenarios.iteration AS iteration, execution_scenarios.execute AS run, execution_scenarios.dependency AS dependency, " & _
                                                    "case_scenarios.id AS cs_id, case_procedures.id AS cp_id, procedure_steps.id AS ps_id, execution_scenarios.order_no AS es_order, case_scenarios.order_no AS cs_order, " & _
                                                    "case_procedures.order_no AS cp_order, procedure_steps.order_no AS ps_order, test_scenarios.id AS ts_id, test_scenarios.name AS ts_name, test_cases.id AS tc_id, " & _
                                                    "test_cases.name AS tc_name, test_procedures.id AS tp_id, test_procedures.name AS tp_name, test_objects.id AS to_id, test_objects.name AS to_name, procedure_steps.keyword_name AS keyword_name " & _
                                                    "FROM plan_executions INNER JOIN (((test_scenarios INNER JOIN ((test_procedures INNER JOIN (test_cases INNER JOIN case_procedures ON test_cases.id = case_procedures.tc_id) ON " & _
                                                    "test_procedures.id = case_procedures.tp_id) INNER JOIN case_scenarios ON test_cases.id = case_scenarios.tc_id) ON test_scenarios.id = case_scenarios.ts_id) INNER JOIN " & _
                                                    "(procedure_steps LEFT JOIN test_objects ON procedure_steps.to_id = test_objects.id) ON test_procedures.id = procedure_steps.tp_id) INNER JOIN execution_scenarios ON " & _
                                                    "test_scenarios.id = execution_scenarios.ts_id) ON plan_executions.id = execution_scenarios.pe_id " & _
                                                    "WHERE (((plan_executions.id) = " & exp & ")) " & _
                                                    "ORDER BY execution_scenarios.order_no, case_scenarios.order_no, case_procedures.order_no,procedure_steps.order_no;", _
                                                "SELECT test_scenario_links.id AS id, test_scenario_links.cs_id AS cs_id, test_scenario_links.cp_id AS cp_id, test_scenario_links.tp_id AS tp_id, test_scenario_links.ps_id AS ps_id, test_datas.name AS data_name, test_scenario_links.data_value_ref_id AS data_value_ref_id, function_reference_values.name AS ref_name, function_reference_values.item AS ref_value, test_scenario_links.data_value_id_in AS data_value_id_in, test_scenario_links.data_value_in AS data_value_in, test_scenario_links.data_value_id_out AS data_value_id_out, test_scenario_links.data_value_out AS data_value_out, " & _
                                                    "test_scenario_links.execute AS run, test_scenario_links.screenshot AS screenshot " & _
                                                    "FROM plan_executions INNER JOIN (((test_scenario_links LEFT JOIN test_datas ON test_scenario_links.data_value_id_in = test_datas.id) LEFT JOIN function_reference_values ON test_scenario_links.data_value_ref_id = function_reference_values.id) INNER JOIN execution_scenarios ON test_scenario_links.ts_id = execution_scenarios.ts_id) ON plan_executions.id = execution_scenarios.pe_id " & _
                                                    "WHERE (((plan_executions.id) = " & exp & ")) " & _
                                                    "ORDER BY execution_scenarios.order_no;", _
                                                "SELECT test_option_links.ps_id, test_options.name, test_option_links.item " & _
                                                    "FROM test_options INNER JOIN test_option_links ON test_options.id = test_option_links.option_id " & _
                                                    "WHERE (((test_option_links.tp_id) In (SELECT DISTINCT test_scenario_links.tp_id " & _
                                                    "FROM execution_scenarios INNER JOIN test_scenario_links ON execution_scenarios.ts_id = test_scenario_links.ts_id " & _
                                                    "WHERE execution_scenarios.pe_id = " & exp & ")));" _
                                            ) _
                                        )
    '"SELECT test_objects.id AS to_id, test_objects.name AS to_name, test_object_values.type_code AS type_code, test_object_values.item  AS item" & _
    '    "FROM test_objects INNER JOIN test_object_values ON test_objects.id = test_object_values.to_id " & _
    '    "WHERE (((test_objects.id) In (SELECT procedure_steps.to_id " & _
    '    "FROM procedure_steps INNER JOIN (execution_scenarios INNER JOIN test_scenario_links ON execution_scenarios.ts_id = test_scenario_links.ts_id) ON procedure_steps.id = test_scenario_links.ps_id WHERE (((procedure_steps.to_id) Is Not Null) AND ((execution_scenarios.pe_id)=" & exp & ")))));"
    ExecutionScenarioDataFactory = rs
End Function

Public Function ScenarioDataFactory(exp)
    Set cn = New connection
    rs = cn.connection.get_batch_records(Array("SELECT case_scenarios.id AS cs_id, case_procedures.id AS cp_id, procedure_steps.id AS ps_id, case_scenarios.order_no AS cs_order, case_procedures.order_no AS cp_order, procedure_steps.order_no AS ps_order, test_cases.id AS tc_id, " & _
                                               "    test_cases.name AS tc_name, test_procedures.id AS tp_id, test_procedures.name AS tp_name, test_objects.name AS to_name, procedure_steps.keyword_name AS keyword_name " & _
                                               "    FROM (test_scenarios INNER JOIN ((test_procedures INNER JOIN (test_cases INNER JOIN case_procedures ON test_cases.id = case_procedures.tc_id) ON test_procedures.id =  " & _
                                               "    case_procedures.tp_id) INNER JOIN case_scenarios ON test_cases.id = case_scenarios.tc_id) ON test_scenarios.id = case_scenarios.ts_id) INNER JOIN (procedure_steps LEFT  " & _
                                               "    JOIN test_objects ON procedure_steps.to_id = test_objects.id) ON test_procedures.id = procedure_steps.tp_id " & _
                                               "    WHERE (((test_scenarios.id)=" & exp & ")) " & _
                                               "    ORDER BY case_scenarios.order_no, case_procedures.order_no, procedure_steps.order_no;", _
                                               "SELECT test_scenario_links.id AS id, test_scenario_links.cs_id AS cs_id, test_scenario_links.cp_id AS cp_id, test_scenario_links.tp_id AS tp_id, test_scenario_links.ps_id AS ps_id, test_datas.name AS name, test_scenario_links.data_value_ref_id AS data_value_ref_id, function_reference_values.name AS ref_name, test_scenario_links.data_value_id_in AS data_value_id_in, test_scenario_links.data_value_in AS data_value_in, " & _
                                               "test_scenario_links.screenshot AS screenshot, test_scenario_links.execute AS execute " & _
                                               "     FROM (test_scenario_links LEFT JOIN test_datas ON test_scenario_links.data_value_id_in = test_datas.id)  LEFT JOIN function_reference_values ON function_reference_values.id = test_scenario_links.data_value_ref_id " & _
                                               "     WHERE (((test_scenario_links.ts_id)=" & exp & ")); ", _
                                               "SELECT test_scenario_links.id AS id, test_scenario_links.cp_id AS cp_id, test_scenario_links.tp_id AS tp_id, test_scenario_links.ps_id AS ps_id, test_datas.name AS name, test_scenario_links.data_value_id_out AS data_value_id_out, test_scenario_links.data_value_out AS data_value_out " & _
                                               "     FROM test_scenario_links LEFT JOIN test_datas ON test_scenario_links.data_value_id_out = test_datas.id " & _
                                               "     WHERE (((test_scenario_links.ts_id)=" & exp & ")); " _
                                            ) _
                                        )
    ScenarioDataFactory = rs
End Function

Public Function ProcedureStepDataFactory(exp)
    Set cn = New connection
    rs = cn.connection.get_batch_records(Array( _
                                                "SELECT procedure_steps.id AS id, procedure_steps.order_no AS order_no, procedure_steps.keyword_name AS keyword_name, test_objects.name AS name " & _
                                                "FROM test_procedures INNER JOIN (procedure_steps LEFT JOIN test_objects ON procedure_steps.to_id = test_objects.id) ON test_procedures.id = procedure_steps.tp_id " & _
                                                "     WHERE (((test_procedures.id)=" & exp & ")) " & _
                                                "     ORDER BY procedure_steps.order_no;", _
                                                "SELECT test_scenario_links.id AS id, test_scenario_links.tp_id AS tp_id, test_scenario_links.ps_id AS ps_id, test_datas.name AS name, test_scenario_links.data_value_ref_id AS data_value_ref_id, function_reference_values.name AS ref_name, test_scenario_links.data_value_id_in AS data_value_id_in, test_scenario_links.data_value_in AS data_value_in " & _
                                                "     FROM (test_scenario_links LEFT JOIN test_datas ON test_scenario_links.data_value_id_in = test_datas.id)  LEFT JOIN function_reference_values ON function_reference_values.id = test_scenario_links.data_value_ref_id " & _
                                                "     WHERE (((test_scenario_links.tp_id)=" & exp & ")) " & _
                                                "     AND (test_scenario_links.ts_id) IS NULL " & _
                                                "     AND (test_scenario_links.tc_id) IS NULL " & _
                                                "     AND (test_scenario_links.cs_id) IS NULL " & _
                                                "     AND (test_scenario_links.cp_id) IS NULL;", _
                                                "SELECT test_scenario_links.id AS id, test_scenario_links.tp_id AS tp_id, test_scenario_links.ps_id AS ps_id, test_datas.name AS name, test_scenario_links.data_value_id_out AS data_value_id_out, test_scenario_links.data_value_out AS data_value_out " & _
                                                "     FROM test_scenario_links LEFT JOIN test_datas ON test_scenario_links.data_value_id_out = test_datas.id " & _
                                                "     WHERE (((test_scenario_links.tp_id)=" & exp & ")) " & _
                                                "     AND (test_scenario_links.ts_id) IS NULL " & _
                                                "     AND (test_scenario_links.tc_id) IS NULL " & _
                                                "     AND (test_scenario_links.cs_id) IS NULL " & _
                                                "     AND (test_scenario_links.cp_id) IS NULL;", _
                                                "SELECT test_option_links.ps_id AS ps_id, test_options.name AS name, test_option_links.item AS item " & _
                                                "     FROM test_option_links INNER JOIN test_options ON test_option_links.option_id = test_options.id " & _
                                                "     WHERE (((test_option_links.tp_id)=" & exp & "));" _
                                                ) _
                                        )
    ProcedureStepDataFactory = rs
End Function

Public Function RecordFactory(exp)
    Set cn = New connection
    Set rs = cn.connection.get_records(exp)
    Set RecordFactory = rs
End Function

Public Function create_temp_text()
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set tfolder = fso.GetSpecialFolder(2)
   tname = fso.GetTempName
   Set tfile = tfolder.CreateTextFile(tname)
   create_temp_text = tname
End Function

Public Function write_text(filename, text, iomode)
    Set object_file = CreateObject("Scripting.FileSystemObject").openTextFile(filename, iomode, True)
    object_file.WriteLine text
    object_file.Close
    Set object_file = Nothing
End Function

Function columnLetter(ByVal columnNumber)
				Dim n
				Dim C
				Dim s
				s = ""
				n = columnNumber
				Do
								C = ((n - 1) Mod 26)
								s = Chr(C + 65) & s
								n = (n - C) \ 26
				Loop While n > 0
				columnLetter = s
End Function

Function RandomString(ByVal strLen)
				Dim str, min, max

				Const LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789"
				min = 1
				max = Len(LETTERS)

				Randomize
				For i = 1 To strLen
								str = str & Mid(LETTERS, Int((max - min + 1) * Rnd + min), 1)
				Next
				RandomString = str
End Function

Public Function IsNullOrEmpty(value)
				out = True
				If IsObject(value) Then
								out = False
				Else
								If IsEmpty(value) Then
												out = True
								ElseIf value = "" Then
												out = True
								ElseIf IsNull(value) Then
												out = True
								Else
												out = False
								End If
				End If
				IsNullOrEmpty = out
End Function

Function reg_replace(original_text, find_text, replace_with)
        Set regex = CreateObject("VBScript.RegExp")
        regex_pattern = find_text
        regex.Global = True
        regex.IgnoreCase = True
        regex.Pattern = regex_pattern
        If regex.Test(original_text) Then
            new_string = regex.Replace(original_text, replace_with)
        Else
            new_string = original_text
        End If
        reg_replace = new_string
    End Function

    Function reg_match(original_text, find_text)
        Set regex = CreateObject("VBScript.RegExp")
        regex_pattern = find_text
        regex.Global = True
        regex.IgnoreCase = True
        regex.Pattern = regex_pattern
        Set reg_match = regex.execute(original_text)
    End Function

    Function reg_match_first(original_text, find_text)
        Set regex = CreateObject("VBScript.RegExp")
        regex_pattern = find_text
        regex.Global = True
        regex.IgnoreCase = True
        regex.Pattern = regex_pattern
        If regex.Test(original_text) Then
                reg_matched = regex.execute(original_text)(0).value
        Else
                reg_matched = Null
        End If
        reg_match_first = reg_matched
    End Function

    Function db_remove_delimeter(text)
        db_remove_delimeter = Replace(text, "|", "~!!~")
    End Function

    Function db_restore_delimeter(text)
        db_restore_delimeter = Replace(text, "~!!~", "|")
    End Function

    Function db_qoutes(text)
        db_qoutes = Replace(text, "'", "''")
    End Function

    Function format_text_to_param(ByVal theString)
        Dim strAlphaNumeric, iCounter, cleanedString, strChar
        cleanedString = ""
        strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_" 'Used to check for numeric characters.
        For iCounter = 1 To Len(theString)
            strChar = Mid(theString, iCounter, 1)
            If InStr(strAlphaNumeric, strChar) Then
                cleanedString = cleanedString & strChar
            Else
                cleanedString = cleanedString & "_"
            End If
        Next
        format_text_to_param = cleanedString
    End Function
