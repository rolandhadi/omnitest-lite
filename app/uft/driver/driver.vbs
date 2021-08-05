Public LOG_FOLDER
Public SCREENSHOT_FOLDER
Public FILES_FOLDER
Public DYNAMIC_OBJECTS_LIST
Public TEST_DATA_ITERATION
Public stPass
Public stFail
Public stDone
Public stSkip

Set DYNAMIC_OBJECTS_LIST = CreateObject("Scripting.Dictionary")

stPass = "Passed"
stFail = "Failed"
stDone = "Done"
stSkip = "Skipped"

For Each f In CreateObject("Scripting.FileSystemObject").GetFolder(ROOT_FOLDER & "app\uft\core\").Files
	ExecuteFile  f.Path
Next

omnilite_init

For Each f in read_function_libraries
    ExecuteFile f(0)
Next

Sub start_plan_execution(pe_id)

	ess = ExecutionScenarioDataFactory(pe_id)

	Set fso = CreateObject("Scripting.FileSystemObject")
	If NOT fso.FolderExists(ROOT_FOLDER & "result") Then fso.CreateFolder(ROOT_FOLDER & "result")
	LOG_FOLDER = ROOT_FOLDER & "result\" & get_timestamp("_") & "_" & ess(0).data("pe_name")
	If NOT fso.FolderExists(LOG_FOLDER) Then fso.CreateFolder(LOG_FOLDER)
	LOG_FOLDER = LOG_FOLDER & "\"
	SCREENSHOT_FOLDER = LOG_FOLDER & "\Screenshots"
	If NOT fso.FolderExists(SCREENSHOT_FOLDER) Then fso.CreateFolder(SCREENSHOT_FOLDER)
	SCREENSHOT_FOLDER = SCREENSHOT_FOLDER & "\"
	FILES_FOLDER = LOG_FOLDER & "\Files"
	If NOT fso.FolderExists(FILES_FOLDER) Then fso.CreateFolder(FILES_FOLDER)
	FILES_FOLDER = FILES_FOLDER & "\"

	append_result_summary "pe_name:=" & ess(0).data("pe_name")
	append_result_summary "pe_status:=3"
	append_result_summary "pe_start:=" & format_now

	append_execution_result Chr(34) & "Test Scenario" & Chr(34) & "," & _
	 												Chr(34) & "TC Order" & Chr(34) & "," & _
													Chr(34) & "Test Case" & Chr(34) & "," & _
													Chr(34) & "TP Order" & Chr(34) & "," & _
													Chr(34) & "Test Procedure" & Chr(34) & "," & _
													Chr(34) & "Step Order" & Chr(34) & "," & _
													Chr(34) & "Step" & Chr(34) & "," & _
													Chr(34) & "Start" & Chr(34) & "," & _
													Chr(34) & "End" & Chr(34) & "," & _
													Chr(34) & "Status" & Chr(34) & "," & _
													Chr(34) & "Remarks" & Chr(34)

	If Not ess(0).eof Then
	  es_ids = ess(0).column("es_id", Array("es_id", "*"))
	  cs_ids = ess(0).column("cs_id", Array("cs_id", "*"))
	  cp_ids = ess(0).column("cp_id", Array("cp_id", "*"))
	  Set test_procedures = CreateObject("Scripting.Dictionary")
	  For Each cp_id In cp_ids
	    tp_steps = ess(0).column("ps_order", Array("cp_id", cp_id))
	    Set test_procedures(cp_id) = CreateObject("Scripting.Dictionary")
	    tp_row = ess(0).find("tp_name", Array("cp_id", cp_id), True)
	    test_procedures(cp_id)("tp_name") = tp_row(0)
	    Set test_procedures(cp_id)("steps") = CreateObject("Scripting.Dictionary")
	    For Each tp_step in tp_steps
	      ps_row =  ess(0).find(Array("ps_id", "keyword_name", "to_id", "to_name"), _
	                              Array( _
	                                Array("cp_id", cp_id), _
	                                Array("ps_order", tp_step) _
	                              )_
	                            , True _
	                          )
	      Set test_procedures(cp_id)("steps")(tp_step) = CreateObject("Scripting.Dictionary")
	      test_procedures(cp_id)("steps")(tp_step)("ps_id") = ps_row(0)(0)
	      test_procedures(cp_id)("steps")(tp_step)("keyword_name") = ps_row(1)(0)
	      test_procedures(cp_id)("steps")(tp_step)("to_id") = ps_row(2)(0)
	      test_procedures(cp_id)("steps")(tp_step)("to_name") = ps_row(3)(0)
	    Next
	  Next

	  Set test_cases = CreateObject("Scripting.Dictionary")
	  For Each cs_id In cs_ids
	    tc_procedures = ess(0).column("cp_order", Array("cs_id", cs_id))
	    Set test_cases(cs_id) = CreateObject("Scripting.Dictionary")
	    tc_row = ess(0).find("tc_name", Array("cs_id", cs_id), True)
	    test_cases(cs_id)("tc_name") = tc_row(0)
	    Set test_cases(cs_id)("test_procedures") = CreateObject("Scripting.Dictionary")
	    For Each tc_procedure in tc_procedures
	      cp_row =  ess(0).find(Array("cp_order", "cp_id", "tp_id"), _
	                              Array( _
	                                Array("cs_id", cs_id), _
	                                Array("cp_order", tc_procedure) _
	                              )_
	                            , True _
	                          )
	      Set test_cases(cs_id)("test_procedures")(cp_row(1)(0)) = CreateObject("Scripting.Dictionary")
	      test_cases(cs_id)("test_procedures")(cp_row(1)(0))("cp_order") = cp_row(0)(0)
	      test_cases(cs_id)("test_procedures")(cp_row(1)(0))("cp_id") = cp_row(1)(0)
	      test_cases(cs_id)("test_procedures")(cp_row(1)(0))("tp_id") = cp_row(2)(0)
	    Next
	  Next

	  Set test_scenarios = CreateObject("Scripting.Dictionary")
	  es_order = 0
	  For Each es_id In es_ids
	    es_order = es_order + 1
	    tc_cases = ess(0).column("cs_order", Array("es_id", es_id))
	    Set test_scenarios(es_id) = CreateObject("Scripting.Dictionary")
	    ts_row = ess(0).find(Array("ts_name", "iteration", "run"), Array("es_id", es_id), True)
	    test_scenarios(es_id)("es_order") = es_order
	    test_scenarios(es_id)("ts_name") = ts_row(0)(0)
	    test_scenarios(es_id)("iteration") = ts_row(1)(0)
			test_scenarios(es_id)("run") = ts_row(2)(0)
	    Set test_scenarios(es_id)("test_cases") = CreateObject("Scripting.Dictionary")
	    For Each ts_case in tc_cases
	      cs_row =  ess(0).find(Array("cs_order", "cs_id", "tc_id"), _
	                              Array( _
	                                Array("es_id", es_id), _
	                                Array("cs_order", ts_case) _
	                              ) _
	                            , True _
	                          )
	      Set test_scenarios(es_id)("test_cases")(cs_row(1)(0)) = CreateObject("Scripting.Dictionary")
	      test_scenarios(es_id)("test_cases")(cs_row(1)(0))("cs_order") = cs_row(0)(0)
	      test_scenarios(es_id)("test_cases")(cs_row(1)(0))("cs_id") = cs_row(1)(0)
	      test_scenarios(es_id)("test_cases")(cs_row(1)(0))("tc_id") = cs_row(2)(0)
	    Next
	  Next
	End If

	execution_failed = False

	For Each ts In test_scenarios
	  If execution_failed = True Then Exit For
	  append_result_summary "ts_count:=+1"
	  For tx = 1 To test_scenarios(ts)("iteration")
			TEST_DATA_ITERATION = tx
	    If execution_failed = True Then
				append_result_summary "ts_" & test_scenarios(ts)("es_order") & "_iteration:=+1"
				append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_name:=" & test_scenarios(ts)("ts_name")
				append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_status:=4"
			Else
				append_result_summary "ts_" & test_scenarios(ts)("es_order") & "_iteration:=+1"
				append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_name:=" & test_scenarios(ts)("ts_name")
				append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_status:=3"
	    End If
			If test_scenarios(ts)("run") = 1 Then
		    For Each tc In test_scenarios(ts)("test_cases")
		      If execution_failed = True Then
						append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_count:=+1"
						append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" & test_scenarios(ts)("test_cases")(tc)("cs_order") & "_name:=" & test_cases(tc)("tc_name")
						append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" & test_scenarios(ts)("test_cases")(tc)("cs_order") & "_status:=4"
					Else
						append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_count:=+1"
						append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" & test_scenarios(ts)("test_cases")(tc)("cs_order") & "_name:=" & test_cases(tc)("tc_name")
						append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" & test_scenarios(ts)("test_cases")(tc)("cs_order") & "_status:=3"
		      End If
		      For Each tp In test_cases(tc)("test_procedures")
		        If execution_failed = True Then
							append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_count:=+1"
							append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_" & test_cases(tc)("test_procedures")(tp)("cp_order") & "_name:=" & test_procedures(tp)("tp_name")
							append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_" & test_cases(tc)("test_procedures")(tp)("cp_order") & "_status:=4"
						Else
							append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_count:=+1"
							append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_" & test_cases(tc)("test_procedures")(tp)("cp_order") & "_name:=" & test_procedures(tp)("tp_name")
							append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_" & test_cases(tc)("test_procedures")(tp)("cp_order") & "_status:=3"

							For Each st In test_procedures(tp)("steps")

								If execution_failed = True Then
			            Exit For
			          End If

			          current_data_name = ess(1).find("data_name", Array	( _
												                                              Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
												                                              Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
												                                              Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
												                                            ), True)(0)

			          current_data_value_ref_id = ess(1).find("data_value_ref_id", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

			          current_ref_name = ess(1).find("ref_name", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

			          current_ref_value = ess(1).find("ref_value", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

								current_ref_value = get_data_reference_value(current_ref_value, tx)

			          current_data_value_id_in = ess(1).find("data_value_id_in", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

			          current_data_value_in = ess(1).find("data_value_in", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

			          current_data_value_id_out = ess(1).find("data_value_id_out", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

			          current_data_value_out = ess(1).find("data_value_out", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

								current_run_flag = ess(1).find("run", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

								current_screenshot_flag = ess(1).find("screenshot", Array	( _
																												Array("cs_id", test_scenarios(ts)("test_cases")(tc)("cs_id")), _
																												Array("cp_id", test_cases(tc)("test_procedures")(tp)("cp_id")), _
																												Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")) _
			                                            		), True)(0)

								If current_run_flag = "" Then current_run_flag = 1
								If current_screenshot_flag = "" Then current_screenshot_flag = 1

			          current_options = ess(2).find(Array("name", "item"), Array("ps_id", test_procedures(tp)("steps")(st)("ps_id")), False)

			          Set step_evaluation_data = CreateObject("scripting.dictionary")

			        	step_evaluation_data("keyword_name") = test_procedures(tp)("steps")(st)("keyword_name")

			        	If Trim(test_procedures(tp)("steps")(st)("to_id")) <> "" Then
			        		Set step_object = TestObjectFactory(test_procedures(tp)("steps")(st)("to_id"))
			        		If step_object.count > 0 Then
			        			step_evaluation_data("obj") = step_object.first.test_object_value("")
			        		Else
			        			step_evaluation_data("obj") = ""
			        		End If
			        		'object_type_code = "WEB"
			        		'step_evaluation_data("obj") = ess(3).find("item", Array	( _
			        		'																						Array("to_id", ess(0).data("to_id")), _
			        		'																						Array("type_code", object_type_code) _
			        		'																					), True)(0)
			        	Else
			              step_evaluation_data("obj") = ""
			        	End If

			        	step_evaluation_data("options") = current_options

			        	If Trim(current_data_value_ref_id) <> "" Then
			         		step_evaluation_data("data_in") = current_ref_value
			        	ElseIf Trim(current_data_value_id_in) <> "" Then
			        	 	Set iteration_data = TestDataFactory(current_data_value_id_in)
			        		If iteration_data.count > 0 Then
			        			step_evaluation_data("data_in") = iteration_data.first.iteration(tx)
			        		Else
			        			step_evaluation_data("data_in") = ""
			        		End If
			        	Else
			        		step_evaluation_data("data_in") = current_data_value_in
			        	End If

			        	If Trim(current_data_value_id_out) <> "" Then
			         		step_evaluation_data("data_out") = current_data_value_id_out
			        	ElseIf Trim(current_data_value_out) <> "" Then
			        		step_evaluation_data("data_out") = current_data_value_out
			        	Else
			        		step_evaluation_data("data_out") = ""
			        	End If

			        	On Error Resume Next
			        	If IsArray(current_options) Then
			        		option_list = ""
			        		For option_index = 0 To UBound(current_options)
			        				option_list = option_list & current_options(0)(option_index) & "(" & Chr(34) & current_options(1)(option_index) & Chr(34) & ")" & " "
			        		Next
			        	End If
								Err.Clear: On Error Goto 0
								On Error Resume Next

			          Print test_scenarios(ts)("ts_name") & " - " & _
			                test_cases(tc)("tc_name") & " - " & _
			                test_procedures(tp)("tp_name") & " - " & _
			                step_evaluation_data("keyword_name") & " - " & _
			                step_evaluation_data("obj") & " - " & _
			                step_evaluation_data("data_in") & " - " & _
			                step_evaluation_data("data_out") & " - " & _
			                option_list

								If current_run_flag = 1 Then

				          'Set step_execution_status = CreateObject("Scripting.Dictionary")
				          'step_execution_status("status") = 1
									step_execution_start = Now
				          Set step_execution_status = Eval(step_evaluation_data("keyword_name") & " (step_evaluation_data)")
									step_execution_end = Now
				        	If Err.Number <> 0 Or step_execution_status("status") = 0 Then
										append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_status:=0"
										append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" & test_scenarios(ts)("test_cases")(tc)("cs_order") & "_status:=0"
										append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_" & test_cases(tc)("test_procedures")(tp)("cp_order") & "_status:=0"
				            append_execution_result Chr(34) & test_scenarios(ts)("ts_name") & Chr(34) & "," & _
				          													Chr(34) & test_scenarios(ts)("test_cases")(tc)("cs_order") & Chr(34) & "," & _
				          													Chr(34) & test_cases(tc)("tc_name") & Chr(34) & "," & _
				          													Chr(34) & test_cases(tc)("test_procedures")(tp)("cp_order") & Chr(34) & "," & _
				          													Chr(34) & test_procedures(tp)("tp_name") & Chr(34) & "," & _
				          													Chr(34) & st & Chr(34) & "," & _
				          													Chr(34) & step_evaluation_data("keyword_name") & Chr(34) & "," & _
				          													Chr(34) & step_execution_start & Chr(34) & "," & _
				          													Chr(34) & step_execution_end & Chr(34) & "," & _
				          													Chr(34) & "Failed" & Chr(34) & "," & _
				          													Chr(34) & step_execution_status("err_log") & Chr(34)
				            execution_failed = True
										Err.Clear: On Error Goto 0
				        		Exit For
				        	Else
				            append_execution_result Chr(34) & test_scenarios(ts)("ts_name") & Chr(34) & "," & _
																						Chr(34) & test_scenarios(ts)("test_cases")(tc)("cs_order") & Chr(34) & "," & _
																						Chr(34) & test_cases(tc)("tc_name") & Chr(34) & "," & _
																						Chr(34) & test_cases(tc)("test_procedures")(tp)("cp_order") & Chr(34) & "," & _
																						Chr(34) & test_procedures(tp)("tp_name") & Chr(34) & "," & _
																						Chr(34) & st & Chr(34) & "," & _
																						Chr(34) & step_evaluation_data("keyword_name") & Chr(34) & "," & _
				          													Chr(34) & step_execution_start & Chr(34) & "," & _
				          													Chr(34) & step_execution_end & Chr(34) & "," & _
				          													Chr(34) & "Passed" & Chr(34) & "," & _
				          													Chr(34) & "Remarks" & Chr(34)
				        	End If

									If current_screenshot_flag = 1 Then
										Desktop.CaptureBitmap SCREENSHOT_FOLDER & "ts_" & LPad(test_scenarios(ts)("es_order"), 3) & "-" & LPad(tx, 3) & "_tc_" &  LPad(test_scenarios(ts)("test_cases")(tc)("cs_order"), 3) & "_tp_" & LPad(test_cases(tc)("test_procedures")(tp)("cp_order"), 3) & "_" & LPad(st, 4) & ".png", True
									End If
				        	Err.Clear: On Error Goto 0

								Else

									append_execution_result Chr(34) & test_scenarios(ts)("ts_name") & Chr(34) & "," & _
																					Chr(34) & test_scenarios(ts)("test_cases")(tc)("cs_order") & Chr(34) & "," & _
																					Chr(34) & test_cases(tc)("tc_name") & Chr(34) & "," & _
																					Chr(34) & test_cases(tc)("test_procedures")(tp)("cp_order") & Chr(34) & "," & _
																					Chr(34) & test_procedures(tp)("tp_name") & Chr(34) & "," & _
																					Chr(34) & st & Chr(34) & "," & _
																					Chr(34) & step_evaluation_data("keyword_name") & Chr(34) & "," & _
																					Chr(34) & Now & Chr(34) & "," & _
																					Chr(34) & Now & Chr(34) & "," & _
																					Chr(34) & "Skipped" & Chr(34) & "," & _
																					Chr(34) & "Remarks" & Chr(34)

								End If

			        Next

						End If
						If execution_failed = False Then
		        	append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" &  test_scenarios(ts)("test_cases")(tc)("cs_order") & "_tp_" & test_cases(tc)("test_procedures")(tp)("cp_order") & "_status:=1"
						End If
					Next
					If execution_failed = False Then
						append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_tc_" & test_scenarios(ts)("test_cases")(tc)("cs_order") & "_status:=1"
					End If
				Next
				If execution_failed = False Then
		    	append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_status:=1"
				End If
			Else
				append_result_summary "ts_" & test_scenarios(ts)("es_order") & ":" & tx & "_status:=4"
			End If
		Next
	Next

	If execution_failed = True Then
	  append_result_summary "pe_status:=0"
	Else
	  append_result_summary "pe_status:=1"
	End If
	append_result_summary "pe_end:=" & format_now

End Sub

Function ExecuteFile_(filename)
	Print filename
   ExecuteGlobal CreateObject("Scripting.FileSystemObject").openTextFile(filename).readAll()
End Function

Function Print_(string_text)
   Set object_file = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Logs\" & SESSION_NAME & ".txt", 8, True)
	object_file.WriteLine string_text
	object_file.Close
	Set object_file = Nothing
End Function

Function Evaluate(e)

	On Error Resume Next
	Set o = CreateObject("scripting.dictionary")
	If Trim(UCase(e("data_in"))) = "[DEFAULT]" Then
		a = 0
	Else
		o("mobile") = 0
		If (e("obj") <> "") Then
			If IsObject(Eval(e("obj"))) Then
				If DYNAMIC_OBJECTS_LIST.exists("RT" & e("obj")) Then
					Set o("obj") = Eval(DYNAMIC_OBJECTS_LIST("RT" & o("ebj")))
				Else
					Set o("obj") = Eval(e("obj"))
				End If
			Else
				o("mobile") = 1
				If DYNAMIC_OBJECTS_LIST.exists("RT" & e("obj")) Then
					o("obj") = Eval(DYNAMIC_OBJECTS_LIST("RT" & e("obj")))
				Else
					o("obj") = Eval(e("obj"))
				End If
			End If
		Else
			Set o("obj") = Nothing
		End If
		a = 1
		If IsArray(e("options")) Then
   		For option_index = 0 To UBound(e("options"))
				If IsArray(e("options")(0)) Then
					a = a * Eval(e("options")(0)(option_index) & "(o(""obj""), " & Chr(34) & e("options")(1)(option_index) & Chr(34) & ")")
				End If
   		Next
		End If
	End If
	o("run") = a
	Set Evaluate = o
	Err.Clear

End Function

Function EvaluateExecution(err_log, output)

	Set o = CreateObject("scripting.dictionary")
	If err_log.Number <> 0 Then
		o("status") = 0
	Else
		o("status") = 1
	End If
	o("err_log") = err_log.Description
	o("output") = output
	Set EvaluateExecution = o

End Function

Function read_function_libraries()
	Set session_lib = New FunctionLibrary
 	session_lib.load_libraries ROOT_FOLDER & "app\uft\library\", "vbs", True
	read_function_libraries = session_lib.get_libraries
End Function

Function append_result_summary(text)
		result_filename = LOG_FOLDER & "summary.result"
		Set log_file = CreateObject("Scripting.FileSystemObject").OpenTextFile(result_filename, 8, True, False)
		log_file.WriteLine text
End Function

Function append_execution_result(text)
		result_filename = LOG_FOLDER & "test_result.csv"
		Set log_file = CreateObject("Scripting.FileSystemObject").OpenTextFile(result_filename, 8, True, False)
		log_file.WriteLine text
End Function

Function get_timestamp(delimiter)
	get_timestamp = LPad(Year(Now), 2) & delimiter & LPad(Month(Now), 2) & delimiter & LPad(Day(Now), 2) & delimiter & LPad(Hour(Now), 2) & delimiter & LPad(Minute(Now), 2) & delimiter & LPad(Second(Now), 2)
End Function

Function format_now()
	format_now = LPad(Year(Now), 2) & "-" & LPad(Month(Now), 2) & "-" & LPad(Day(Now), 2) & " " & LPad(Hour(Now), 2) & ":" & LPad(Minute(Now), 2) & ":" & LPad(Second(Now), 2)
End Function

Function LPad(v, l)
	If Len(v) > l Then l = Len(v)
	LPad = Right(String(l, "0") & v, l)
End Function

Function get_data_reference_value(data_value, iteration_number)
	data_value_out = data_value
	Set data_input_matches = reg_match(data_value, "INPUT\[(\w+)\]")
	For i = 0 To data_input_matches.count - 1
		data_identifier_name = data_input_matches(i).SubMatches(0)
		Set data_identifier = TestDataFactory(Array("name", "'" & data_identifier_name & "'"))
		If data_identifier.count > 0 Then
				data_value_out = reg_replace(data_value_out, Replace(Replace(data_input_matches(i).value, "[", "\["), "]", "\]"), data_identifier.first.iteration(iteration_number))
		End If
	Next

	Set data_input_matches = reg_match(iteration, "VAR\[(\w+)\]")
	For i = 0 To data_input_matches.count - 1
		data_identifier_name = data_input_matches(i).SubMatches(0)
		data_value_out = reg_replace(data_value_out, Replace(Replace(data_input_matches(i).value, "[", "\["), "]", "\]"), Eval(data_identifier_name))
	Next
	get_data_reference_value = data_value_out
End Function
