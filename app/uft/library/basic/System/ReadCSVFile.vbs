Function ReadCSVFile(e) ' [KEYWORD=TRUE] '@L4	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		Set stepReference = s("obj")		stepData = e("data_in") 	End If	Set ReadCSVFile = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function