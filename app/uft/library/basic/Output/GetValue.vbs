Function GetValue(e) ' [KEYWORD=TRUE] '
	Dim outValue, strGetFrom, strTarget, inputId, targetID
	Err.Clear: On Error GoTo 0
	On Error Resume Next
	Set s = Evaluate(e)
	If s("run") = 1 Then
		Set stepReference = s("obj")
		stepData = e("data_in") 
		omnilite_update_data	stepData, e("data_out")
	End If
	Set GetValue = EvaluateExecution(Err, "")
	Err.Clear: On Error GoTo 0
End Function