Function GetText(e) ' [KEYWORD=TRUE] '@L1S2 @L2S2	Dim outValue, strTarget, inputId	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		Set stepReference = s("obj")		If IsObject(stepReference) Then			outValue = stepReference.GetROProperty(myPactera_commonFunction_getObjectTextValue(stepReference))		Else			helper_function__findSwipe stepReference, stepOptions			outValue=mobileClientInstance.ElementGetText("NATIVE", stepReference, 0)			mobileClientCaptureFromMobile = True		End If		omnilite_update_data	outValue, e("data_out")	End If	Set GetText = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function