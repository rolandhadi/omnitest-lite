Function TickCheckBox(e) ' [KEYWORD=TRUE] '@L1S1	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		Set stepReference = s("obj")		stepData = e("data_in") 		If stepData = "" Then stepData = "ON"		If IsObject(stepReference) Then			Select Case stepReference.GetTOProperty("micclass")			Case "JavaCheckBox"				If Trim(UCase(stepData)) <> "%SKIP%" Then stepReference.Set stepData			Case Else				If Trim(UCase(stepData)) <> "%SKIP%" Then stepReference.Set stepData			End Select		Else			If Trim(UCase(stepData)) = "ON" Or Trim(UCase(stepData)) = "YES" Or Trim(UCase(stepData)) = "Y" Then				stepReference = s("obj")				helper_function__findSwipe stepReference, stepOptions				mobileClientInstance.Click "NATIVE", stepReference, 0, 1				mobileClientCaptureFromMobile = True			End If		End If	End If	Set TickCheckBox = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function