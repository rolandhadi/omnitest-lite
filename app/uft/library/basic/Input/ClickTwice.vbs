Function ClickTwice(e) ' [KEYWORD=TRUE] '@L1S1 @L2S1	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		Set stepReference = s("obj")		If IsObject(stepReference) Then			stepReference.Click			mobileClientCaptureFromMobile = False		Else			helper_function__findSwipe stepReference, stepOptions			mobileClientInstance.Click "NATIVE", stepReference, 0, 1			Wait (1)			If mobileClientInstance.IsElementFound ("NATIVE", stepReference, 0) = "True" Then				mobileClientInstance.Click "NATIVE", stepReference, 0, 1			End If			MOBILE_LAST_OBJECT_CLICKED_01 = stepReference			mobileClientCaptureFromMobile = True		End If	End If	Set ClickTwice = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function