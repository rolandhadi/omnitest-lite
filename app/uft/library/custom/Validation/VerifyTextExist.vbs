Function VerifyTextExist(e) ' [KEYWORD=TRUE] '@L1S1 @L2S1	Dim objLink, objDesc, iCounter	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		If s("mobile") <> 1 Then			Set stepReference = s("obj")				Set objDesc = Description.Create()			objDesc("micclass").Value = "WebElement"			objDesc("innertext").Value = stepData			objDesc("visible").Value = True			Set objLink = stepReference.ChildObjects(objDesc)			For iCounter = 1 To Setting("DefaultTimeOut") \ 1000				If objLink.Count = 0 Then					Set objLink = stepReference.ChildObjects(objDesc)					Wait(1)				Else					Exit For				End If			Next			If objLink(0).Exist Then				lastCheckPointStatus = stPass				objLink(0).Highlight			Else				LastCheckPointMessage = "Text does not exist"				Err.Raise 424			End If			mobileClientCaptureFromMobile = False		Else			stepReference = s("obj")			helper_function__findSwipe stepReference, stepOptions			mobileClientInstance.Click "NATIVE", stepReference, 0, 1			mobileClientCaptureFromMobile = True		End If	End If	Set VerifyTextExist = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function