Function VerifyApplicationText(e) ' [KEYWORD=TRUE] '@L4	Dim intNumScrollDown, GetTextLocation_Left, GetTextLocation_Top, GetTextLocation_Right, GetTextLocation_Bottom, iCounter	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		Set stepReference = s("obj")		stepData = e("data_in") 		tmpValues = Split(stepData, "~!!!~")		strCheckText = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(0))		intNumScrollDown = CInt(myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(1)))		bolCloseWindow = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(2))		If stepReference.Exist(ObjectWaitTime) Then			On Error Resume Next			stepReference.Maximize			stepReference.Highlight			Set WShell = CreateObject("WScript.Shell")			If CInt(intNumScrollDown) = 0  Then intNumScrollDown = 1			For iCounter = 1 To intNumScrollDown + 1				captureScreen ""				stepReference.GetTextLocation strCheckText, GetTextLocation_Left, GetTextLocation_Top, GetTextLocation_Right, GetTextLocation_Bottom				If CInt(GetTextLocation_Left) > 0 Or CInt(GetTextLocation_Top) > 0 Or CInt(GetTextLocation_Right) > 0 Or CInt(GetTextLocation_Bottom) > 0 Then					If Trim(UCase(bolCloseWindow)) = "Y" Then stepReference.Close					Exit For				End If				WShell.SendKeys "{PGDN}"				Wait (1)			Next		End If	End If	Set VerifyApplicationText = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function