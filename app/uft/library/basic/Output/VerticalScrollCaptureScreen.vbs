Function VerticalScrollCaptureScreen(e) ' [KEYWORD=TRUE] '@L4	Dim iCounter, wShell	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		Set stepReference = s("obj")		stepData = CInt(myPactera_commonFunction_convertToText(stepData))		If stepReference.Exist(GlobalShortWait) Then			iCounter = 0			stepReference.Highlight			stepReference.Click			captureScreen ""			Set wShell = CreateObject("WScript.Shell")			If CInt(stepData) = 0 Then stepData = 1			For iCounter = 1 To stepData				wShell.SendKeys "{PGDN}"				Wait (1)				captureScreen ""			Next		End If	End If	Err.Clear: On Error GoTo 0	Set VerticalScrollCaptureScreen = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function