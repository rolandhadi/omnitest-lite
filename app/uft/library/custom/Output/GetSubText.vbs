Function GetSubText(e) ' [KEYWORD=TRUE] '@L4	Dim tmpValues, strText, strLeftBound, strRighBound, outValue	Err.Clear: On Error GoTo 0	On Error Resume Next	Set s = Evaluate(e)	If s("run") = 1 Then		Set stepReference = s("obj")		stepData = e("data_in")		tmpValues = Split(stepData, "~!!!~")		If IsObject(stepReference) Then			strText = stepReference.GetROProperty(myPactera_commonFunction_getObjectTextValue(stepReference))		Else			strText = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(0))		End If		strLeftBound = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(1))		strRighBound = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(2))		If strLeftBound = "%skip%" Then			strLeftBound = ""		End If		If strRighBound = "%skip%" Then			strRighBound = ""		End If		If strRighBound <> "" Then			outValue = Trim((myPactera_commonFunction_getSubString(strText, strLeftBound, strRighBound)))		Else			outValue = Trim(Mid(strText, InStr(strText,strLeftBound) + Len(strLeftBound)))		End If		Execute "temporaryContainer = outValue"		omnilite_update_data	"VAR[temporaryContainer]", e("data_out")	End If	Set GetSubText = EvaluateExecution(Err, "")	Err.Clear: On Error GoTo 0End Function