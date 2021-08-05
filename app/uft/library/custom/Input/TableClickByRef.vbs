Function TableClickByRef(e) ' [KEYWORD=TRUE] '
	Dim objObject, intRow, intCheckCol, strTargetColumn, strChildClass, intChildIndex, strFindValue
	Err.Clear: On Error GoTo 0
	On Error Resume Next
	Set s = Evaluate(e)
	If s("run") = 1 Then
		Set stepReference = s("obj")
		stepData = e("data_in") 
		If Trim(UCase(stepData)) <> "%SKIP%" Then
			tmpValues = Split(stepData, "~!!!~")
			intRow = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(0))
			intCheckCol = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(1))
			strTargetColumn = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(2))
			Select Case stepReference.GetTOProperty("micclass")
			Case "WebTable"
				strChildClass = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(3))
				intChildIndex = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(4))
				strFindValue = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(5))
			Case "JavaTable"
				strChildClass = ""
				intChildIndex = 0
				strFindValue = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(5))
			End Select
			If Trim(intRow) = "" Or (Trim(intRow) = "%skip%" And Trim(intRow) = "%blank%") Then intRow = 0
			If Trim(intCheckCol) = "" Or (Trim(intCheckCol) = "%skip%" And Trim(intCheckCol) = "%blank%") Then intCheckCol = 0
			If Trim(strTargetColumn) = "" Or (Trim(strTargetColumn) = "%skip%" And Trim(strTargetColumn) = "%blank%") Then strTargetColumn = 0
			Set objObject = stepReference
			If objObject.Exist(0) Then
				If Trim(strFindValue) <> "" And (Trim(strFindValue) <> "%skip%" And Trim(strFindValue) <> "%blank%") Then
					If objObject.GetROProperty("cols") > 1 Then
						For iCounter = 0 To objObject.GetROProperty("rows")
							On Error Resume Next
							Select Case stepReference.GetTOProperty("micclass")
							Case "WebTable"
								If Trim(UCase(objObject.GetCellData(iCounter, intCheckCol))) = Trim(UCase(strFindValue)) Then
									objObject.ChildItem(iCounter, strTargetColumn, strChildClass, intChildIndex).highlight
									MouseLClick objObject.ChildItem(iCounter, strTargetColumn, strChildClass, intChildIndex), "", "", "", ""
									Exit For
								End If
							Case "JavaTable"
								If Trim(UCase(objObject.GetCellData(iCounter, intCheckCol))) = Trim(UCase(strFindValue)) Then
									objObject.Highlight
									objObject.ClickCell iCounter, strTargetColumn
									Exit For
								End If
							End Select
						Next
					Else
						tmpTabArr = myPactera_commonFunction_extractInnerHTMLTable(objObject.GetROProperty("innerhtml"))
						For iCounter = 0 To UBound(tmpTabArr, 1)
							If Trim(UCase(tmpTabArr(iCounter, strTargetColumn))) = Trim(UCase(strFindValue)) Then
								Set tmpObj = Browser("CreationTime:=" & myPactera_commonFunction_getActiveBrowserIndex).Page("index:=0").WebElement("visible:=True", "index:=0", "innertext:=" & tmpTabArr(iCounter, strTargetColumn))
								If tmpObj.Exist(0) Then
									On Error Resume Next
									tmpObj.Click
									Exit For
								Else
									For jCounter = 0 To UBound(tmpTabArr, 1)
										If Trim(UCase(tmpTabArr(jCounter, intCheckCol))) = Trim(UCase(strFindValue)) Or InStr(1, tmpTabArr(jCounter, intCheckCol), ">" & Trim(UCase(strFindValue)) & "<") <> 0 Then
											Execute "Set tmpObj = Browser(""CreationTime:="" & myPactera_commonFunction_getActiveBrowserIndex).Page(""index:=0"")." & strChildClass & "(""visible:=True"", ""index:=" & jCounter + intChildIndex & """)"
											If tmpObj.Exist(0) Then
												On Error Resume Next
												tmpObj.Click
												Exit For
											End If
										End If
									Next
									Exit For
								End If
							End If
						Next
					End If
				Else
					If objObject.GetROProperty("cols") > 1 Then
						On Error Resume Next
						Select Case stepReference.GetTOProperty("micclass")
						Case "WebTable"
							objObject.ChildItem(intRow, strTargetColumn, strChildClass, intChildIndex).highlight
							MouseLClick objObject.ChildItem(intRow, strTargetColumn, strChildClass, intChildIndex), "", "", "", ""
						Case "JavaTable"
							objObject.ClickCell intRow, strTargetColumn
						End Select
					Else
						tmpTabArr = myPactera_commonFunction_extractInnerHTMLTable(objObject.GetROProperty("innerhtml"))
						Set tmpObj = Browser("CreationTime:=" & myPactera_commonFunction_getActiveBrowserIndex).Page("index:=0").WebElement("visible:=True", "index:=0", "innertext:=" & tmpTabArr(intRow, strTargetColumn))
						If tmpObj.Exist(ObjectWaitTime) Then
							On Error Resume Next
							tmpObj.Click
						End If
					End If
				End If
			End If
		End If
	End If
	Set TableClickByRef = EvaluateExecution(Err, "")
	Err.Clear: On Error GoTo 0
End Function