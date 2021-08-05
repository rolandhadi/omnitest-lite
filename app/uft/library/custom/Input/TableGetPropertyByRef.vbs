Function TableGetPropertyByRef(e) ' [KEYWORD=TRUE] '
	Dim objObject, outputValue, intRow, intCheckCol, strTargetColumn, strChildClass, intChildIndex, strFindValue, strPropertyToGet, strTarget
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
			strChildClass = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(3))
			intChildIndex = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(4))
			strFindValue = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(5))
			strPropertyToGet = myPactera_commonFunction_fixFuntionReferenceParameters(tmpValues(6))
			If Trim(intRow) = "" Or (Trim(intRow) = "%skip%" And Trim(intRow) = "%blank%") Then intRow = 0
			If Trim(intCheckCol) = "" Or (Trim(intCheckCol) = "%skip%" And Trim(intCheckCol) = "%blank%") Then intCheckCol = 0
			If Trim(strTargetColumn) = "" Or (Trim(strTargetColumn) = "%skip%" And Trim(strTargetColumn) = "%blank%") Then strTargetColumn = 0
			Set objObject = stepReference
			If objObject.Exist(0) Then
				If Trim(strFindValue) <> "" And (Trim(strFindValue) <> "%skip%" And Trim(strFindValue) <> "%blank%") Then
					If objObject.GetROProperty("cols") > 1 Then
						For iCounter = 0 To objObject.RowCount
							If Trim(UCase(objObject.GetCellData(iCounter, intCheckCol))) = Trim(UCase(strFindValue)) Then
								If InStr(1, UCase(strChildClass), "WEBELEMENT") <> 0 Then
									Set tmpObj = Browser("CreationTime:=" & myPactera_commonFunction_getActiveBrowserIndex).Page("index:=0").WebElement("index:=0", "innertext:=" & objObject.GetCellData(iCounter, strTargetColumn))
									If tmpObj.Exist(0) Then
										On Error Resume Next
										tmpValue = tmpObj.GetROProperty(strPropertyToGet)
										outputValue = tmpValue
										Exit For
									End If
								Else
									On Error Resume Next
									objObject.ChildItem(iCounter, strTargetColumn, strChildClass, intChildIndex).Highlight
									tmpValue = objObject.ChildItem(iCounter, strTargetColumn, strChildClass, intChildIndex).GetROProperty(strPropertyToGet)
									outputValue = tmpValue
									Exit For
								End If
							End If
						Next
					Else
						tmpTabArr = myPactera_commonFunction_extractInnerHTMLTable(objObject.GetROProperty("innerhtml"))
						For iCounter = 0 To UBound(tmpTabArr, 1)
							If Trim(UCase(tmpTabArr(iCounter, intCheckCol))) = Trim(UCase(strFindValue)) Then
								tmpValue = tmpTabArr(iCounter, strTargetColumn)
								outputValue = tmpValue
								Exit For
							End If
						Next
					End If
				Else
					If objObject.GetROProperty("cols") > 1 Then
						If InStr(1, UCase(strChildClass), "WEBELEMENT") <> 0 Then
							Set tmpObj = Browser("CreationTime:=" & myPactera_commonFunction_getActiveBrowserIndex).Page("index:=0").WebElement("index:=0", "innertext:=" & objObject.GetCellData(intRow, strTargetColumn))
							If tmpObj.Exist(0) Then
								On Error Resume Next
								tmpValue = tmpObj.GetROProperty(strPropertyToGet)
								outputValue = tmpValue
							End If
						Else
							On Error Resume Next
							objObject.ChildItem(intRow, strTargetColumn, strChildClass, intChildIndex).Highlight
							If Trim(strChildClass) <> "" Then
								objObject.ChildItem(intRow, strTargetColumn, strChildClass, intChildIndex).Highlight
								tmpValue = objObject.ChildItem(intRow, strTargetColumn, strChildClass, intChildIndex).GetROProperty(strPropertyToGet)
							Else
								tmpValue = objObject.GetCellData(intRow, strTargetColumn)
							End If
							outputValue = tmpValue
						End If
					Else
						tmpTabArr = myPactera_commonFunction_extractInnerHTMLTable(objObject.GetROProperty("innerhtml"))
						tmpValue = tmpTabArr(intRow, strTargetColumn)
						outputValue = tmpValue
					End If
				End If
			End If
			omnilite_update_data	outputValue, e("data_out")
		End If
	End If
	Set TableGetPropertyByRef = EvaluateExecution(Err, "")
	Err.Clear: On Error GoTo 0
End Function

