Public stFail
Public stPass
Public stSkip

stFail = "Failed"
stPass = "Passed"
stSkip = "Skipped"

Function myPactera_commonFunction_getActiveBrowserIndex()
	Dim desc, children
	desc = Description.create
	desc("micclass").value = "Browser"
	children = Desktop.ChildObjects(desc)
	If children.Count > 0 Then
		m_activeBrowserIndex = children.Count - 1
		myPactera_commonFunction_getActiveBrowserIndex = m_activeBrowserIndex
	Else
		m_activeBrowserIndex = 0
		myPactera_commonFunction_getActiveBrowserIndex = m_activeBrowserIndex
	End If
	desc = Nothing
	children = Nothing
End Function

Function myPactera_commonFunction_fixFuntionReferenceParameters(ByVal strValue)
	If Not IsDBNull(strValue) Then
		myPactera_commonFunction_fixFuntionReferenceParameters = Replace(Replace(strValue, "\\q", ""), Chr(34), "")
	Else
		myPactera_commonFunction_fixFuntionReferenceParameters = ""
	End If
End Function

Function myPactera_commonFunction_checkTest(ByVal strOperation, ByVal strValue, ByVal strCheckValue)
	Set myPactera_commonFunction_checkTest = checkPropertyTest(strOperation, strValue, strCheckValue)
End Function

Function myPactera_commonFunction_GetScreenX()
	Dim objWMISer, ColItems, objItem
	Dim curResWidth, curResHeight
	On Error Resume Next
	curResWidth = 0
	Set objWMISer = GetObject("Winmgmts:\\.\root\cimv2")
	Set ColItems = objWMISer.ExecQuery("Select * FROM Win32_DesktopMonitor WHERE DeviceID = 'DesktopMonitor1'", , 0)
	For Each objItem In ColItems
					curResWidth = objItem.ScreenWidth
					curResHeight = objItem.ScreenHeight - 25
	Next
	If Err.Number <> 0 Then
					curResWidth = 800
					curResHeight = 600
	End If
	myPactera_commonFunction_GetScreenX = curResWidth
	Err.Clear()
	On Error GoTo 0
	objItem = Nothing
	ColItems = Nothing
	objWMISer = Nothing
End Function

Function myPactera_commonFunction_GetScreenY()
	Dim objWMISer, ColItems, objItem
	Dim curResWidth, curResHeight
	On Error Resume Next
	curResHeight = 0
	Set objWMISer = GetObject("Winmgmts:\\.\root\cimv2")
	Set ColItems = objWMISer.ExecQuery("Select * FROM Win32_DesktopMonitor WHERE DeviceID = 'DesktopMonitor1'", , 0)
	For Each objItem In ColItems
					curResWidth = objItem.ScreenWidth
					curResHeight = objItem.ScreenHeight - 25
	Next
	If Err.Number <> 0 Then
					curResWidth = 800
					curResHeight = 600
	End If
	myPactera_commonFunction_GetScreenY = curResHeight
	Err.Clear()
	On Error GoTo 0
	objItem = Nothing
	ColItems = Nothing
	objWMISer = Nothing
End Function

Function myPactera_commonFunction_extractObjectTable(ByVal stepReference)
	Dim iCounter, jCounter, tableDetails, countLimit, arrayTable
	tableDetails = New libPacteraExecution.clsSubContent
	tableDetails.r.Item("Rows") = stepReference.GetROProperty("rows")
	tableDetails.r.Item("Cols") = stepReference.GetROProperty("cols")
	Select Case stepReference.GetTOProperty("micclass")
					Case "JavaTable"
									countLimit = 1
					Case Else
									countLimit = 0
	End Select
	If stepReference.GetROProperty("cols") > 1 Then
					For iCounter = 1 To stepReference.GetROProperty("rows")
									For jCounter = 1 To stepReference.GetROProperty("cols")
													tableDetails.r.Item("R" & iCounter & "C" & jCounter) = stepReference.GetCellData(iCounter - countLimit, jCounter - countLimit)
									Next
					Next
	Else
					arrayTable = extractInnerHTMLTable(stepReference.GetROProperty("innerhtml"))
					For iCounter = 1 To UBound(arrayTable, 1) - 1
									For jCounter = 1 To UBound(arrayTable, 2) - 1
													tableDetails.r.Item("R" & iCounter & "C" & jCounter) = arrayTable(iCounter, jCounter)
									Next
					Next
	End If
	myPactera_commonFunction_extractObjectTable = tableDetails
End Function

Function myPactera_commonFunction_nextScreenShotFile()
	myPactera_commonFunction_nextScreenShotFile = 1
End Function

Function myPactera_commonFunction_getObjectTextValue(ByRef stepReference)
	Select Case UCase(stepReference.GetTOProperty("micclass"))
			Case "WEBLIST", "WINLIST", "WINLISTVIEW", "WEBCOMBOBOX", "WINCOMBOBOX", "JAVALIST"
				myPactera_commonFunction_getObjectTextValue = "all items"
			Case "WEBELEMENT", "WEBTABLE"
				myPactera_commonFunction_getObjectTextValue = "innertext"
			Case "JAVAOBJECT"
				myPactera_commonFunction_getObjectTextValue = "text"
			Case "JAVAINTERNALFRAME"
				myPactera_commonFunction_getObjectTextValue = "text"
			Case "SWFOBJECT", "SWFEDIT"
				myPactera_commonFunction_getObjectTextValue = "text"
			Case Else
				myPactera_commonFunction_getObjectTextValue = "value"
	End Select
End Function

Function checkPropertyTest(ByVal strOperation, ByVal strValue, ByVal strCheckValue)
		Dim tmpValue, tmpResult, tmpArray
		Set tmpResult = CreateObject("Scripting.Dictionary")
		tmpArray = Split(strCheckValue, "|")
		On Error Resume Next
		Select Case UCase(Trim(strOperation))
						'For String
						Case "EQUAL"
										For Each tmpValue In tmpArray
														If StrComp(Trim(strValue), Trim(tmpValue)) = 0 Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: EQUAL TO (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: EQUAL TO (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "EQUAL (IGNORE CASE)"
										For Each tmpValue In tmpArray
														If UCase(Trim(strValue)) = UCase(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: EQUAL (IGNORE CASE) TO (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: EQUAL (IGNORE CASE) TO (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NOT EQUAL"
										For Each tmpValue In tmpArray
														If StrComp(Trim(strValue), Trim(tmpValue)) <> 0 Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NOT EQUAL TO (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NOT EQUAL TO (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NOT EQUAL (IGNORE CASE)"
										For Each tmpValue In tmpArray
														If UCase(Trim(strValue)) <> UCase(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NOT EQUAL (IGNORE CASE) TO (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NOT EQUAL (IGNORE CASE) TO (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "CONTAIN"
										For Each tmpValue In tmpArray
														If InStr(1, Trim(strValue), Trim(tmpValue)) > 0 Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: CONTAIN (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: CONTAIN (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "CONTAIN (IGNORE CASE)"
										For Each tmpValue In tmpArray
														If InStr(1, UCase(Trim(strValue)), UCase(Trim(tmpValue))) > 0 Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: CONTAIN (IGNORE CASE) (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: CONTAIN (IGNORE CASE) (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NOT CONTAIN"
										For Each tmpValue In tmpArray
														If InStr(1, Trim(strValue), Trim(tmpValue)) = 0 Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NOT CONTAIN (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stPass
										tmpResult("Remarks") = "Expected: NOT CONTAIN (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NOT CONTAIN (IGNORE CASE)"
										For Each tmpValue In tmpArray
														If InStr(1, UCase(Trim(strValue)), UCase(Trim(tmpValue))) = 0 Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NOT CONTAIN (IGNORE CASE) (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NOT CONTAIN (IGNORE CASE) (" & strCheckValue & ") Actual: (" & strValue & ")"
										'For Number
						Case "NUMBER EQUAL"
										For Each tmpValue In tmpArray
														If Val(Trim(strValue)) = Val(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NUMBER EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NUMBER EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NUMBER NOT EQUAL"
										For Each tmpValue In tmpArray
														If Val(Trim(strValue)) <> Val(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NUMBER NOT EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NUMBER NOT EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NUMBER GREATER THAN"
										For Each tmpValue In tmpArray
														If Val(Trim(strValue)) > Val(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NUMBER GREATER THAN (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NUMBER GREATER THAN (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NUMBER GREATER THAN or EQUAL"
										For Each tmpValue In tmpArray
														If Val(Trim(strValue)) > Val(Trim(tmpValue)) Or Val(Trim(strValue)) = Val(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NUMBER GREATER THAN or EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NUMBER GREATER THAN or EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NUMBER LESS THAN"
										For Each tmpValue In tmpArray
														If Val(Trim(strValue)) < Val(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NUMBER LESS THAN (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NUMBER LESS THAN (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case "NUMBER LESS THAN OR EQUAL"
										For Each tmpValue In tmpArray
														If Val(Trim(strValue)) < Val(Trim(tmpValue)) Or Val(Trim(strValue)) = Val(Trim(tmpValue)) Then
																		tmpResult("Status") = stPass
																		tmpResult("Remarks") = "Expected: NUMBER LESS THAN OR EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
																		Set checkPropertyTest = tmpResult
																		Exit Function
														End If
										Next
										tmpResult("Status") = stFail
										tmpResult("Remarks") = "Expected: NUMBER LESS THAN OR EQUAL (" & strCheckValue & ") Actual: (" & strValue & ")"
						Case Else
		End Select
		If Err.Number <> 0 And tmpResult("Status") = stFail Then tmpResult("Remarks") = Err.Description
		Err.Clear() : On Error GoTo 0
		Set checkPropertyTest = tmpResult
	End Function

	Sub fixObjectReference(ByRef stepReference)
		On Error Resume Next
		If Not IsObject(stepReference) Then
			stepReference = Replace(stepReference, currentRunningPlatform & "-", "")
			If Left(stepReference, 1) = Chr(34) And Right(stepReference, 1) = Chr(34) Then
				stepReference = stepReference
			ElseIf stepReference = "" Then
				stepReference = ""
			ElseIf stepReference = Chr(34) & Chr(34) Then
				stepReference = Chr(34) & Chr(34)
			Else
				Execute "Set stepReference = "  & stepReference
			End If
		End If
	End Sub

	Function formatToParameter(ByVal theString)
				Dim strAlphaNumeric, iCounter, cleanedString, strChar
				cleanedString = ""
				strAlphaNumeric = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ" 'Used to check for numeric characters.
				For iCounter = 1 To Len(theString)
								strChar = Mid(theString, iCounter, 1)
								If InStr(strAlphaNumeric, strChar) Then
												cleanedString = cleanedString & strChar
								Else
												cleanedString = cleanedString & ""
								End If
				Next
				formatToParameter = cleanedString
	End Function

	Public Function omnilite_update_data(data_in, data_out)

			data_from_value = ""
			data_to_value = ""

			If Left(data_in, Len("INPUT[")) = "INPUT[" Then
					data_in_name = Replace(Replace(data_in, "INPUT[", ""), "]", "")
					data_from_value = GetData ("data", data_in_name)
			ElseIf Left(data_in, Len("VAR[")) = "VAR[" Then
					data_in_name = Replace(Replace(data_in, "VAR[", ""), "]", "")
					data_from_value = Eval(data_in_name)
			Else
					data_from_value = data_in
			End If

			If Left(data_out, Len("OUTPUT[")) = "OUTPUT[" Then
					data_out_name = Replace(Replace(data_out, "OUTPUT[", ""), "]", "")
					omni_core.SetData "data", data_out_name, data_from_value
			ElseIf Left(data_out, Len("VAR[")) = "VAR[" Then
					data_out_name = Replace(Replace(data_out, "VAR[", ""), "]", "")
					Execute data_out_name & " = " & Chr(34) & data_from_value & Chr(34)
			ElseIf IsNumeric(data_out) Then
					omni_core.SetData "data", data_out, data_from_value
			Else
					Execute data_out & " = " & data_from_value
			End If

	End Function

	Function performTextOptions(ByVal stepData, ByVal stepOptions)
		If stepOptions = "TEXT[]" Or stepOptions = "TEXT" Then
			stepData = formatToParameter(stepData)
		ElseIf stepOptions = "XDECIMAL_00[]" Or stepOptions = "XDECIMAL_00" Then
			If InStr(1, stepData, ".") <> 0 Then
				stepData = Split(stepData, ".")(0) & "." & Left(Split(stepData, ".")(1),2)
			End If
		ElseIf stepOptions = "RDECIMAL_00[]" Or stepOptions = "RDECIMAL_00" Then
			stepData = myPactera_commonFunction_formatParameter("NUMBER", stepData, "0.00")
		End If
		performTextOptions = stepData
	End Function

	Function myPactera_curTestStep_testScenarioName()
		myPactera_curTestStep_testScenarioName = e("tsn." & ts_i)
	End Function

	Function myPactera_commonFunction_formatParameter(ByVal parameterType, ByVal parameter, ByVal format)
			Dim strLength, strString
			If Trim(UCase(parameterType)) = "NUMBER" Then
					If Len(format) > Len(parameter) Then
							If InStr(1, format, ".", vbTextCompare) = 0 Then
									strLength = Len(format) - Len(parameter)
									strString = StrDup(strLength, "0")
									parameter = strString & parameter
							End If
					End If
			ElseIf Trim(UCase(parameterType)) = "DATE" Then
					parameter = myPactera_commonFunction_formatDate(parameter, format)
			End If
			myPactera_commonFunction_formatParameter = parameter
	End Function

	Function myPactera_commonFunction_formatDate(ByVal dtmInputDate, ByVal strDateTimeFormat)
			'Date/Time Standard Formats - Case-Sensitive.
			'********************************************
			'M : Months 1-12
			'd : Days 1-31
			'yy : Two-digit Year
			'h : Hours 1-12, 12-hour format
			'H : Hours 0-23, 24-hour format
			'm : Minutes 0-59
			'mm : Minutes 00-59
			's : Seconds 0-59
			't : AM or PM
			If Not IsDate(dtmInputDate) Then
					dtmInputDate = Now()
			End If
			Dim strFormattedDateTime
			Dim objFormattedDateRegExp
			objFormattedDateRegExp = CreateObject("VBScript.RegExp")
			objFormattedDateRegExp.IgnoreCase = False
			objFormattedDateRegExp.Global = True
			strFormattedDateTime = strDateTimeFormat
			Dim intMonth
			Dim intDay
			Dim intYear
			Dim intHour
			Dim intMinute
			Dim intSecond
			Dim strWeekday
			Dim strMonth
			Dim strAMPM
			intMonth = Month(dtmInputDate)
			intDay = Day(dtmInputDate)
			intYear = Year(dtmInputDate)
			intHour = Hour(dtmInputDate)
			intMinute = Minute(dtmInputDate)
			intSecond = Second(dtmInputDate)
			strWeekday = WeekdayName(Weekday(dtmInputDate))
			strMonth = MonthName(intMonth)
			strAMPM = "AM"
			objFormattedDateRegExp.Pattern = "MMMM"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, strMonth)
			objFormattedDateRegExp.Pattern = "MMM"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, UCase(Left(strMonth, 3)))
			objFormattedDateRegExp.Pattern = "MM"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, Right("0" & intMonth, 2))
			objFormattedDateRegExp.Pattern = "(M(?=[^AaBbOo])|M$)~^([Ee]M)"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, intMonth)
			objFormattedDateRegExp.Pattern = "dddd"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, strWeekday)
			objFormattedDateRegExp.Pattern = "ddd"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, UCase(Left(strWeekday, 3)))
			objFormattedDateRegExp.Pattern = "dd"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, Right("0" & intDay, 2))
			objFormattedDateRegExp.Pattern = "(d(?=[^EeAaNn])|d$)^([Nn]d|[Ss]d|[Ii]d|[Rr]d)"
			strFormattedDateTime = Replace(strFormattedDateTime, "d", intDay)
			objFormattedDateRegExp.Pattern = "yyyy"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, intYear)
			objFormattedDateRegExp.Pattern = "yy"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, Right(intYear, 2))
			objFormattedDateRegExp.Pattern = "HH"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, Right("0" & intHour, 2))
			objFormattedDateRegExp.Pattern = "(H(?=[^Uu])|H$)^([Cc]H|[Tt]H)"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, intHour)
			If intHour > 12 Then
					intHour = intHour - 12
					strAMPM = "PM"
			End If
			If intHour = 0 Then
					intHour = 12
			End If
			objFormattedDateRegExp.Pattern = "hh"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, Right("0" & intHour, 2))
			objFormattedDateRegExp.Pattern = "(h(?=[^Uu])|h$)^([Cc]h|[Tt]h)"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, intHour)
			objFormattedDateRegExp.Pattern = "mm"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, Right("0" & intMinute, 2))
			objFormattedDateRegExp.Pattern = "(m(?=[^AaBbOo])|m$)^([Ee]m)"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, intMinute)
			objFormattedDateRegExp.Pattern = "ss"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, Right("0" & intSecond, 2))
			objFormattedDateRegExp.Pattern = "(s(?=[^TtEeUuDdAa])|s$)^([Uu]s|[Ee]s|[Rr]s)"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, intSecond)
			objFormattedDateRegExp.Pattern = "(t(?=[^HhUuEeOo])|t$)^([Ss]t|[Aa]t|[Pp]t|[Cc]t)"
			strFormattedDateTime = objFormattedDateRegExp.Replace(strFormattedDateTime, strAMPM)
			myPactera_commonFunction_formatDate = strFormattedDateTime
	End Function

	Function myPactera_filesAndFolders_logOutputFolder()
		myPactera_filesAndFolders_logOutputFolder = e("log_folder")
	End Function

	Function myPactera_commonFunction_convertToText(ByVal strText)
      Dim tmp
      tmp = returnApos(strText)
      tmp = returnDblQoute(tmp)
      Select Case tmp
          Case "[BLANK]"
              tmp = ""
          Case "[SKIP]", "[DEFAULT]"
              tmp = "%SKIP%"
          Case Else
              If InStr(1, tmp, "RANDOM[") <> 0 Then
                  tmp = replaceAllRandom(CStr(tmp))
              End If
      End Select
      myPactera_commonFunction_convertToText = tmp
  End Function

	Function myPactera_commonFunction_getSubString(ByVal strString, ByVal strLeftBound, ByVal strRightBound)
	    Dim tmpValue, tmpOut, LB, RB
	    tmpValue = strString
	    tmpOut = ""
	    If Trim(tmpValue) <> "" Then
	        If Trim(strLeftBound) = "" Then
	            LB = 1
	        Else
	            LB = InStr(1, tmpValue, strLeftBound, vbTextCompare) + Len(strLeftBound)
	        End If
	        If Trim(strRightBound) = "" Then
	            RB = Len(tmpValue) + 1
	        Else
	            RB = InStr(LB, tmpValue, strRightBound, vbTextCompare)
	        End If
	        tmpOut = (Mid(tmpValue, LB, RB - LB))
	    End If
	    myPactera_commonFunction_getSubString = tmpOut
	End Function
	
	Function removeApos(ByVal strText)
					If IsDBNull(strText) Then strText = ""
					removeApos = Replace(strText, "'", "`", vbTextCompare)
	End Function

	Function returnApos(ByVal strText)
					If IsDBNull(strText) Then strText = ""
					returnApos = Replace(strText, "`", "'", vbTextCompare)
	End Function

	Function removeDblQoute(ByVal strText)
					If IsDBNull(strText) Then strText = ""
					removeDblQoute = Replace(strText, Chr(34), "\\q", vbTextCompare)
	End Function

	Function returnDblQoute(ByVal strText)
					If IsDBNull(strText) Then strText = ""
					returnDblQoute = Replace(strText, "\\q", Chr(34), vbTextCompare)
	End Function

	Function convertToDBText(ByVal strText)
					Dim tmp
					tmp = removeApos(strText)
					tmp = removeDblQoute(tmp)
					convertToDBText = tmp
	End Function
	
	Function replaceAllRandom(ByVal strTarget)
        Dim strOutput
        Dim counter
        Dim i
        Dim j
        strOutput = strTarget
        i = InStr(strOutput, "RANDOM[")
        Do While i > 0
            j = InStr(Mid(strOutput, i), "]")
            If j > 0 Then
                j = j + i - 1
                For counter = i + 7 To j - 1
                    strOutput = replaceSubstring(strOutput, generateRandomChar(Mid(strOutput, counter, 1)), counter, counter)
                Next
                strOutput = replaceSubstring(strOutput, "", j, j)
                strOutput = replaceSubstring(strOutput, "", i, i + 6)
            Else
                Exit Do
            End If
            i = InStr(strOutput, "RANDOM[")
        Loop
        replaceAllRandom = strOutput
    End Function

    Function generateRandomChar(ByVal chrTemp)
        Dim chrOutput
        chrOutput = chrTemp
        Randomize()
        Select Case chrOutput
            Case "@" 'a-z'
                chrOutput = Chr(Int((122 - 97 + 1) * Rnd() + 97))
            Case "#" '0-9'
                chrOutput = Chr(Int((57 - 48 + 1) * Rnd() + 48))
            Case "$" 'A-Z'
                chrOutput = Chr(Int((90 - 65 + 1) * Rnd() + 65))
            Case Else
                chrOutput = chrOutput
        End Select
        generateRandomChar = chrOutput
    End Function
