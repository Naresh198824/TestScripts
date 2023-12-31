'*********************************************************************************************************************
'**  Project            : 
'**  File Name          : Functions.vbs
'**  Version            : 
'**  Created on         : 
'**  Updated on         : 
'**  Description        : Generic functions that can be used across the project
'**  Copyright          : Arsin Corporation.
'**  Author             :       
'*********************************************************************************************************************

Sub SetObjectValue (oObject, sData)

		' @HELP
		' @group	: Functions	
		' @method	: SetObjectValue (oObject, sData)
		' @returns	: None
		' @parameter: oObject: Name of the Object to set the value
		' @parameter: sData: Value to set for the Object
		' @notes	: To set values for WebEdit,WebList and WebRadiogroup
		' @END
		
		Dim sTypeName
	
		If IsNull (sData) OR IsEmpty (sData) Then
			MsgBox "Expression - True"
		End If
	
		sTypeName = TypeName (oObject)
	
		Select Case sTypeName
			Case "WebEdit"			' WebEdit object
				object.Set sData
			Case "WebList"			' WebList object
				object.Select sData
			Case "WebRadioGroup"	' WebRadioGroup Object
				object.Select sData
			Case Else
				ReportWriter "Fail", "Wrong Usage: SetObjectValue", "Cannot call this function on the object of type: " & sTypeName,0
				ExitActionIteration 
		End Select

End Sub
'-----------------------------------------------------------------------
Function GetObjectValue (object)

		' @HELP
		' @group	: Functions	
		' @method	: GetObjectValue (object)
		' @returns	: Returns Object value
		' @parameter: object: Type of the Object(HtmlEdit,HtmlList,......)
		' @notes	: To get Object values from HtmlEdit,HtmlList,.......
		' @END	
		
		Dim sTypeName, sValue
	
		sTypeName = TypeName (object)
	
		Select Case sTypeName
			Case "HtmlEdit"			' WebEdit object
				sValue = object.GetROProperty ("value")
			Case "HtmlList"			' WebList object
				sValue = object.GetROProperty ("value")
			Case "HtmlRadioGroup"	' WebRadioGroup Object
				sValue = object.GetROProperty ("value")
			Case "ImageLink"		' Image Object
				sValue = object.GetROProperty ("alt")
			Case "TextLink"			' Link Object
				sValue = object.GetROProperty ("text")
			Case Else
				ReportWriter "Fail", "Wrong Usage: [GetObjectValue]", "Cannot call this function on the object of type: " & sTypeName,0
				ExitActionIteration 
		End Select
	
		GetObjectValue = sValue

End Function
'-----------------------------------------------------------------------
Function CaptureErrorAsBitmap ()

		' @HELP
		' @group	: Functions	
		' @method	: CaptureErrorAsBitmap
		' @returns	: None
		' @parameter: 
		' @notes	: To Capture an Error as Bitmap Image
		' @END
		
		Dim sImageName, sTestName, sTestcaseID
	
		sTestName	= Environment.Value("TestName")
		sTestcaseID	= Environment.Value("TestcaseId")
	
		sImageName	= ERRORS_DIR & sTestName &"_" & sTestcaseID &".bmp"
	
		Desktop.CaptureBitmap sImageName, TRUE
		ReportWriter micDone, "Desktop Image", "The Desktop Image location is: " &sImageName,0
	
End Function
'-----------------------------------------------------------------------
Function GetTestcaseID_TD ()
	
		' @HELP
		' @group	: Functions	
		' @method	: GetTestcaseID_TD
		' @returns	: None
		' @parameter:  
		' @notes	: To get the Test Case Id from Test Director
		' @END
		
		Dim currentTestSetTest
		set currentTestSetTest = TDUtil.CurrentTestSetTest
		GetTestcaseID_TD = currentTestSetTest.Field(TESTCASE_ID)
End Function
'-----------------------------------------------------------------------
Function GetTestcaseID_QC ()
	
		' @HELP
		' @group	: Functions	
		' @method	: GetTestcaseID_TD
		' @returns	: None
		' @parameter:   
		' @notes	: To get the Test Case Id from Quality Center
		' @END
		
		Dim currentTestSetTest
		set currentTestSetTest = QCUtil.CurrentTestSetTest
		GetTestcaseID_QC = currentTestSetTest.Field(TESTCASE_ID)
		
End Function
'-----------------------------------------------------------------------
Function CaptureAndCloseError(Object, Method, Arguments, retVa)

		' @HELP
		' @group	: Functions	
		' @method	: CaptureAndCloseError(Object, Method, Arguments, retVa)
		' @returns	: None
		' @parameter: Object:
		' @parameter: Method:
		' @parameter: Arguments:
		' @parameter: retVa:
		' @notes	: To Capture and Close Error 
		' @END
		
	sTestcaseID = Environment.Value("TestcaseId")
	sImageName	= sErrorPath&Environment.Value("TestName")+"_"+sTestcaseID&".bmp"
	Desktop.CaptureBitmap sImageName,TRUE
	ReportWriter 1, "Testcase Failed", "Test case failed, closing the omni application",0

End Function
'-----------------------------------------------------------------------
Function ReportError(Object, Method, Arguments, retVal)

	ReportWriter 1, "Testcase Failed", "Test case failed, closing the application",0
	
End Function
'-----------------------------------------------------------------------
' Takes the actual and expected data as input and verifies if sExpData pattern exists in sActData
Function StringPartialMatch (sExpData,sActData)
		
		' @HELP
		' @group	: Functions	
		' @method	: StringPartialMatch (sExpData,sActData)
		' @returns	: None
		' @parameter: sExpData: Expected data to verify
		' @parameter: sActData: Actual data to verify
		' @notes	: Verifies if sExpData exists in sActData
		' @END
		
   		For j = 1 to Len(sActData)
			If  Trim(Mid(sActData, j, Len(sExpData))) = sExpData Then
				bFlag = True
                		ReportWriter 0,"String Partial Comparision","String Partial Match Successful. Actual is: " +Trim(Mid(sActData, j, Len(sExpData))) + "  Expected is: "+ sExpData,0
				Exit for
			End If
		Next
		If not bFlag = True then
			ReportWriter 1,"String Partial Comparision","String Partial Match Not Successful. Actual is: " + sActData + "  Expected is: "+ sExpData,0
		End if
		
End Function
'-----------------------------------------------------------------------
'Genrate Random number for a given a range.
Function RandomNumberGen()

		' @HELP
		' @group	: Functions	
		' @method	: RandomNumberGen
		' @returns	: Returns Random Number
		' @parameter:   
		' @notes	: To generate Random Numbers
		' @END
		
  	Randomize
	sUpperbound = "199999"
	sLowerbound = "100000"
    sValue = Int((sUpperbound - sLowerbound + 1005) * Rnd + sLowerbound)
	RandomNumberGen = sValue
	
End Function

Function RandomNumberGenRange(sUpperbound, sLowerbound)

		' @HELP
		' @group	: Functions	
		' @method	: RandomNumberGen
		' @returns	: Returns Random Number
		' @parameter:   
		' @notes	: To generate Random Numbers
		' @END
		
  	Randomize
'	sUpperbound = "199999"
'	sLowerbound = "100000"
    sValue = Int((sUpperbound - sLowerbound) * Rnd + sLowerbound)
	RandomNumberGenRange = sValue
	
End Function
'-----------------------------------------------------------------------
'WinZip evaluation version: When you open WinZip file it will open with 3 buttons. Button position "Use Evaluation Verison" is not constant. 
'Hot keys also changing from Ctrl+E to Ctrl+V. This function will handle both issues.
Function ClickOnWinZipUseEvaluationButton ()

		' @HELP
		' @group	: Functions	
		' @method	: ClickOnWinZipUseEvaluationButton
		' @returns	: None
		' @parameter:   		
		' @notes	: To Click on WinZip Evaluation Button and to Extract Files
		' @END
	
	Set WS = CreateObject("Wscript.Shell")
	WS.AppActivate "WinZip"
	WS.SendKeys "^{V}"

	Const WM_CLOSE = "&H10"
	Extern.Declare micHwnd, "FindWindow", "user32.dll", "FindWindowA", micString, micString 
	Extern.Declare micHwnd, "SendMessage", "user32.dll", "SendMessageA", micHwnd, micLong, micLong, micString 
	hwnd = Extern.FindWindow(vbNullString, "WinZip")
	
	If hwnd <> 0 then 
		res=Extern.SendMessage(hwnd, WM_CLOSE, 256, String(256," ")) 
		WS.AppActivate "WinZip"
		WS.SendKeys "^{E}"
	End If
	
End Function
'--------------------------------------------------------------------------------------------
Function ClickOnWinzipOpen_Extract(sZipFileName)

		' @HELP
		' @group	: Functions	
		' @method	: ClickOnWinzipOpen_Extract(sZipFileName)
		' @returns	: None
		' @parameter: sZipFileName	: Zip file name to Extarct
		' @notes	: To Capture an Error as Bitmap Image
		' @END

		' This function is for winzip 11 Eval version	
			'Set WS = CreateObject("Wscript.Shell")
			'WS.AppActivate "File Download"
			'WS.SendKeys "{LEFT}{LEFT}{ENTER}"
			'MsgBox "Opening Files"
			'wait 5
	
	Set WS1 = CreateObject("Wscript.Shell")
	WS1.AppActivate "WinZip"
	WS1.SendKeys "^{V}"

	Const WM_CLOSE = "&H10"
	Extern.Declare micHwnd, "FindWindow", "user32.dll", "FindWindowA", micString, micString 
	Extern.Declare micHwnd, "SendMessage", "user32.dll", "SendMessageA", micHwnd, micLong, micLong, micString 
	hwnd = Extern.FindWindow(vbNullString, "WinZip")
	
	'If  hwnd is not equal then there is a popup window
	If hwnd <> 0 then 
		res=Extern.SendMessage(hwnd, WM_CLOSE, 256, String(256," ")) 
		WS1.AppActivate "WinZip"
		WS1.SendKeys "^{E}"
	end if
	
	set ws2	= CreateObject("Wscript.Shell")
	ws2.AppActivate "WinZip (Evaluation Version) - "&ZipFileName&".*"
	ws2.SendKeys "^{A}" 
  	wait 5
  		
	Window("regexpwndclass:=WinZipWClass").WinToolbar("regexpwndclass:=ToolbarWindow32","object class:=ToolbarWindow32").Press "Extract"
	Window("regexpwndclass:=WinZipWClass").Dialog("text:=Extract.*").WinEdit("attached text:=E&xtract to:","x:=113").Set NEW_FILE_LOCATION
	Window("regexpwndclass:=WinZipWClass").Dialog("text:=Extract.*").WinButton("text:=&Extract").Click
	wait 4
	If Window("regexpwndclass:=WinZipWClass").Dialog("text:=Confirm File Overwrite").Exist Then
		Window("regexpwndclass:=WinZipWClass").Dialog("text:=Confirm File Overwrite").WinButton("text:=Yes to &All").Click
	End If
	Wait 3
	Window("regexpwndclass:=WinZipWClass").Close
	
End Function
'-----------------------------------------------------------------------
Function VerifyValues(sV1, sV2)

		' @HELP
		' @group	: Functions	
		' @method	: VerifyValues(sV1, sV2)
		' @returns	: None
		' @parameter: sV1: Value1 to Verify
		' @parameter: sV2: Second Value to Verify
		' @notes	: To Verify whether two given Values are equal or not
		' @END
		
'	If (sV1 = sV2) Then
'		Services.LogMessage "Data verification Successful.",OutputMsg
'		Services.LogMessage " Original Value:" &sV1,OutputMsg 
'		Services.LogMessage " Expected Value:" &sV2,OutputMsg
'	Else
'		Services.LogMessage "Data verification failed.",ErrorMsg
'		Services.LogMessage " Original Value:" &sV1,OutputMsg 
'		Services.LogMessage " Expected Value:" &sV2,OutputMsg
'	End If
	
	If (sV1 = sV2) Then
		'Services.LogMessage "Data verification Successful.",OutputMsg
		ReportWriter "Pass", "Compare Results", "Original Value: " +sV1 &" , Actual Value:" +sV2,0
	Else
		'Services.LogMessage "Data verification failed.",ErrorMsg
		ReportWriter "Fail", "Compare Results ", "Original Value: " +sV1 &" , Actual Value:" +sV2,0
		End If
	
End Function
'-----------------------------------------------------------------------
Function DateAfterDays(iInterval)
	
		' @HELP
		' @class	: Functions
		' @method	: DateAfterDays(iInterval)
		' @returns	: Date Expression
		' @parameter: iInterval:  Numeric expression that is the number of days after.
		' @notes	: This method  returns the  date after iInterval days of current Date.
		' @END
		
		DateAfterDays=FormatDateTime(DateAdd("d",iInterval,Date),2)
		MsgBox FormatDateTime(DateAdd("d",iInterval,Date),2)
End Function
'-----------------------------------------------------------------------
'Usage: MessageBox "OK"
Function MessageBox(Message)

	  ' @HELP
	  ' @group   	: Functions
	  ' @method  	: MessageBox(Message)
	  ' @returns 	: None
	  ' @parameter 	: Message: Any message which we need to be displayed
	  ' @notes  	: Shows Any message which we need to be displayed
	  ' @END 

Set WshShell = CreateObject("WScript.Shell")
WshShell.Popup Message, 5, "UserMessage"

End Function

'-------------------------------------------------------------------------------------------------------------------------------------------
Sub ChangeTimeZoneToIST()
	Set WshShell=createobject("wscript.shell")
	WshShell.Run "RunDLL32 shell32.dll,Control_RunDLL  %SystemRoot%\system32\TIMEDATE.cpl,,/Z (GMT+05:30) Chennai, Kolkata, Mumbai, New Delhi", 0, True
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------
Sub ChangeTimeZoneToPST()
    Set WshShell=createobject("wscript.shell")
	WshShell.Run "RunDLL32 shell32.dll,Control_RunDLL  %SystemRoot%\system32\TIMEDATE.cpl,,/Z Pacific Standard Time", 0, True
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------
Function getPSTDate()
If Reporter.GetTimeZone="IST" Then
	If Time>"12:30 PM" then
		sPSTDate = Date
		Else
		sPSTDate = DateAdd("d",-1,Date)
	End if
	Else
	sPSTDate = Date
End If
sPSTDate = Replace(sPSTDate,"-","/")
aDate = Split(sPSTDate,"/")
If aDate(0)<10 Then
	smonth ="0"&int(aDate(0))
	Else
	smonth =aDate(0)
End If
If aDate(1)<10 Then
	sDay ="0"&int(aDate(1))
	Else
	sDay =aDate(1)
End If
sPSTDate = smonth&"/"&sDay&"/"&aDate(2)
getPSTDate = sPSTDate
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------
Function ChangeDateFormat(sReqDateFormat,sDateToBeChanged)
	
	Set WshShell = CreateObject("WScript.Shell")
	WshShell.RegWrite("HKEY_USERS\.Default\Control Panel\International\sDate"),"/","REG_EXPAND_SZ"
	WshShell.RegWrite("HKEY_USERS\.Default\Control Panel\International\sShortDate"),sReqDateFormat,"REG_EXPAND_SZ"
	WshShell.RegWrite("HKEY_CURRENT_USER\Control Panel\International\sDate"),"/","REG_EXPAND_SZ"
	WshShell.RegWrite("HKEY_CURRENT_USER\Control Panel\International\sShortDate"),sReqDateFormat,"REG_EXPAND_SZ" 
	wait(5)
	varDate = sDateToBeChanged
	
	ChangeDateFormat=FormatDateTime(varDate, 2)
	WshShell.RegWrite("HKEY_USERS\.Default\Control Panel\International\sDate"),"/","REG_EXPAND_SZ"
	WshShell.RegWrite("HKEY_USERS\.Default\Control Panel\International\sShortDate"),"MM/dd/yyyy","REG_EXPAND_SZ"
	WshShell.RegWrite("HKEY_CURRENT_USER\Control Panel\International\sDate"),"/","REG_EXPAND_SZ"
	WshShell.RegWrite("HKEY_CURRENT_USER\Control Panel\International\sShortDate"),"MM/dd/yyyy","REG_EXPAND_SZ" 
	wait (5)
	
End Function
'---------------------------------------------------------------------------------------------------------------- 
'To Check for the Security popup window and click on "Yes" button
Function VerifySecurityInformationPopup()

		' @HELP
		' @group	: Functions
		' @method	: VerifySecurityInformationPopup()
		' @returns	: None
		' @parameter: None
		' @notes	: This function is used to check for Security Popup and click on "yes" button
		' @END
		
	If Browser("micClass:=Browser").Dialog("regexpwndtitle:=Security Information").Exist Then		
		Browser("micClass:=Browser").Dialog("regexpwndtitle:=Security Information").WinButton("regexpwndtitle:=&Yes").Click
		ReportWriter "Pass", " Security PoPup Exists ", " Clicked on Yes Button",0	
	Else
		ReportWriter "Pass", " Security PoPup Not Exists ", " Continue with Login "	,0
End If
End Function
'---------------------------------------------------------------------------------------------------------------- 

'To Check for the Security Warning popup window and click on "Yes" button
Function VerifySecurityWarningPopup()

		' @HELP
		' @group	: Functions
		' @method	: VerifySecurityWarningPopup()
		' @returns	: None
		' @parameter: None
		' @notes	: This function is used to check for Security Popup and click on "yes" button
		' @END
		
	If Browser("micClass:=Browser").Dialog("regexpwndtitle:=Security Warning").Exist Then		
		Browser("micClass:=Browser").Dialog("regexpwndtitle:=Security Warning").WinButton("regexpwndtitle:=&Yes").Click
		ReportWriter "Pass", " Security PoPup Exists ", " Clicked on Yes Button",0	
	Else
		ReportWriter "Pass", " Security PoPup Not Exists ", " Continue with Login "	,0
End If
End Function

'Usage sDate = ReadPSTDate()
Function ReadPSTDate()

  ' @HELP
  ' @method   : ReadPSTDate()
  ' @returns  : Returns date of Pacific Standard Time
  ' @notes    : This function checks the Time zone. If time zone is IST, it converts the time zone to 
  ' 			PST and then reads the PST date and changes time zone again to IST. If time zone is PST it directly gets PST date
  ' @END 
  
If Reporter.GetTimeZone="IST" Then
	If Time>"12:30 PM" then
		sPSTDate = Date
		Else
		sPSTDate = DateAdd("d",-1,Date)
	End if
Else
	sPSTDate = Date
End If

aDate = Split(sPSTDate,"/")
If aDate(0)<10 Then
	smonth ="0"&int(aDate(0))
	Else
	smonth =aDate(0)
End If
If aDate(1)<10 Then
	sDay ="0"&int(aDate(1))
	Else
	sDay =aDate(1)
End If
sPSTDate = smonth&"/"&sDay&"/"&aDate(2)
	
ReadPSTDate = sPSTDate
End Function
'---------------------------------------------------------------------

Function VerifyProductExpiredPSTDate(sPSTDate,iTermInYears)			
			
	'USDate = DateAdd("n",-810,now)
	PSTDate = FormatDateTime(sPSTDate, 2)
	tempExpDate = DateAdd("YYYY",iTermInYears,PSTDate)
	ExpDate = DateAdd("d",-1,tempExpDate)
	VerifyProductExpiredPSTDate = ExpDate
End Function

'---------------------------------------------------------------------


Function checkASC(arr)
	
	' @HELP
		' @group	: Functions
		' @method	: checkASC()
		' @returns	: Boolean Value(true or false)
		' @parameter: Array containing integer elements
		' @notes	: This function is used to check wheather Array is in Ascending order
		' @END
		
checkASC=true
For j=0 to ubound(arr)-1
	If cint(arr(j))<cint(arr(j+1)) Then
		checkASC=true
	Else
		checkASC=false
		Exit for
	End if

Next

End Function
'---------------------------------------------------------------------

Function checkDSC(arr)

	' @HELP
		' @group	: Functions
		' @method	: checkDSC()
		' @returns	: Boolean Value(true or false)
		' @parameter: Array containing integer elements
		' @notes	: This function is used to check wheather Array is in Descending order
		' @END
checkDSC=true
For j=0 to ubound(arr)-1
	If cint(arr(j))>cint(arr(j+1)) Then
		checkDSC=true
	Else
		checkDSC=false
		Exit for
	End if

Next
End Function
'---------------------------------------------------------------------
Function ConvertArrayToDictionary(cArrayName)

Dim cDictionary

   Set cDictionary=CreateObject("scripting.dictionary")

	For aIndex=0 to UBound(cArrayName)
			cDictionary.Add cArrayName(aIndex),""
	Next

Set ConvertArrayToDictionary=cDictionary

End Function
'---------------------------------------------------------------------

Function GetRandomDate(sDate1, sDate2)
'
'Date1 = "10/02/2017"
'Date2 = "10/06/2017"
iDiff = DateDiff("d", cdate(sDate1), cdate(sDate2))
stemp = RandomNumberGenRange(0, cstr(iDiff))
RndDate = DateAdd("d",stemp , cdate(sDate1))
RndDate = cdate(RndDate)
GetRandomDate = Day(RndDate) & "/" & Month(RndDate) & "/" & Year(RndDate)

End Function

