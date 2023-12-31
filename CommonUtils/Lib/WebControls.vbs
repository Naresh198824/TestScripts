'****************************************************************************************************************************
' $Filename:		WebControls.vbs
' $Description: 	WebControls 
' $Copyright: 		Arsin Corporation
'****************************************************************************************************************************
Dim oObjectDescriptions
Dim oSesObjDescriptions
Dim BaseWindow

Public Const Short_Interval = 2
Public Const Long_Interval = 2

Function GetObjectDescriptions ()
		Set oSesObjDescriptions = Description.Create ()
		oSesObjDescriptions("micClass").Value = "Browser"
		
		Set oObjectDescriptions = Description.Create ()
		oObjectDescriptions("micClass").Value = "Page"
End Function  
'****************************************************************************************************************************************************************************************************************************************************************************************
Sub CloseAllBrowsers(sBrowser) 
 
		' @HELP
		' @group	: Webcontrols	
		' @funcion	: CloseAllBrowsers(sBrowser)
		' @returns	: None
		' @parameter: sBrowser : Browser Type (ex: IE Or FF)
		' @notes	: Closes All Browser based on Browser type
		' @example	: CloseAllBrowsers("IE")
		' @END
		
		If UCase(sBrowser) = "IE" Then
			SystemUtil.CloseProcessByName "IExplore.exe"
		Else
			SystemUtil.CloseProcessByName "firefox.exe"   
		End If
		
End Sub 
'****************************************************************************************************************************************************************************************************************************************************************************************
Function InvokeBrowser (sURL, sBrowser)

		Rem # HELP
		Rem # group	: Webcontrols	
		Rem # funcion	: InvokeBrowser (sURL, sBrowser)
		Rem # returns	: None
		Rem # parameter: sURL : Name of the URL to invoke
		Rem # parameter: sBrowser : Name of the Browser (Ex: IE Or FF)
		Rem # notes	: Invoke the URL for the given browser
		Rem # example	: InvokeBrowser ("http://www.google.com","IE")
		Rem # END
		
		On Error Resume Next 
		
		If UCase(sBrowser) = "IE" Then
			SystemUtil.Run "iexplore.exe", sURL,"","",3
			wait Short_Interval
	
		'To Terminate rundll32.exe 
		SystemUtil.CloseProcessByName "rundll32.exe"
			
	 	ElseIf UCase(sBrowser) = "FF" Then
			SystemUtil.Run "firefox.exe", sURL
			wait Short_Interval
		End If
		
		If Err = 0 Then
	 		ReportWriter "Pass","Open Browser:  '"& UCase(sBrowser)&"'", "Browser:  '","Opened  With URL:  '" & sURL & "'",INDEX_VALUE_ZERO
	 	Else 
	 		ReportWriter "Fail","Open Browser:  '"& UCase(sBrowser)&"'", "Browser:  '","Not Opened  With URL:  '" & sURL & "' :" & Err.Description,INDEX_VALUE_ZERO
	 	End If
	 	
	 	Err.Clear
	 	
	 	
End Function


Function InvokeBrowser_Ora (sURL, sBrowser,iBrowIndex)
		' @HELP
		' @group	:	WebControls	
		' @method	:	InvokeBrowser (sURL, sBrowser, iBrowIndex)
		' @returns	:	None
		' @parameter:   sURL		: Name of the URL to invoke
		' @parameter:   sBrowser	: Name of the Browser
		' @parameter:   iBrowIndex	: Index of the Browser
		' @notes	: 	Invoke the URL for the given browser
		' @END
		
		GetObjectDescriptions()
		If UCase(sBrowser) = "IE" Then
			SystemUtil.Run "iexplore.exe", sURL,"","",3
			wait 8
			If Browser(oSesObjDescriptions,"Creationtime:="&iBrowIndex).Page(oObjectDescriptions).Exist(3) Then
				ReportWriter "PASS","Browser", "URL:"&sURL&" Is Successfully Opened in the Browser",0
            Else
				ReportWriter "FAIL","Failed to Open Browser","Unable to Open the Browser",0
				ExitRun
				sOpen = False
			End If
		ElseIf UCase(sBrowser) = "FF" Then
			SystemUtil.Run "firefox.exe", sURL
			wait 8
		Else
			SystemUtil.Run "netscp6.exe", sURL
		End If
End Function
'*********************************************************************************************************************************************************************
Function GetBrowserCount()
		' @HELP
		' @group	:	WebControls	
		' @method	:	GetBrowserCount
		' @returns	:	Returns the count of active browsers on the desktop
		' @parameter:   None
		' @notes	:	GetBrowserCount
		' @End 

		 
	GetObjectDescriptions()	
	Set oDesc = Description.Create ()
	oDesc("micclass").Value = "Dialog"
	Set oDlg = DeskTop.ChildObjects(oDesc)
	For i = 0 To oDlg.Count-1
		If Dialog("micclass:=Dialog","index:="&i).Exist Then
			Dialog("micclass:=Dialog","index:="&i).Close
		End If
	Next		
	Set oDesc = Nothing

	Set oBrowser = Description.Create()
	oBrowser("application version").Value = "internet explorer.*"
	Set oBrow = Desktop.ChildObjects(oBrowser)	
	For i = 0 to oBrow.count-1
		If oBrow(i).GetROProperty("name") = "Effecta" or oBrow(i).GetROProperty("name") = "Effecta Execution::" Then
			sFound = True
		Else
			If Browser(oSesObjDescriptions,"Creationtime:="&i).Page(oObjectDescriptions).JavaWindow("tagname:=PluginEmbeddedFrame").JavaButton("label:=Cancel").Exist(10) Then
				Browser(oSesObjDescriptions,"Creationtime:="&i).Page(oObjectDescriptions).JavaWindow("tagname:=PluginEmbeddedFrame").JavaButton("label:=Cancel").Click
   			End If 
			ClearAllCookies			
			oBrow(i).close  
		End If
	Next

	'If oBrow.count = 0 then
		'SystemUtil.Run "iexplore.exe", "","","",0  ''Added to overcome cookies problem
		'Wait 2
		'ClearAllCookies	
		'Set oBrow = Desktop.ChildObjects(oBrowser)		
		'oBrow(0).close
	'End If
	
	Set oBrow = Desktop.ChildObjects(oBrowser)
	GetBrowserCount = oBrow.Count
		
	'Close Java Process
'systemutil.CloseProcessByName "javaw.exe"
'	systemutil.CloseProcessByName "java.exe"
	


End Function

'*********************************************************************************************************************************************************************
Sub VerifyPageExists(iCreationtime,sWebTablePropAndVal,iIndex,sPageTxt)
	   '@HELP
	   '@group    :  	WebControls 
	   '@method   :   	VerifyPageExists(iCreationtime,sWebTablePropAndVal,iIndex,sPageTxt)
	   '@returns  :   	None
	   '@parameter:   iCreationtime  :  Creationtime of Browser
	   '@parameter: 	sWebTablePropAndVal: Property and Value of WebTable
	   '@parameter: 	iIndex			   : Index of WebTable
	   '@parameter: 	sPageTxt	       : Page text/caption to be verified
	   '@notes    :   	Checks whether Page Existing Or Not
	   '@END
	   
		GetObjectDescriptions ()
		If  Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sWebTablePropAndVal,"index:="&iIndex).Exist  Then
			sText = Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sWebTablePropAndVal,"index:="&iIndex).GetROProperty("innertext")
			If sText = sPageTxt Then
				ReportWriter "PASS", "Page", sText &" Page exists",0
				
			Else
				ReportWriter "FAIL", "Page", sText &" Page doesn't  exists",0
			End If
    	Else
    		ReportWriter "FAIL", "Page", sText & " Page doesn't  exists",0
    	End If
End Sub	

'*********************************************************************************************************************************************************************
Sub ClickLink(iCreationtime,sLnkPropAndValue)
		' @HELP
		' @group	:	WebControls
		' @method	: 	ClickLink(iCreationtime,sLnkPropAndValue,sLinkText)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sLnkPropAndValue : Property & value of the Link
		' @notes	:	Used to Click the Link 
		' @END 
		
		sLinkText = Split(sLnkPropAndValue,":=")
		GetObjectDescriptions()
		If Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).Link(sLnkPropAndValue).Exist(.1) Then
			Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).Link(sLnkPropAndValue).Click()		
			ReportWriter "PASS","Link:" ,sLinkText(1) & "   is clicked",0
		Else
			ReportWriter "FAIL","Link:" ,sLinkText(1) & "   is not clicked",0
			ExitRun
		End If			
End Sub
'*********************************************************************************************************************************************************************
Sub ClickLinkInTable(iCreationtime,sTblPropnVal,sLnkPropAndValue)
		' @HELP
		' @group	:	WebControls
		' @method	: 	ClickLink(iCreationtime,sLnkPropAndValue,sLinkText)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sLnkPropAndValue : Property & value of the Link
		' @notes	:	Used to Click the Link 
		' @END 
		
		sLinkText = Split(sLnkPropAndValue,":=")
		GetObjectDescriptions()
		If Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropnVal).Link(sLnkPropAndValue).Exist(.1) Then
			Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropnVal).Link(sLnkPropAndValue).Click()		
			ReportWriter "PASS","Link:" ,sLinkText(1) & "   is clicked",0
		Else
			ReportWriter "FAIL","Link:" ,sLinkText(1) & "   is not clicked",0
			ExitRun
		End If			
End Sub
'*********************************************************************************************************************************************************************
Function SelectItemFromWebList(iCreationtime,sWebListPropAndValue,sValue)
		' @HELP
		' @group	:	WebControls
		' @method	:	SelectItemFromWebList (iCreationtime,sWebListPropAndValue,sValue)
		' @returns	:	None
		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sWebListPropAndValue : PropName & Value of the List box
		' @parameter:   sValue	             : Value to be selected from the List
		' @notes	:	Used to Select item in the WebList 
		' @END 
		
		GetObjectDescriptions ()		
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebList(sWebListPropAndValue).Exist Then
			Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebList(sWebListPropAndValue).Select sValue
			Wait 2
			'sVal = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebList(sWebListPropAndValue).GetROProperty("Selection")
		    ReportWriter "PASS", "List Box", sValue &"   is selected from the list",0
		Else
		    ReportWriter "FAIL","List Box", sValue &"   is not selected from the list",0
		End If
		SelectItemFromWebList = sValue
End Function
'*********************************************************************************************************************************************************************
Function SetDataInWebEdit(iCreationtime,sEditPropAndValue,sValue)
'    	 @HELP
'		 @group	    : 	WebControls
'		 @method	: 	SetDataInWebEdit(iCreationtime,sEditPropAndValue,sValue)
'		 @returns	: 	None
' 		 @parameter :   iCreationtime  :  Creationtime of Browser
'		 @parameter :	sEditPropAndValue: Property Name & Value of the Object
'		 @parameter :	sValue: Value to be set In the WebEdit
'		 @notes	    : 	This method is used to set data In WebEdit
'		 @END
		
		GetObjectDescriptions ()
		sEdit = Split(sEditPropAndValue,"=")
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebEdit(sEditPropAndValue).Exist(.1) Then

			Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebEdit(sEditPropAndValue).Set sValue
			ReportWriter "PASS", "Edit Box" , sValue & "  value is set in  " & sEdit(1),0		
    	Else
    		ReportWriter "FAIL", "Edit Box" , sValue & "  value is not set in  " & sEdit(1),1
    	End If 
End Function
'*********************************************************************************************************************************************************************
Sub ClickButton(iCreationtime,sBtnPropAndValue,iIndex)
		' @HELP
		' @group	: 	WebControls
		' @method	:	ClickButton(iCreationtime,sBtnPropAndValue,iIndex)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sBtnPropAndValue :  Property name and value of WebButton
		' @parameter:   iIndex			 :  index of WebButton
		' @notes	:	Used to Click button present in webpage 
		' @END
		 
		GetObjectDescriptions()
		sText = Split(sBtnPropAndValue,":=")
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebButton(sBtnPropAndValue,"index:="&iIndex).Exist(.1) Then
			Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebButton(sBtnPropAndValue,"index:="&iIndex).Click
			ReportWriter "PASS","WebButton",sText(1) & "  is clicked",0
		Else
			ReportWriter "FAIL","WebButton",sText(1) & "  is not clicked",0
		End If
End Sub
'*********************************************************************************************************************************************************************
Function CompareDataInWebTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,iColID,sData2Verify,sText)
		
		' @HELP
		' @group	: 	WebControls
		' @method	:	CompareDataInWebTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,iColID,sData2Verify,sText)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   iRowID 		   :  Row ID in which data is to be compared
		' @parameter:   iColID		   :  Column ID in which data is to be compared
		' @parameter:   sData2Verify   :  Data to be verified
		' @parameter:   sText		   :  Text/Column name for reporting purpose
		' @notes	:	Compares data present in WebTable 
		' @END
		
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
			sValue = Trim(Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).GetCellData(iRowID,iColID))

			If UCase(trim(sValue) )= UCase(trim(sData2Verify)) Then
				ReportWriter "PASS", "WebTable" , sText & " Data is verified." & vbLf & "Expected : " & sData2Verify & vbLf & "Actual : " & sValue,0
			Else
    			ReportWriter "FAIL", "WebTable" , sText & " data doesn't match." & vbLf & "Expected : " & sData2Verify & vbLf & "Actual : " & sValue,1
    		End If
    	Else
    		ReportWriter "FAIL", "WebTable" , "WebTable doesn't exist",0
    	End If 
    	CompareDataInWebTableByRowAndColID = sValue
    	
End Function
'*********************************************************************************************************************************************************************
Function VerifyValExistInWebTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,iColID,sText)
		
		' @HELP
		' @group	: 	WebControls
		' @method	:	VerifyValExistInWebTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,iColID,sText)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   iRowID 		   :  Row ID in which data is to be compared
		' @parameter:   iColID		   :  Column ID in which data is to be compared
		' @parameter:   sText		   :  Text/Column name for reporting purpose
		' @notes	:	verifies whether any data is present in WebTable or not
		' @END
		
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
			sValue = Trim(Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).GetCellData(iRowID,iColID))
			If IsNumeric(sValue) OR Len(sValue) > 0 Then
				ReportWriter "PASS", "WebTable" , sText & " is present. Value is " & sValue,0
			Else
    			ReportWriter "FAIL", "WebTable" , sText & " is not present. Value is " & sValue,1
    		End If
    	Else
    		ReportWriter "FAIL", "WebTable" , "WebTable doesn't exist",0
    	End If 
    	VerifyValExistInWebTableByRowAndColID = sValue
End Function
'*********************************************************************************************************************************************************************
Sub ClickLinkInTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowId,iColID)

		' @HELP
		' @group	: 	WebControls
		' @method	:	ClickLinkInTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowId,iColID)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   iRowID 		   :  Row ID in which data is to be compared
		' @parameter:   iColID		   :  Column ID in which data is to be compared
		' @notes	:	Clicks link present in WebTable based on specifed row and column ID
		' @END
		
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
			Set oDesc = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).ChildItem(iRowId,iColID,"Link",0)
			oDesc.Click
			ReportWriter "PASS","Link","Clicked on the Link  " & oDesc.GetROProperty("name"),0
		Else
			ReportWriter "FAIL","WebTable","WebTable Doesn't exist",0
		End If	
End Sub
'*********************************************************************************************************************************************************************
 Function GetDefaultValOfWebEdit(iCreationtime,sEditPropAndVal)
    	' @HELP
		' @group    :  WebControls
		' @method   :  GetDefaultValOfWebEdit(iCreationtime,sEditPropAndVal)
		' @returns  :  Default Value of WebEdit
		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:  sEditProp : Property name and value of the WebEdit Object
		' @notes    :  Returns Default Value of WebEdit
		' @END
		
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebEdit(sEditPropAndVal).Exist(.1) Then
			GetDefaultValOfWebEdit = Trim(Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebEdit(sEditPropAndVal).GetROProperty("default value"))
			ReportWriter "PASS", "WebEdit" , "Default Value of WebEdit is : " & GetDefaultValOfWebEdit,0
		Else
    		ReportWriter "FAIL", "WebEdit" , "WebEdit does not exists",0
    	End If 
  End Function
'*********************************************************************************************************************************************************************
Function GetDateInRequiredFormat(iDate,sDateFormat)
		' @HELP
		' @group	: 	WebControls
		' @method	:	GetDateInRequiredFormat(iDate,sDateFormat)
   		' @returns	:   Date in Required format     
   		' @parameter:   iDate      : Required date in "MM/DD/YYYY" Format
   		' @parameter:   sDateFormat: Required date format  (Ex:"DD-MMM-YYYY ->05-May-2008","MMM-DDD-YY ->May-Wed-2008")
		' @notes	:	Used to display required date in the required format
    	' @END 
    	
	    Set sDateVal = GetObject("","Excel.Application")
		GetDateInRequiredFormat = sDateVal.Text(iDate,sDateFormat)
		Set xDate = Nothing
 End Function
'*********************************************************************************************************************************************************************
Function CloseBrowserByIndex(iIndex)
	    ' @HELP
	    ' @group	:	WebControls
	    ' @method	:   CloseBrowserByIndex(iIndex)
	    ' @returns	:   None 
	    ' @parameter:	iIndex : Browser Index
	    ' @notes	:   Used to Close Browser based on index.
	    ' @END
	    
	    Wait 2
		GetObjectDescriptions
		If Browser(oSesObjDescriptions,"CreationTime:="&iIndex).Exist  Then		  
			ClearAllCookies
			Browser(oSesObjDescriptions,"CreationTime:="&iIndex).Close
	   		ReportWriter "PASS","Browser"," Browser Closed",0
	  	Else
	  		iBrowClose = False
	   		'ReportWriter "FAIL","Browser"," Browser Not Closed",0
	  	End If
End Function
'*********************************************************************************************************************************************************************
Sub ClickLinkByIndex(iCreationtime,sLnkPropAndValue,iLnkIndex)
		' @HELP
		' @group	:	WebControls
		' @method	: 	ClickLink(iCreationtime,sLnkPropAndValue,sLinkText)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sLnkPropAndValue : Property & value of the Link
		' @notes	:	Used to Click the Link 
		' @END 
		
		sLinkText = Split(sLnkPropAndValue,":=")
		GetObjectDescriptions()
		If Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).Link(sLnkPropAndValue,"index:="&iLnkIndex).Exist Then
			Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).Link(sLnkPropAndValue,"index:="&iLnkIndex).Click()		
			ReportWriter "PASS","Link:" ,sLinkText(1) & "   is clicked",0
		Else
			ReportWriter "FAIL","Link:" ,sLinkText(1) & "   is not clicked",0
				End If			
End Sub
'*********************************************************************************************************************************************************************
Function GetRequiredDate(sInterval,iNumber,iDate,sDateFormat)
		' @HELP
		' @group	: 	WebControls
		' @method	:	GetRequiredDate(sInterval,iNumber,iDate)
   		' @returns	:   None
   		' @parameter:   sInterval	  : String expression that is the interval you want to add(for date "d",for month "m")
		' @parameter:   iNumber       : Numeric expression that is the number of interval you want to add
		' @parameter:   iDate	      : literal representing the date to which interval is added 
		' @notes	:	Used to display required date(we can display tomorrow date or else yesterday date or any date u want)
    	' @END

	    sdate = dateadd(sInterval,iNumber,iDate)
	    Set sDateVal = GetObject("","Excel.Application")
		sdate = sDateVal.Text(sdate,sDateFormat)
	    GetRequiredDate = sdate

 End Function
 '*********************************************************************************************************************************************************************
 Function VerifyWebPageExists(iBrowIndex,sWebElemPropAndVal1,sWebElemPropAndVal2,sWebPageText)
    	' @HELP
		' @group	:	 	WebControls
		' @method	:	 	VerifyWebPageExists(iBrowIndex,sWebElemPropAndVal1,sWebElemPropAndVal2,sWebPageText)
		' @returns	:	 	None
		' @parameter:       iBrowIndex          : Browser Creationtime
		' @parameter:   	sWebElemPropAndVal1 : Property and value of the webelement
		' @parameter:       sWebElemPropAndVal2 : Another Property and value of the webelement
		' @parameter:       sWebPageText        : Text of WebPage for verifying
		' @notes	:		This method is used to Verify Webpage exists or not
		' @END
		 
		GetObjectDescriptions ()
		If  Browser(oSesObjDescriptions,"Creationtime:=" & iBrowIndex).Page(oObjectDescriptions).WebElement(sWebElemPropAndVal1,sWebElemPropAndVal2).Exist  Then
			sActPageTxt = Browser(oSesObjDescriptions,"Creationtime:=" & iBrowIndex).Page(oObjectDescriptions).WebElement(sWebElemPropAndVal1,sWebElemPropAndVal2).GetROProperty("innertext")
			If InStr(sActPageTxt,sWebPageText) > 0 Then
				ReportWriter "PASS", "WebPage", sWebPageText + "   page exists",0
			Else
				ReportWriter "FAIL", "WebPage", sWebPageText + "  page does not exists",0
			End If
	    Else
	    	ReportWriter "FAIL", "WebElement", sWebPageText + "  not exists",0
	    End If	
End Function
 '*********************************************************************************************************************************************************************
Function  ClickImage(iCreationtime,sImgPropAndVal,iIndex,sText)
        ' @HELP
		' @group	: 	WebControls
		' @method	:	ClickImage(iCreationtime,sImgPropAndVal,iIndex,sText)
   		' @returns	:   None
   		' @parameter:   iCreationtime	  : Creation Time of the Browser
		' @parameter:   sImgPropAndVal: PropertyName & Value of the Image 
		' @parameter:   iIndex	      : Index of the Image 
	    ' @parameter:   sText         : Text of Image for reporting purpose
		' @notes	:	Used to Click specified Image in the specified browser
    	' @END 
    	
		GetObjectDescriptions ()
	    If Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).Image(sImgPropAndVal,"index:="&iIndex).Exist(.1) Then
			Browser(oSesObjDescriptions,"Creationtime:="&iCreationtime).Page(oObjectDescriptions).Image(sImgPropAndVal,"index:="&iIndex).Click
			ReportWriter "PASS","Image",sText & "  is clicked",0
		Else
			ReportWriter "FAIL","Image",sText & "  is not clicked",0
		End If
End Function

'*********************************************************************************************************************************************************************

Function SelectRadiobutton(iCreationtime,sPropAndVal,iIndex)
        ' @HELP
		' @group	: 	WebControls
		' @method	:	SelectRadiobutton(iBrowIndex,sImgPropAndVal,iIndex,sText)
   		' @returns	:   None
   		' @parameter:   iCreationtime	  : Creation Time of the Browser
		' @parameter:   sPropAndVal: PropertyName & Value of the Image 
		' @parameter:   iIndex	      : Index of the Image 
		' @notes	:	Used to Select the Radiobutton
    	' @END 

		GetObjectDescriptions ()		
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebRadioGroup(sPropAndVal).Exist Then
			Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebRadioGroup(sPropAndVal).Select iIndex
			ReportWriter "PASS", "Radiobutton", iIndex &"   is selected",0
		Else
		    ReportWriter "FAIL","List Box", sVal &"   is not selected",0
		End If
End Function

'*********************************************************************************************************************************************************************

Function GetTextFromWebElement(iCreationTime,sPropAndVal)
        ' @HELP
		' @group	: 	WebControls
		' @method	:	lement(iCreationTime,sPropAndVal)
   		' @returns	:   Text Present in the WebElement
   		' @parameter:   iCreationtime	  : Creation Time of the Browser
		' @parameter:   sPropAndVal: PropertyName & Value of the Image 
		' @notes	:	Used to get the WebElement Text
    	' @END 

         GetObjectDescriptions ()
         If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebElement(sPropAndVal).Exist Then
			sText = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebElement(sPropAndVal).getROProperty("innertext")
			ReportWriter "PASS", "WebElemt", sText &"   is Present",0
		Else
		    ReportWriter "FAIL","WebElemt", sText &"   is not Present",0
		End If
		GetTextFromWebElement = sText
End Function


'*********************************************************************************************************************************************************************
Function ClickTableCellbyContent(iCreationtime,sTblPropAndVal,iTblIndex,sValue,iColumnNo)
		
		' @HELP
		' @group	: 	WebControls
		' @method	:	CompareDataInWebTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,iColID,sData2Verify,sText)
   		' @returns	:   Click on the Cell against to the given data
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable		
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   iColumnNo: Coloumn Number
		' @parameter:   sValue : Value to be Verify
		' @END
		
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
			iRowCount = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).RowCount
		    For i = 1 To iRowCount
		      sActValue = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).GetCellData(i,iColumnNo)
			  If sActValue = sValue Then 
			  	Set sLink=Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).ChildItem(i,iColumnNo,"Link",0)
				sLink.Click
				ReportWriter "PASS","WebTable" , "Clicked on the Given data in the Table",0
				Exit For
			  End If
		    Next 		
		Else
    		ReportWriter "FAIL", "WebTable" , "WebTable doesn't exist",0
    	End If   
    	
End Function
'*********************************************************************************************************************************************************************
Function GetWebTableRowCount(iCreationtime,sTblPropAndVal,iTblIndex)
		' @HELP
		' @group	: 	WebControls
		' @method	:	GetWebTableRowCount(iCreationtime,sTblPropAndVal,iTblIndex)
   		' @returns	:   The WebTable RowCount
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable		
		' @parameter:   iTblIndex	   :  index of WebTable
		' @END
	GetObjectDescriptions ()	
	If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist Then
	   iRowCount = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).RowCount
	End If
	GetWebTableRowCount = iRowCount
End Function
'*********************************************************************************************************************************************************************
Function SetDataInWebEditByIndex(iCreationtime,sEditPropAndValue,iIndex,sValue)
'    	 @HELP
'		 @group	    : 	WebControls
'		 @method	: 	SetDataInWebEdit(iCreationtime,sEditPropAndValue,sValue)
'		 @returns	: 	None
' 		 @parameter :   iCreationtime  :  Creationtime of Browser
'		 @parameter :	sEditPropAndValue: Property Name & Value of the Object
'		 @parameter :	sValue: Value to be set In the WebEdit
'		 @parametre :   iIndex: Index Value of WebEdit
'		 @notes	    : 	This method is used to set data In WebEdit
'		 @END
		
		GetObjectDescriptions ()
		sEditValue = Split(sEditPropAndValue,":=")
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebEdit(sEditPropAndValue,"Index:="&iIndex).Exist(1) Then
			Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebEdit(sEditPropAndValue,"Index:="&iIndex).Set sValue
			ReportWriter "PASS", "Edit Box" , sValue & "  value is set in  " & sEditValue(1),0   		
    	Else
    		ReportWriter "FAIL", "Edit Box" , sValue & "  value is not set in  " & sEditValue(1),1
    	End If 
End Function
'*********************************************************************************************************************************************************************
Sub CheckWebElement(iCreationtime,sWebElementTxt,sPassStat,sFailStat)
    
    ' @HELP
	' @class	:	clsWebControls
	' @method	:	CheckWebEliment(sWebElementTxt,sPassStat,sFailStat)
	' @returns	:	None
	' @parameter:	sWebElementTxt: Innertext of WebElement
	' @parameter:	sPassStat: Pass - object exists
	' @parameter:	sFailStat: Fail - object does not exists
	' @notes	:	Verifies Whether A WebElement Exists or Not
	' @END
			
	On Error Resume Next
	
	Found=0

	Set oDesc = Description.Create() 
		oDesc("micclass").Value = "WebElement"
		Set Lists = Browser("micClass:=Browser","CreationTime:="&iCreationtime).Page("micClass:=Page").ChildObjects(oDesc) 
		NumberOfLists = Lists.Count() 
		For Itr = 2 To NumberOfLists-1 
			If  Lists(Itr).Exist  Then
					sText = Trim(Lists(Itr).GetROProperty("innertext") )
					If  LCase(sText) = LCase(Trim(sWebElementTxt)) Then
						Found=1
						Exit For
					End If
			End If
		Next
			
	If Found=1 Then
'		msgbox "Found : 1"
			ReportWriter "PASS",sWebElementTxt&" Tag Verification",sPassStat,0
		Else
'		msgbox "Found : 0"
			ReportWriter "FAIL",sWebElementTxt & " Tag Verification",sFailStat,0
	End If
	
	On Error GoTo 0
End Sub
'*********************************************************************************************************************************************************************
Function CheckImgExist(iCreationtime,sImgProp,sImgValue,sPassStatus,sFailStatus)
		' @HELP
   		' @class	:  	clsWebControls 
   		' @method	:   CheckImgExist(sImgProp,sImgValue,sPassStatus,sFailStatus)
   		' @returns	:   None
   		' @parameter:	sImgProp: property of Image
   		' @parameter:	sImgValue:property value of Image
   		' @parameter:	sPassStatus: Pass - object exists
		' @parameter:	sFailStatus: Fail - object do not exists 		
   		' @notes	:   Checks whether an Image is existing or not.
   		' @END
   		
     GetObjectDescriptions()
     sImgFile = sImgProp & ":=" & sImgValue       
     If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).Image(sImgFile).Exist(5) Then
     	ReportWriter "PASS","Image/Page Check",sPassStatus,0
     Else
     	ReportWriter "FAIL","Image/Page Check",sFailStatus,1
     End If       
    End Function
'*********************************************************************************************************************************************************************
Sub CheckTxtInWebEliment(iCreationtime,sWebElementTxt,sPassStat,sFailStat)
	    ' @HELP
		' @class:	 	clsWebControls
		' @method:	 	CheckTxtInWebEliment(sWebElementTxt,sPassStat,sFailStat)
		' @returns:	 	
		' @notes:	 	Verifies Whether Specified Text is available in WebElement Text or Not
		' @END		
		
			Found=0
			For Each objElement In Browser("micClass:=Browser","CreationTime:="&iCreationtime).Object.document.all
				If InStr(Trim(LCase(objElement.innerText)),Trim(LCase(sWebElementTxt)))>0 Then
					Found=1
					Exit For
				End If
			Next
			
		If Found=1 Then
			ReportWriter "PASS",sWebElementTxt&" Tag Verification",sPassStat,0
		Else
			ReportWriter "FAIL",sWebElementTxt&" Tag Verification",sFailStat,1
		End If
End Sub
'*********************************************************************************************************************************************************************
Function RetrieveValInWebTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,iColID,sText)
		
		' @HELP
		' @group	: 	WebControls
		' @method	:	RetrieveValInWebTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,iColID,sText)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   iRowID 		   :  Row ID in which data is to be compared
		' @parameter:   iColID		   :  Column ID in which data is to be compared
		' @parameter:   sText		   :  Text/Column name for reporting purpose
		' @notes	:	Retrieves data from Row and col
		' @END
		
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
			sValue = Trim(Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).GetCellData(iRowID,iColID))
        Else
    		ReportWriter "FAIL", "WebTable" , "WebTable doesn't exist",0
    	End If 
    	RetrieveValInWebTableByRowAndColID = sValue
End Function
'*********************************************************************************************************************************************************************
Sub SelectCheckBox(iCreationtime,sBoxName)
		' @HELP
		' @class:	 	clsWebControls
		' @method:	 	SelectCheckBox(sBoxName)
		' @returns:	 	
		' @notes:	 	Checks a specified Check Box
		' @END
		If Browser("micClass:=Browser","CreationTime:="&iCreationtime).Page("micClass:=Page").WebCheckBox("name:="&sBoxName).Exist Then
			Browser("micClass:=Browser","CreationTime:="&iCreationtime).Page("micClass:=Page").WebCheckBox("name:="&sBoxName).Set "ON"
			ReportWriter "PASS","CheckBox",sBoxName&" CheckBox Cheacked ",0
		Else
			 ReportWriter "FAIL","CheckBox",sBoxName& "CheckBox  Is not available ",0
			 ExitRun
		End If 
	End Sub
'*********************************************************************************************************************************************************************
Function ClickLinkInTableByRowAndColIDAndGetText(iCreationtime,sTblPropAndVal,iTblIndex,iRowId,iColID)

		' @HELP
		' @group	: 	WebControls
		' @method	:	ClickLinkInTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowId,iColID)
   		' @returns	:   None
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   iRowID 		   :  Row ID in which data is to be compared
		' @parameter:   iColID		   :  Column ID in which data is to be compared
		' @notes	:	Clicks link present in WebTable based on specifed row and column ID
		' @END
		
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
			Set oDesc = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).ChildItem(iRowId,iColID,"Link",0)
			oDesc.Click
		   sOrder =  oDesc.GetROProperty("text")
           ReportWriter "PASS","Link","Clicked on the Link  " & oDesc.GetROProperty("name"),0
		Else
			ReportWriter "FAIL","WebTable","WebTable Doesn't exist",0
		End If	
		ClickLinkInTableByRowAndColIDAndGetText = sOrder
End Function
'*********************************************************************************************************************************************************************
Function VerifyLessThanOrGreaterThan (sExp,sAct,sVerification,sExpText,sActText)
		' @HELP
   		' @group	:  	functions
   		' @method	:   VerifyLessThanOrGreaterThan (sExp,sAct,sVerification,sText)
   		' @returns	:   None
   		' @parameter:	sExp          : Expected Value
   		' @parameter:	sAct          : Actual Value
   		' @parameter:	sVerification : Pass what to verify less than or greater than..
   		' @parameter:	sText         : Text for reporting purpose
   		' @notes	:   Compares two Values wheather Expected value is less than actual or not
   	   	' @END
   	   	
   	   	Select Case sVerification   	   	
   	   	Case "<"
	   	   	If Abs(sExp) < Abs(sAct) Then
	   	   	ReportWriter "PASS","Data",sExpText & " is less than " & sActText,0
	   	   	Else
	   	   	ReportWriter "FAIL","Data",sActText & " is not less than " & sActText,1
	   	    End If
   	    Case ">"   	    
	   	   	If Abs(sExp) > Abs(sAct) Then
	   	   	ReportWriter "PASS","Data",sExpText & " is Greater than " & sActText,0
	   	   	Else
	   	   	ReportWriter "FAIL","Data",sExpText & " is not Greater than " & sActText,1
	   	   	End if
   	   End Select
 End Function 
 
 '****************************************************************************************************************************************************************************************
 Function GetRowID(iCreationtime,sTblPropAndVal,iIndex,sVal,iCol)
		' @HELP
		' @class	:	clsOrders
		' @method	:	GetRowCountID(sTblPropAndValue,iIndex,iItemNumber)
		' @returns	:	ID of required column 
		' @parameter:   sVal:Cell Value based on which, rowID will be written 
		' @parameter:   sTblProp : Property name & value of WebTable
		' @parameter:   iCol : Column ID (get this ID using "Catalog.GetColumnID" - function)
		' @parameter:   iIndex   : WebTable index
		' @notes	:	Retrieves Row ID in a WebTable based on given input
		' @END
		GetObjectDescriptions ()
		Wait(1)
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="& iIndex).Exist Then
		 Set oTab = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="& iIndex)
		  iRowCnt = oTab.GetROProperty("rows")
		    For i = 2 to iRowCnt
		     If Len(sVal) = 0 Then
		      If Trim(oTab.GetCellData(i,iCol)) <> sVal Then
			    GetRowID = i
			      Exit For
			    	End If
		    	      Else
					 If Lcase(Trim(oTab.GetCellData(i,iCol))) = Trim(Lcase(sVal))  Then
			        GetRowID = i
			       Exit For
				 End If
				End If
			  Next
		    Else
		   ReportWriter "FAIL","WebTable","Search results table is not found",0
		 End If
  End Function 
  
'*********************************************************************************************************************************************************************
  
  Function VerifyWebElement(iCreationtime,sWebElemPropAndVal1,sWebElemPropAndVal2,sWebElemText)
    	' @HELP
		' @group	:	 	WebControls
		' @method	:	 	VerifyWebElement(sWebElemPropAndVal1,sWebElemPropAndVal2,sWebElemText)
		' @returns	:	 	None
		' @parameter:   	sWebElemPropAndVal1 : Property and value of the webelement
		' @parameter:       sWebElemPropAndVal2 : Another Property and value of the webelement
		' @parameter:       sWebElemText        : Text of Webelement for reporting purpose
		' @notes	:		This method is used to VerifyWebElement
		' @END
		 
		GetObjectDescriptions ()
		If  Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebElement(sWebElemPropAndVal1,sWebElemPropAndVal2).Exist(.1)  Then
			ReportWriter "PASS", "WebElement", sWebElemText + "   text exists",0
			'VerifyWebElement = 1
	    Else
	    	ReportWriter "FAIL", "WebElement", sWebElemText + "  text not exists",0
	    End If	
End Function

'****************************************************************************************************************************************************************

Function GetCellValue(iCreationtime,sTblProp,iIndex,iRow,iColID)

		' @HELP
		' @class	:	clsCatalog
		' @method	:	GetSearchResultsCellData(sTblProp,iIndex,sColName)
		' @returns	:	Search results cell data 
		' @parameter:   sColName : ColName of the required cell
		' @parameter:   sTblProp : Property name & value of WebTable
		' @parameter:   iIndex   : WebTable index
		' @notes	:	Retrieves cell data from Search Results page
		' @END
		GetObjectDescriptions ()
		If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblProp,"index:="& iIndex).Exist Then		
			sCellValue = Trim(Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblProp,"index:="& iIndex).GetCellData(iRow,iColID))
			ReportWriter "PASS","WebElement","Value in cell is" &  sCellValue,0
			GetCellValue   =  sCellValue
		Else
			ReportWriter "FAIL","WebTable","Search results table Doesn't Exist",0
		End If	
End Function

'**********************************************************************************************************************************************************************************

Sub VerifyEditBoxVal(iCreationtime,sObjPropNameAndValue,sExpTxt)
	        ' @HELP
			' @group    :	WebControls
			' @method   :	VerifyEditBoxVal(sObjPropNameAndValue,sExpTxt)
			' @returns  :	None
			' @parameter:	sObjName: Name of WebEdit Object
			' @parameter:	sExpTxt: Value of Editbox to Verify
			' @notes:	 	Verifies Whether a Edit Box has the expected value in it or not
			' @END
			
			
			GetObjectDescriptions ()	
			ActTxt=Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebEdit(sObjPropNameAndValue).GetRoProperty("value")
			sObjName = Split(sObjPropNameAndValue,":=")
			If Trim(sExpTxt) = "" Then
				If Trim(sExpTxt) = Trim(ActTxt) Then
					ReportWriter "PASS","EditBox Text Verification","Null Value or Value got cleared in "&sObjName(1),0
		  		Else
					ReportWriter "FAIL","EditBox Text Verification","Value is not cleared from "&sObjName(1),1
		  		End If 
			ElseIf Not IsNumeric(ActTxt) then
		    	If Trim(LCase(sExpTxt))=Trim(LCase(ActTxt)) Then
					ReportWriter "PASS","EditBox Text Verification","Specified Text "&sExpTxt&" Displayed in Edit Box "&sObjName(1),0
		  		Else
					ReportWriter "FAIL","EditBox Text Verification","Specified Text "&sExpTxt&" Not Displayed in Edit Box "&sObjName(1),1
		  		End If   
			Else
				If Trim(Int(sExpTxt))=Trim(Int(ActTxt)) Then
					ReportWriter "PASS","EditBox Text Verification","Specified Text "&sExpTxt&" Displayed in Edit Box "&sObjName(1),0
			  	Else
					ReportWriter "FAIL","EditBox Text Verification","Specified Text "&sExpTxt&" Not Displayed in Edit Box "&sObjName(1),1
		  		End If   
			End If		  
End Sub
'**********************************************************************************************************************************************************************************
  
  Function GetWebTableInnertext(sSubStringValueOfInnertext)
	'@HELP
	'@group		:	WebControls
	'@method	:	GetWebTableInnertext(sSubStringValueOfInnertext)
	'@parameter :    sSubStringValueOfInnertext : A sub string of webTable innertext value	
	'@End
	
 GetObjectDescriptions ()
 Found=0
 sInnertext =""
 set oWebTable = Description.Create
 oWebTable("name").value="WebTable"
   Set WebTables = Browser(oSesObjDescriptions).Page(oObjectDescriptions).ChildObjects(oWebTable)
    For i = 0 to WebTables.count-1
	If not Trim(WebTables(i).GetROProperty("innertext")) = "" then
	If  Instr(Trim(WebTables(i).GetROProperty("innertext")),sSubStringValueOfInnertext) > 0 Then
    sInnertext= Trim(WebTables(i).GetROProperty("innertext"))
	Found = 1
	Exit for
	End If
	End if
    Next
	If Found=1 Then
		ReportWriter "PASS","WebTableInnertext", "'" & sSubStringValueOfInnertext  & " is found",0
		GetWebTableInnertext = sInnertext
	Else
		ReportWriter "FAIL","WebTableInnertext", "'" & sCheckBoxValue  & " is not found",0
		GetWebTableInnertext = ""
	End If

End Function

'*************************************************************************************************************************************************************************************

Sub ClickCheckBoxByValue(iCreationtime,sCheckBoxValue)
		'@HELP
		'@group		:	WebControls
		'@method	:	ClickCheckBoxByValue(sCheckBoxValue)
		'@parameter :   sCheckBoxValue : CheckBox value
		'@notes	    :	No need to specify the entair value of the CheckBox
		'@End 
	GetObjectDescriptions ()
	Found=0
	set oWebCheckBox = Description.Create
	oWebCheckBox("micclass").value="WebCheckBox"
	set  CheckBoxs=Browser(oSesObjDescriptions,"CreationTime:=" & iCreationtime).Page(oObjectDescriptions).ChildObjects(oWebCheckBox)
	For i = 0 to   CheckBoxs.count-1
		If Instr(CheckBoxs(i).GetROProperty("value"),sCheckBoxValue) > 0 then
			CheckBoxs(i).set "ON"
			wait 4
			Found=1
		End if 
	Next
	If Found=1 Then
		ReportWriter "PASS","ClickCheckBox",sCheckBoxValue  & " is selected",0
	Else
		ReportWriter "FAIL","ClickCheckBox",sCheckBoxValue  & " is not selected",0
	End If
	
End Sub
	

'*********************************************************************************************************************************************************************

Function ClickOnWebElementByIndex(iCreationTime,sPropValue,sIndex)
'		 @HELP
'	     @group	    :	WebControls
'		 @method	:	ClickOnWebElement(sPropValue)
'		 @parameter :   sPropValue : property name and value of the WebElemnt.
'		 @notes	    :   Click on WebElement 
'		 @END
			
	    GetObjectDescriptions ()	
		If Browser(oSesObjDescriptions,"CreationTime:=" & iCreationTime).Page(oObjectDescriptions).WebElement(sPropValue,"index:=" & sIndex).Exist  Then
		   sText=Browser(oSesObjDescriptions,"CreationTime:=" & iCreationTime).Page(oObjectDescriptions).WebElement(sPropValue,"index:=" & sIndex).GetRoProperty("innertext")
		   Browser(oSesObjDescriptions,"CreationTime:=" & iCreationTime).Page(oObjectDescriptions).WebElement(sPropValue,"index:=" & sIndex).click
		   	ReportWriter "PASS","WebElement","WebElement " &sText&" is clicked",0
        Else
        	ReportWriter "FAIL","WebElement","WebElement " &sText&" does not exist",0
			end if
 End Function
 '********************************************************************************************************************************************************************************
Function ClickOnWebElementInTblByIndex(iCreationTime,sTblPropnVal,sPropValue,sIndex)
'		 @HELP
'	     @group	    :	WebControls
'		 @method	:	ClickOnWebElement(sPropValue)
'		 @parameter :   sPropValue : property name and value of the WebElemnt.
'		 @notes	    :   Click on WebElement 
'		 @END
			
	    GetObjectDescriptions ()	
		If Browser(oSesObjDescriptions,"CreationTime:=" & iCreationTime).Page(oObjectDescriptions).WebTable(sTblPropnVal).WebElement(sPropValue,"index:=" & sIndex).Exist  Then
		   sText=Browser(oSesObjDescriptions,"CreationTime:=" & iCreationTime).Page(oObjectDescriptions).WebTable(sTblPropnVal).WebElement(sPropValue,"index:=" & sIndex).GetRoProperty("innertext")
		   Browser(oSesObjDescriptions,"CreationTime:=" & iCreationTime).Page(oObjectDescriptions).WebTable(sTblPropnVal).WebElement(sPropValue,"index:=" & sIndex).click
		   	ReportWriter "PASS","WebElement","WebElement " &sText&" is clicked",0
        Else
        	ReportWriter "FAIL","WebElement","WebElement " &sText&" does not exist",0
			end if
 End Function
 '********************************************************************************************************************************************************************************
 
 Function CheckWebElementInnertext(iCreationtime,sWebElemPropAndVal1,sWebElemPropAndVal2,sWebElemText,sPassStat,sFailStat)
    	' @HELP
		' @group	:	 	WebControls
		' @method	:	 	VerifyWebElement(sWebElemPropAndVal1,sWebElemPropAndVal2,sWebElemText)
		' @returns	:	 	None
		' @parameter:   	sWebElemPropAndVal1 : Property and value of the webelement
		' @parameter:       sWebElemPropAndVal2 : Another Property and value of the webelement
		' @parameter:       sWebElemText        : Text of Webelement for reporting purpose
		' @notes	:		This method is used to VerifyWebElement
		' @END
		 
		GetObjectDescriptions ()
		If  Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebElement(sWebElemPropAndVal1,sWebElemPropAndVal2).Exist(.1)  Then
		sText=Browser(oSesObjDescriptions,"CreationTime:=" & iCreationTime).Page(oObjectDescriptions).WebElement(sWebElemPropAndVal1,sWebElemPropAndVal2).GetRoProperty("innertext")
		
		If  LCase(sText) = LCase(Trim(sWebElemText)) Then
		
			ReportWriter "PASS",sWebElemText &" Tag Verification",sPassStat,0
			'VerifyWebElement = 1
	    Else
	    	ReportWriter "FAIL",sWebElemText&" Tag Verification",sFailStat,1
	    	
	    	End If
	    End If	
End Function

'*****************************************************************************************************************************************************************

Function VerifyWebElementDoeNotExists (iBrowserIndex,sWebElemText,sPassStat,sFailStat)

	Found = 0

   Set oDesc = Description.Create
   oDesc("micclass").value = "WebElement"
  ' oDesc("class").value = "sn"
   oDesc("html tag").value = "DIV"
   Set aObjts =  Browser("micclass:=Browser","CreationTime:=" & iBrowserIndex).page("micclass:=page").ChildObjects(oDesc)

   If aObjts.Count>0 Then
		For LCount = 0 to aObjts.Count-1
			sActual = aObjts(LCount).GetRoProperty("innertext")
			'MsgBox sActual
			If  LCase(sActual)  = LCase(Trim(sWebElemText)) Then
		       	Found=1
				Exit For
			End If
		Next
	End If

If Found=1 Then
	ReportWriter "FAIL", sWebElemText&" Tag Verification",sFailStat,1
Else
	ReportWriter "PASS",sWebElemText &" Tag Verification",sPassStat,0
End If
End Function
'*************************************************************************************************************************************************************************************************************

Function VerifyWebTableNotExist(iCreationtime,sTblPropAndVal,iTblIndex,sText)
		' @HELP
		' @group	: 	WebControls
		' @method	:	GetWebTableRowCount(iCreationtime,sTblPropAndVal,iTblIndex)
   		' @returns	:   The WebTable RowCount
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable		
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   sText		   :  Reporting message
		' @END
	GetObjectDescriptions ()	
	If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist Then
	   ReportWriter "FAIL", "WebTable",sText & " is Created",0
	  Else
	  	ReportWriter "PASS", "WebTable",sText & " is not Created",0
	End If
End Function
'*************************************************************************************************************************************************************************************************************

Function GetRowIDOfTextPresentInWebTable(iCreationtime,sTblPropAndVal,iTblIndex,sText)
		' @HELP
		' @group	: 	WebControls
		' @method	:	GetRowIDOfTextPresentInWebTable(iCreationtime,sTblPropAndVal,iTblIndex)
   		' @returns	:   The WebTable RowCount
   		' @parameter:   iCreationtime  :  Creationtime of Browser
		' @parameter:   sTblPropAndVal :  Property name and value of WebTable		
		' @parameter:   iTblIndex	   :  index of WebTable
		' @parameter:   sText		   :  text present in webtable
		' @END
	GetObjectDescriptions ()	
	If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist Then
	   GetRowIDOfTextPresentInWebTable = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).GetRowWithCellText(sText)
	  Else
	  	ReportWriter "FAIL", "WebTable",sText & "WebTable is not present",0
	End If
End Function

'********************************************************************************************************************************************************************************************************
Function ClickDialogButton()
	Set oDesc = Description.Create()
	oDesc("micclass").value = "Dialog"
   Set oDialog = Desktop.ChildObjects(oDesc)
	 If  oDialog.Count >= 1 Then
		Set oShell = CreateObject("WScript.Shell")
				oDialog(0).Activate
				'oShell.SendKeys "{TAB}"
			   Wait 1
			 oShell.SendKeys "~"
			 Wait 1
	Else
		ClickDialogButton = False	
	End If		   
End Function
'************************************************************************************************************************************************************************************************************************
Sub VerifyWebElementVal(iCreationtime,sObjPropNameAndValue,sExpTxt)
	        ' @HELP
			' @group    :	WebControls
			' @method   :	VerifyWebElementVal(sObjPropNameAndValue,sExpTxt)
			' @returns  :	None
			' @parameter:	sObjName: Name of WebElementObject
			' @parameter:	sExpTxt: Value of WebElement to Verify
			' @notes:	 	Verifies Whether aWeb Element has the expected value in it or not
			' @END
		
			GetObjectDescriptions ()
			If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebElement(sObjPropNameAndValue).exist Then	
               ActTxt=Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebElement(sObjPropNameAndValue).GetRoProperty("innertext")
               'MsgBox ActTxt
            End If
			sObjName = Split(sObjPropNameAndValue,":=")
			If Trim(sExpTxt) = "" Then
				If Trim(sExpTxt) = Trim(ActTxt) Then
					ReportWriter "PASS","WebElement Verification","Null Value is Present in "&sObjName(1),0
		  		Else
					ReportWriter "FAIL","WebElement Verification","Value is prsent in "&sObjName(1),1
		  		End If 
		  	ElseIf  IsDate(ActTxt) Then
			   ActTxt=GetDateInRequiredFormat(ActTxt,MM_DD_YYYY_FORMAT)
		    	If Trim(LCase(sExpTxt))=Trim(LCase(ActTxt)) Then
				
				ReportWriter "PASS","Web Element Text Verification","Specified Text "&sExpTxt&" Displayed in Web Element "&sObjName(1),0
		  		Else
				ReportWriter "FAIL","Web Element Text Verification","Specified Text "&sExpTxt&" Not Displayed in Web Element "&sObjName(1),1
		  		End If   
			Else
				If Trim(int(sExpTxt))=Trim(Int(ActTxt)) Then
					ReportWriter "PASS","Web Element Text Verification","Specified Text "&sExpTxt&" Displayed in Web Element "&sObjName(1),0
			  	Else
					ReportWriter "FAIL","Web Element Text Verification","Specified Text "&sExpTxt&" Not Displayed in Web Element "&sObjName(1),1
		  		End If   
			End If		  
 End Sub
'************************************************************************************************************************************************************************************************************************
Function RetrieveColumnNumberFromWebTableByRowIDAndColumnName(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,sColumnName)
	
	' @HELP
	' @group	: 	WebControls
	' @method	:	RetrieveColumnNumberFromWebTableByRowIDAndTextPresent(iCreationtime,sTblPropAndVal,iTblIndex,iRowID,sColumnName)
   	' @returns	:   Column Number
   	' @parameter:   iCreationtime  :  Creationtime of Browser
	' @parameter:   sTblPropAndVal :  Property name and value of WebTable
	' @parameter:   iTblIndex	   :  index of WebTable
	' @parameter:   iRowID 		   :  Row ID in which data is to be compared
	' @parameter:   sColumnName	   :  Name of the Column whose Column Number is to be retrieved
	' @notes	:	Returns the Column Number based on Column Name 
	' @END
	
	GetObjectDescriptions ()
	If Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
		sCols = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).ColumnCount(iRowID)
		For iCol = 1 To sCols
			sActualColumnName = Browser(oSesObjDescriptions,"CreationTime:="&iCreationtime).Page(oObjectDescriptions).WebTable(sTblPropAndVal,"index:="&iTblIndex).GetCellData(1,iCol)
			If LCase(Trim(sActualColumnName)) = LCase(Trim(sColumnName)) Then 
				ReqColumnNumber = iCol
				Exit For 
			End If 
		Next
	End If
	RetrieveColumnNumberFromWebTableByRowIDAndColumnName = ReqColumnNumber
End Function  
 
 '************************************************************************************************************************************************************************************************************************
Function VerifyValueByFieldNameinWebTable(sFieldName,sFieldValue,iCreationTime)

	' @HELP
	' @group    : 
	' @method   : VerifyValueByFieldNameinWebTable()
	' @returns  : None
	' @parameter: sFieldName : Field Name of the Table 
	' @parameter: sFieldValue : Field value of the Table of which the Data is to be read
	' @parameter: iCreationTime	: Creation Time of the Browser
	' @notes    : To Verify whether a cell corresponding to a specified cell contains valid data or not
	' @example	: VerifyValueByFieldNameinWebTable("Order Id","1456784",0)
	' @END
		
		On Error Resume Next 
		
		Set BaseWindow = Browser("micclass:=Browser","creationtime:="& iCreationTime).Page("micclass:=page")
				
		Set oDesc = Description.Create() 
		oDesc("micclass").Value = "WebTable" 
		
		Set Lists = BaseWindow.ChildObjects(oDesc) 
		
		NumberOfLists = Lists.Count()
		sTag = 1
		For i = 0 to NumberOfLists - 1  
			rows = Lists(i).rowcount()
			For r = 1 to rows
				cols = Lists(i).ColumnCount(r)
				For C = 1 to cols
					sTextToSearch = Lists(i).GetCellData(r,c)
					If (trim(sTextToSearch)= Trim(sFieldName)) Then
						sTag = 0
                        sWebElementVal = Lists(i).GetCellData(r,c+1)
                 		If InStr(UCase(Trim(sWebElementVal)),UCase(Trim(sFieldValue))) > 0 Or InStr(UCase(Trim(sFieldValue)),UCase(Trim(sWebElementVal)))>0 Then 
						'Modified by sudheer
                        'If UCase(Replace(sWebElementVal," ",""))=UCase(Replace(sFieldValue," ","")) Then 
                    		ReportWriter "PASS", ""&sFieldName&"Verification",""&sFieldName& " Verification Passed. Expected : "&sFieldValue&" Actual Value : = " &sWebElementVal,0
							VerifyValueByFieldNameinWebTable = sWebElementVal
							Exit Function
						Else 
							ReportWriter "FAIL", ""&sFieldName&" Verification", ""&sFieldName&" Verification Failed. Expected : "&sFieldValue&" Actual Value := " &sWebElementVal,1
							Exit Function
						End If
                    End If
				Next
			Next
		Next
		If 	sTag = 1 Then
			ReportWriter "FAIL", ""&sFieldName&" Verification", sFieldName&" filed does not exist",0
		End If 
End Function
'***************************************************************************************************************************************************************************************************************************
Function ClickOnWebObjectsByIndex(sObjType,sObjPropAndValue,iCreationTime,iIndex)

	' @HELP
	' @group                		: WebControls
	' @method               		: ClickOnWebObjectsByIndex(sObjType,sObjPropAndValue,iCreationTime,iIndex)
	' @returns              		: None
	' @parameter: sObjType 			: Type of Object(webbutton,Link,Image)
	' @parameter: sObjPropAndValue 	: Object Property & Value (ex: micClass:=Browser)
	' @parameter: iCreationTime		: Browser creation time (ex: 1,2 ...)
	' @parameter: iIndex			: Object Index & Value (ex: 1,2 ...)
	' @notes                		: To get Default Value of any Object (Webedit,Weblist,Webradiogroup,...)
	' @END
	
	On Error Resume Next 
	sObjType = LCase(Trim(sObjType))
	
	If iCreationTime = 0 Then
		Set BaseWindow = Browser("micClass:=Browser").Page("micClass:=Page")
	Else
		Set BaseWindow = Browser("micClass:=Browser","CreationTime:=" &iCreationTime).Page("micClass:=Page")
	End If 
	
	Select Case sObjType
	
	Case "webbutton"
	
			sWebButton = Split(sObjPropAndValue,":=")
			If BaseWindow.WebButton(sObjPropAndValue,"Index:="&iIndex).Exist Then
				BaseWindow.WebButton(sObjPropAndValue,"Index:="&iIndex).Highlight
				BaseWindow.WebButton(sObjPropAndValue,"Index:="&iIndex).Click
				ReportWriter "PASS","WebButton: "&sWebButton(1),"Pass:  Clicked On Button: " &sWebButton(1),0
			Else
				ReportWriter "FAIL","WebButton: "&sWebButton(1),"Fail:  Not Clicked On Button: " &sWebButton(1),1 
			End If
			
	Case "link"
	
		    sLinkName = Split(sObjPropAndValue,":=")
			If BaseWindow.Link(sObjPropAndValue,"Index:="&iIndex).Exist Then
				BaseWindow.Link(sObjPropAndValue,"Index:="&iIndex).Click()
				ReportWriter "PASS" ,"WebLink: "& sLinkName(1) ,"Pass:  '"&sLinkName(1)&"' Link is Clicked.",0
			Else
				ReportWriter "FAIL" ,"WebLink: "& sLinkName(1) ,"Fail:  '"&sLinkName(1)&"' Link is Not Clicked.",0
			End If	
			
	Case "image"
	
			sImageLink = Split(sObjPropAndValue,":=")
		    If BaseWindow.Image(sObjPropAndValue,"Index:="&iIndex).Exist Then
				BaseWindow.Image(sObjPropAndValue,"Index:="&iIndex).Click
				ReportWriter "PASS" ,"Image: "& sImageLink(1) ,"Pass:  '"&sImageLink(1)&"' Image is Clicked.",0
			Else
				ReportWriter "FAIL" ,"Image: "& sImageLink(1) ,"Fail:  '"&sImageLink(1)&"' Image is Not Clicked.",0
			End If
			
	Case "webcheckbox"
	
			sWebCheckBox = Split(sObjPropAndValue,":=")
		    If BaseWindow.WebCheckBox(sObjPropAndValue,"Index:="&iIndex).GetROProperty("checked") = 0 Then
			   BaseWindow.WebCheckBox(sObjPropAndValue,"Index:="&iIndex).Click
			   ReportWriter "PASS" ,"CheckBox: "& sWebCheckBox(1) ,"Pass:  '"&sWebCheckBox(1)&"' Check Box Selected.",0
			 Else
			   ReportWriter "FAIL" ,"CheckBox: "& sWebCheckBox(1) ,"Fail:  '"&sWebCheckBox(1)&"' Check Box not Selected..",0
			 End If
	
	Case "webradiogroup"
			sRadio = Split(sObjPropAndValue,":=")
			If BaseWindow.WebRadiogroup(sObjPropAndValue,"Index:="&iIndex).Exist Then
	   			BaseWindow.WebRadiogroup(sObjPropAndValue,"Index:="&iIndex).Click
				ReportWriter "PASS" ,"Web Radio: " & sRadio(1) ,"Pass:  'Selected the " & sRadio(1),0
				Exit Function
				Else
				ReportWriter "FAIL" ,"Web Radio: " & sRadio(1) ,"Fail:  'Data not Selected for : "& sRadio(1)&"'.",0
			End If 
	
	Case Else
	
			 ReportWriter "FAIL", "Web Object  : " &sEdit(1) , "'"& sValue & "'  " & "  value not set in  object " & "'"  & sEdit(1)& "'",0
			 ReportWriter "FAIL", "Default Value of "&sObjType&": ","Fail:  'Default Value of "&sObjType&": not in case selection'",0
	End Select
	
	If Err <> 0 Then
		ReportWriter "FAIL","UnKnown Error","Fail:  '"&Err.Description&"'.",0
	End If 
	Err.Clear
	
End Function
'---------------------------------------------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------------------------------

Function SetEditOrSelectListOrRadio(sObjType,sObjPropAndValue,sValue,iCreationTime)

	' @HELP
	' @group                		: WebControls
	' @method               		: ClickOnWebObjectsByIndex(sObjType,sObjPropAndValue,sValue,iCreationTime)
	' @returns              		: None
	' @parameter: sObjType 			: Type of Object(webbutton,Link,Image)
	' @parameter: sObjPropAndValue  	: Object Property & Value (ex: micClass:=Browser)
	' @parameter: sValue			: Value to be Set for respective object
	' @parameter: iCreationTime		: Browser creation time (ex: 1,2 ...)
	' @notes                		: To get Default Value of any Object (Webedit,Weblist,Webradiogroup,...)
	' @END


	On Error Resume Next                                      
	sObjType = LCase(Trim(sObjType))
	
	If iCreationTime = 0 Then
		Set BaseWindow = Browser("micClass:=Browser").Page("micClass:=Page")
	Else
		Set BaseWindow = Browser("micClass:=Browser","CreationTime:=" &iCreationTime).Page("micClass:=Page")
	End If 

	
	Select Case sObjType
	
	Case "webedit"

			sEdit = Split(sObjPropAndValue,":=")
			
			If BaseWindow.WebEdit(sObjPropAndValue).Exist(3) Then
		
			 	BaseWindow.WebEdit(sObjPropAndValue).Set sValue
'				If sValue = "" Then
'					ReportWriter "PASS" ,"Web Edit: " & sEdit(1) ,"Pass:  'Cleared the " & sEdit(1),0
'				Else
'					ReportWriter "PASS" ,"Web Edit: " & sEdit(1) ,"Pass:  '"&sValue&"' Data Set for Object : "&sEdit(1)&".",0
'				End If  
				ReportWriter "PASS" ,"Web Edit: " & sEdit(1) ,"Pass:  '"&sValue&"' Data Set for Object : "&sEdit(1)&"."  ,0		
			Else
				ReportWriter "FAIL" ,"Web Edit: " & sEdit(1) ,"Fail:  'Data Was not Set for Object : "&sEdit(1)&"'.",0
			End If          
	
	Case "weblist"
	
			sList = Split(sObjPropAndValue,":=")
			If BaseWindow.WebList(sObjPropAndValue).Exist Then
				'Browser(oSesObjDescriptions).Page(oObjectDescriptions).WebList(sObjPropAndValue).Select sValue
				BaseWindow.WebList(sObjPropAndValue).Select sValue
				sVal = BaseWindow.WebList(sObjPropAndValue).GetROProperty("Selection")
				ReportWriter "PASS" ,"WebList: " & sList(1) ,"Pass:  '"&sValue&"' is Selected from the list",0
			Else
			    ReportWriter "FAIL" ,"WebList: " & sList(1) ,"Fail:  '"&sValue&"' is not Selected from the list",0
			End If
			SetEditOrSelectListOrRadio = sVal

	Case "webfile"
	
			sList = Split(sObjPropAndValue,":=")
			If BaseWindow.WebFile(sObjPropAndValue).Exist Then
				'Browser(oSesObjDescriptions).Page(oObjectDescriptions).WebList(sObjPropAndValue).Select sValue
				BaseWindow.WebFile(sObjPropAndValue).Set sValue
				sVal = BaseWindow.WebList(sObjPropAndValue).GetROProperty("Selection")
				ReportWriter "PASS" ,"WebFile: " & sList(1) ,"Pass:  '"&sValue&"' Data Set for Object",0
			Else
			    ReportWriter "FAIL" ,"WebFile: " & sList(1) ,"Fail:  '"&sValue&"' Data Was not Set for Object",0
			End If
			SetEditOrSelectListOrRadio = sVal
			
'	Case "webradio"
	Case "webradiogroup"
	      
		If (IsNumeric(sValue)) Then
			sValue = "#"&sValue
		End If
		
			sWebRadioName = Split(sObjPropAndValue,":=")
		    If BaseWindow.WebRadioGroup(sObjPropAndValue).Exist Then
				BaseWindow.WebRadioGroup(sObjPropAndValue).Select sValue  
				ReportWriter "PASS","WebRadio: "&sWebRadioName(1),"Pass:  '"&sWebRadioName(1)&"' Radio button is selected: " ,0
			Else
				ReportWriter "FAIL","WebRadio: "&sWebRadioName(1),"Fail:  '"&sWebRadioName(1)&"' Radio button is not selected: ",0
			End If
			
	Case Else
			ReportWriter "FAIL", "Default Value of "&sObjType&": ","Fail:  'Default Value of "&sObjType&": not in case selection'",0
	End Select
	
	If Err <> 0 Then
		ReportWriter "FAIL","UnKnown Error","Fail:  '"&Err.Description&"'.",0
	End If 
	Err.Clear
	
End Function

'---------------------------------------------------------------------------------------------------------------
Sub SelectCheckBoxEMS(iCreationtime,sBoxName)
		' @HELP
		' @class:	 	clsWebControls
		' @method:	 	SelectCheckBox(sBoxName)
		' @returns:	 	
		' @notes:	 	Checks a specified Check Box
		' @END
		If Browser("micClass:=Browser","CreationTime:="&iCreationtime).Page("micClass:=Page").WebCheckBox(sBoxName).Exist Then
			Browser("micClass:=Browser","CreationTime:="&iCreationtime).Page("micClass:=Page").WebCheckBox(sBoxName).Set "ON"
			ReportWriter "PASS","CheckBox",sBoxName&" CheckBox Cheacked ",0
		Else
			 ReportWriter "FAIL","CheckBox",sBoxName& "CheckBox  Is not available ",0
			 ExitRun
		End If 
	End Sub
	
'---------------------------------------------------------------------------------------------------------------
'05th Sept 12
Sub UnSelectCheckBox(iCreationtime,sChkBoxPropNameAndVal)
		' @HELP
		' @class:	 	clsWebControls
		' @method:	 	UnSelectCheckBox(sBoxName)
		' @returns:	 	
		' @notes:	 	Checks a specified Check Box
		' @END
		sBoxName = Split(sChkBoxPropNameAndVal,":=")
		If Browser("micClass:=Browser","CreationTime:="&iCreationtime).Page("micClass:=Page").WebCheckBox(sChkBoxPropNameAndVal).Exist Then
			Browser("micClass:=Browser","CreationTime:="&iCreationtime).Page("micClass:=Page").WebCheckBox(sChkBoxPropNameAndVal).Set "OFF"
			ReportWriter "PASS","CheckBox",sBoxName(1)&" CheckBox UnChecked ",0
		Else
			 ReportWriter "FAIL","CheckBox",sBoxName(1)& "CheckBox  Is not available ",0
			 ExitRun
		End If 
	End Sub
'*********************************************************************************************************************************************************************
'''Added as part of portal scripts merge
'*********************************************************************************************************************************************************************
Function AutoWait(sMessage,iTime)

	' @HELP
	' @group    : WebControls
	' @method	: AutoWait
	' @returns	: None 
	' @parameter: sMessage : Control message for whom the Execution is waiting
	' @parameter: iTime : Specified amount of time to wait
	' @notes	: Login to the Account Management Portal
	' @example	:AutoWait("wait for page to be loaded",5)
	' @END
	
	Set oWshell = CreateObject("Wscript.Shell")
	oWshell.Popup sMessage& vbnewline & ": wait for" &iTime &"sec", iTime
	Set oWshell = Nothing
End Function

'*********************************************************************************************************************
Function BaseWebPage(iCreationTime)
	' @HELP
	' @group	: Webcontrols
	' @method	: BaseWebPage(iCreationTime)
	' @returns	: None
	' @parameter: iCreationTime	: Browser creation time (ex: 0, 1, 2 ...)
	' @notes	: 
	' @example	: BaseWebPage(0)
    ' @END
    
	If  iCreationTime=0 Then
		Set BaseWindow = Browser("micClass:=Browser").Page("micClass:=Page")
	Else
		Set BaseWindow = Browser("micClass:=Browser","CreationTime:=" &iCreationTime).Page("micClass:=Page")
	End If 
	Set BaseWebPage=BaseWindow
End Function 
'*********************************************************************************************************************
Function BrowserSynchronization()
	Dim oProc, oWMIServ, colProc
	Dim strPC, iNum, iPrcentUsage, ProcId
	strPC = "."
	Set oWMIServ = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strPC & "\root\cimv2")
	Set colProc = oWMIServ.ExecQuery("Select * from Win32_Process")
	For Each oProc In colProc
		if (instr(oProc.Name, "iexplore.exe") > 0 or instr(oProc.Name, "firefox.exe") > 0 or instr(oProc.Name, "java.exe") > 0 or instr(oProc.Name, "rundll32.exe") > 0) then
			iPrcentUsage = 1
			ProcId = oProc.ProcessId
			While iPrcentUsage > 0
				iPrcentUsage = eveCpuUSage(ProcId) 
				If iPrcentUsage=0 Then
					wait 0,500
					iPrcentUsage = eveCpuUSage(ProcId) 
				End If
			Wend
		end if 
	Next
	Set oWMIServ = nothing
	Set colProc = nothing
End Function

Function eveCpuUSage(vPid)
	On Error Resume Next 
   	Dim objWMI, objInstance1, perf_instance2, PercentProcessorTime, N1, N2, D1, D2
	Set objWMI = GetObject("winmgmts:\\" &  "." & "\root\cimv2")
	For Each objInstance1 in objWMI.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process Where IDProcess = '" & vPid & "'")
		N1 = objInstance1.PercentProcessorTime 
		D1 = objInstance1.TimeStamp_Sys100NS 
	Next
	wait 0,500
   For Each perf_instance2 in objWMI.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process Where IDProcess = '" & vPid & "'")       
	   N2 = perf_instance2.PercentProcessorTime       
	   D2 = perf_instance2.TimeStamp_Sys100NS
	Next
    PercentProcessorTime = ((N2 - N1)/(D2-D1))  * 100
	eveCpuUSage = Round(PercentProcessorTime ,0)
	Set objWMI = nothing
End Function

Function xBrowserSynchronization()

	' @HELP
	' @group	: Webcontrols
	' @funcion	: BrowserSynchronization()
	' @returns	: None
	' @parameter: None
	' @notes	: This is used for Browser Synchronization
	' @example	: BrowserSynchronization()
	' @END 
    
	On Error Resume Next 
	strComputer = "." 
	
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
	iNumBrowsers = 0 
	Set colProcessList = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'IEXPLORE.exe'") 
	
	For Each objProcess in colProcessList 
		iNumBrowsers = iNumBrowsers + 1 
	Next 
	
    CreationTime = iNumBrowsers - 1
	Browser("micClass:=Browser","CreationTime:="&CreationTime).Page("micClass:=Page").Sync
		
End Function
'*********************************************************************************************************************
Function CheckObjExist(sObjPropAndValue,iCreationTime)

		' @HELP
		' @group    : WebControls
		' @method	: CheckObjExist(sObjType,sObjPropAndValue,iCreationTime,iIndex)
		' @returns	: None
		' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
		' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
		' @notes	: Checks Whether an Object Exists in a Browser Window Or Not
		' @example	: CheckObjExist("link;name:=logout",0)
		' @END
	
		On Error Resume Next 
		
		'Get Base web page object
		BaseWebPage iCreationTime
			
		cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
		sObjType = LCase(Trim(cObjPropArray(0)))
		Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
		If sObjPropAndValue("name")="" Then
			objName=sObjPropAndValue(0)
		Else
			objName=sObjPropAndValue("name")
		End If
		
   		Select Case sObjType
   
			   Case "image"
			   		If BaseWindow.Image(sObjPropAndValue).Exist Then
						If BaseWindow.Image(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify Image: '"&objName&"'","Image: '" &objName&"'  Exist In Page",0
							CheckObjExist = True
						Else
							ReportWriter "FAIL","Verify Image: '"&objName&"'","Image: '" &objName&"'  Exists But Not visible In Page",0
							CheckObjExist = False
						End If
					Else
						ReportWriter "PASS","Verify Image: '"&objName&"'","Image: '" &objName&"'  Does Not Exist In Page",0
						CheckObjExist = False						
					End if					
				
			   Case "webarea"
			   
			   		If BaseWindow.WebArea(sObjPropAndValue).Exist Then
						If BaseWindow.WebArea(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify Text Area:  '"&objName&"'","Text Area: '" &objName&"' Exist In Page",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify Text Area:  '"&objName&"'","Text Area: '" &objName&"' Exists But Not visible In Page",0
							CheckObjExist = False
						End If
					Else
						ReportWriter "FAIL","Verify Text Area:  '"&objName&"'","Text Area: '" &objName&"' Does Not Exist In Page",0
						CheckObjExist = False
					End If
					   
			   Case "webedit"
			   
			   
			   		If BaseWindow.WebEdit(sObjPropAndValue).Exist Then
						
						If BaseWindow.WebEdit(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify Text Box: '"&objName&"'","Text Box: '" &objName&"' Exist In Page ",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify Text Box: '"&objName&"'","Text Box: '" &objName&"' Exists But Not visible In Page",0
							CheckObjExist = False
						End If						
					Else 
						ReportWriter "FAIL","Verify Text Box: '"&objName&"'","Text Box: '" &objName&"' Does Not Exist In Page ",0
						CheckObjExist = False
					End If	
							
				Case "webbutton"
				
					If BaseWindow.WebButton(sObjPropAndValue).Exist Then
						
						If BaseWindow.WebButton(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify Button:  '"&objName&"'","Button: '" &objName&"' Exist In Page ",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify Button:  '"&objName&"'","Button: '" &objName&"' Exists But Not visible In Page",0
							CheckObjExist = False
						End If							
					Else
						ReportWriter "FAIL","Verify Button:  '"&objName&"'","Button: '" &objName&"' Does Not Exist In Page ",0
						CheckObjExist = False
					End If
						
				Case "link"
				
					If BaseWindow.Link(sObjPropAndValue).Exist Then

						If BaseWindow.Link(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify Link:  '"&objName&"'","Link: '" &objName&"' Exist In Page ",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify Link:  '"&objName&"'","Link: '" &objName&"' Exists But Not visible In Page",0
							CheckObjExist = False
						End If
					Else	
						ReportWriter "FAIL","Verify Link:  '"&objName&"'","Link: '" &objName&"' Does Not Exist In Page ",0
						CheckObjExist = False
					End If	
				
				Case "weblist"
				
					If BaseWindow.WebList(sObjPropAndValue).Exist Then
						
						If BaseWindow.WebList(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify List Box:  '"&objName&"'","List Box: '" &objName&"' Exist In Page ",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify List Box:  '"&objName&"'","List Box: '" &objName&"' Exists But Not visible In Page",0
							CheckObjExist = False
						End If						
					Else
						ReportWriter "FAIL","Verify List Box:  '"&objName&"'","List Box: '" &objName&"' Does Not Exist In Page ",0
						CheckObjExist = False
					End If
		
				Case "webelement"
				
					If BaseWindow.WebElement(sObjPropAndValue).Exist Then

						If BaseWindow.WebElement(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify Text:  '" & objName & "'","Text: '" & objName&"' Exist In Page ",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify Text:  '" & objName & "'","Text: '" & objName&"'Exists But Not visible In Page",0
							CheckObjExist = False
						End If	           		    
					Else
						ReportWriter "FAIL","Verify Text:  '" & objName & "'","Text: '" & objName&"' Does Not Exist In Page ",0
						CheckObjExist = False
					End If
				Case "webradiogroup"
				
					If BaseWindow.webradiogroup(sObjPropAndValue).Exist Then
						If BaseWindow.webradiogroup(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify Radio Button: '"&objName&"'","Radio Button: '" &objName&"' Exist In Page ",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify Radio Button: '"&objName&"'","Radio Button: '" &objName&"' Exists But Not visible In Page",0
							CheckObjExist = False
						End If
	           		    
					Else
						ReportWriter "FAIL","Verify Radio Button: '"&objName&"'","Radio Button: '" &objName&"' Does Not Exist In Page ",0
						CheckObjExist = False
					End If
				Case "webcheckbox"
				
					If BaseWindow.webcheckbox(sObjPropAndValue).Exist Then
						If BaseWindow.webcheckbox(sObjPropAndValue).GetROProperty("height")<>0 Then
							ReportWriter "PASS","Verify CheckBox: '"&objName&"'","CheckBox: '" &objName&"' Exist In Page ",0
							CheckObjExist = True
						Else
							ReportWriter "PASS","Verify CheckBox: '"&objName&"'","CheckBox: '" &objName&"' Exists But Not visible In Page",0
							CheckObjExist = False
						End If	           		    
					Else
						ReportWriter "FAIL","Verify CheckBox: '"&objName&"'","CheckBox: '" &objName&"' Does Not Exist In Page ",0
						CheckObjExist = False
					End If		
									
			   Case Else
					ReportWriter "FAIL","UnKnown Object : '"&sObjType&"'","'"&sObjType&"'  Does Not Exist in page" ,1
					CheckObjExist = False
   		End Select		
   	  	
	   If Err <> 0 Then
			ReportWriter "FAIL","'"&sObjType&"' >> " & "UnKnown Error","'" & objName &"' >> " & "'"&Err.Description&"'.",0
	   End If 
	   Err.Clear
		   
End Function

'*********************************************************************************************************************
Function CheckObjIsDisabled (sObjPropAndValue,iCreationTime)
	' @HELP
	' @group	: Webcontrols
	' @funcion	: CheckObjIsDisabled()
	' @returns	: None
	' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
	' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
	' @notes    : To verify the object is disabled Or Not(Webedit,Weblist,Webradiogroup,...)
	' @example	: CheckObjIsDisabled ("link;name:=logout",0)
	' @END 
	

	'Get Base web page object
	BaseWebPage iCreationTime
		
	cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
	sObjType = LCase(Trim(cObjPropArray(0)))
	Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
	
	If sObjPropAndValue("name")="" Then
		objName=sObjPropAndValue(0)
	Else
		objName=sObjPropAndValue("name")
	End If
		
   Select Case sObjType

		Case "webbutton"
			iStatus = BaseWindow.WebButton(sObjPropAndValue).GetROProperty("disabled")
			
			'Modified by sudheer on August 29th, 2012
'			if CInt(iStatus) <> 1 Then
			if CInt(iStatus) = 1 Then
				ReportWriter "PASS","Verify Button:  '"&objName&"'","Button: '" &objName&"' Disabled In Page ",0
				CheckObjIsDisabled=True
			Else
				ReportWriter "FAIL","Verify Button:  '"&objName&"'","Button: '" &objName&"' Is Not Disabled In Page ",0
				CheckObjIsDisabled=False
			End If
			
	End Select		
		
End Function

'*********************************************************************************************************************
'*********************************************************************************************************************

Function CheckObjIsEnabled (sObjPropAndValue,iCreationTime)
	' @HELP
	' @group	: Webcontrols
	' @funcion	: CheckObjIsEnabled()
	' @returns	: None
	' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
	' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
	' @notes    : To verify the object is Enabled Or Not(Webedit,Weblist,Webradiogroup,...)
	' @example	: CheckObjIsDisabled ("link;name:=logout",0)
	' @END 
	
	Dim iStatus
	iStatus = 1
	On Error Resume Next
	
	'Get Base web page object
	BaseWebPage iCreationTime
		
	cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
	sObjType = LCase(Trim(cObjPropArray(0)))
	Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
	
	If sObjPropAndValue("name")="" Then
		objName=sObjPropAndValue(0)
	Else
		objName=sObjPropAndValue("name")
	End If
		
   Select Case sObjType

		Case "webbutton"
			iStatus = BaseWindow.WebButton(sObjPropAndValue).GetROProperty("disabled")
			if CInt(iStatus) =0 Then
				ReportWriter "PASS","Verify Button:  '"&objName&"'","Button: '" &objName&"' Enabled IN page",0
				CheckObjIsDisabled=True
			Else
				ReportWriter "FAIL","Verify Button:  '"&objName&"'","Button: '" &objName&"' Is Not Enabled In Page ",0
				CheckObjIsDisabled=False
			End If
			

	End Select		
		
End Function

'*********************************************************************************************************************
Function CheckObjNotExist(sObjPropAndValue,iCreationTime)

		' @HELP
		' @group    : WebControls
		' @method	: CheckObjNotExist(sObjType,sObjPropAndValue,iCreationTime)
		' @returns	: None
		' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
		' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
		' @notes	: Checks Whether an Object Exists in a Browser Window Or Not
		' @example	: CheckObjExist("webbutton;name:=logout",0)
		' @END
	
		On Error Resume Next 
		
		'Get Base web page object
		BaseWebPage iCreationTime
			
		cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
		sObjType = LCase(Trim(cObjPropArray(0)))
		Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
		
		If sObjPropAndValue("name")="" Then
			objName=sObjPropAndValue(0)
		Else
			objName=sObjPropAndValue("name")
		End If
		
	   Select Case sObjType
	   
			   Case "image"
			   
			   		If BaseWindow.Image(sObjPropAndValue).Exist(3) Then
			   			ReportWriter "FAIL","image: "&objName&" Verification In Browser","'"&"Image: "  &objName&"' Existing In Browser: " ,1
			   			CheckObjNotExist=False			   			
					Else 
						ReportWriter "PASS","image: "&objName&" Verification In Browser","'"&"Image: "  &objName&"' Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End if
		
			   Case "webarea"
			   
			   		If BaseWindow.WebArea(sObjPropAndValue).Exist(3) Then
						ReportWriter "FAIL","Text Area: "&objName&" Verification In Browser","'"&"Text Area: " &objName&"' Existing In Browser: " ,1
						CheckObjNotExist=False
					Else 
						ReportWriter "PASS","Text Area: "&objName&" Verification In Browser","'"&"Text Area: " &objName&"' Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If
					   
			   Case "webedit"
			   
			   		If BaseWindow.WebEdit(sObjPropAndValue).Exist(3) Then
						ReportWriter "FAIL","Text Box: "&objName&" Verification In Browser","'"&"Text Box: "  &objName&"' Existing In Browser: " ,1
						CheckObjNotExist=False
					Else 
						ReportWriter "PASS","Text Box: "&objName&" Verification In Browser","'"&"Text Box: " &objName&"'Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If			
				
				Case "webbutton"
				
					If BaseWindow.WebButton(sObjPropAndValue).Exist Then
						ReportWriter "FAIL","Button: "&objName&" Verification In Browser","'"&"Button: " &objName&"' Existing In Browser: " ,1
						CheckObjNotExist=False
					Else 
						ReportWriter "PASS","Button: "&objName&" Verification In Browser","'"&"Button: " &objName&"'Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If	
				
				Case "link"
				
					If BaseWindow.Link(sObjPropAndValue).Exist(3) Then
						ReportWriter "FAIL","Link: "&objName&" Verification In Browser","'"&"Link: " &objName&"' Existing In Browser: " ,1
						CheckObjNotExist=False
					Else 
						ReportWriter "PASS","Link: "&objName&" Verification In Browser","'"&"Link: " &objName&"'Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If	
				
				Case "weblist"
				
					If BaseWindow.WebList(sObjPropAndValue).Exist(3) Then
						ReportWriter "FAIL","List Box: "&objName&" Verification In Browser","'"&"List Box: " &objName&"' Existing In Browser: " ,1
						CheckObjNotExist=False
					Else 
						ReportWriter "PASS","List Box: "&objName&" Verification In Browser","'"&"List Box: " &objName&"'Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If							
		
				Case "webelement"
				
					If BaseWindow.WebElement(sObjPropAndValue).Exist(3) Then
						ReportWriter "FAIL","webelement: "&objName&" Verification In Browser","'"&"WebElement: " &objName&"' Existing In Browser: " ,1
						CheckObjNotExist=False
					Else
						ReportWriter "PASS","webelement: "&objName&" Verification In Browser","'"&"WebElement: " &objName&"'Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If
				Case "webradiogroup"
				
					If BaseWindow.webradiogroup(sObjPropAndValue).Exist Then
	           		    ReportWriter "FAIL","webradiogroup: "&objName&" Verification In Browser","'"&"webradiogroup: " &objName&"' Existing In Browser: ",0 
	           		    CheckObjNotExist=False
					Else
						ReportWriter "PASS","webradiogroup: "&objName&" Verification In Browser","'"&"webradiogroup: " &objName&"'Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If
				Case "webcheckbox"
				
					If BaseWindow.webcheckbox(sObjPropAndValue).Exist Then
	           		    ReportWriter "FAIL","webcheckbox: "&objName&" Verification In Browser","'"&"webcheckbox: " &objName&"' Existing In Browser: " ,1
	           		    CheckObjNotExist=False
					Else
						ReportWriter "PASS","webcheckbox: "&objName&" Verification In Browser","'"&"webcheckbox: " &objName&"'Not Existing In Browser: " ,0
						CheckObjNotExist=True
					End If			
					
					
			   Case Else
					ReportWriter "FAIL","UnKnown Object : "&sObjType,"'"&sObjType&"'  Not Found In Browser " ,1
					CheckObjNotExist=False
	   End Select
   
	   If Err <> 0 Then
			 ReportWriter "FAIL","'"&sObjType&"' >> " & "UnKnown Error","'" & objName &"' >> " & "'"&Err.Description&"'.",0
			 CheckObjNotExist=False
		End If 
		Err.Clear
		   
End Function

'*********************************************************************************************************************
Function CheckObjProperty(sObjPropAndValue,chkPropName,chkPropValue,wTime,iCreationTime)

	' @HELP
	' @group    : WebControls
	' @method	: CheckObjProperty(sObjPropAndValue,chkPropName,chkPropValue,wTime,iCreationTime)
	' @returns	: None
	' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
	' @parameter: chkPropName : The type of prperty to be verified(name,defaultvalue)
	 '@parameter: chkPropValue : the value of property
	' @parameter: wTime	: waiting time(optional)
	' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
	' @notes	: Checks Whether prperty has specified value or not
	' @example	: CheckObjProperty("weblist;name:=company","default value","vmware",0,0)
	' @END
	
	On Error Resume Next
	
	'Get Base web page object
	BaseWebPage iCreationTime
	
	cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
	sObjType = LCase(Trim(cObjPropArray(0)))
	Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
	
	If sObjPropAndValue("name")="" Then
		objName=sObjPropAndValue(0)
	Else
		objName=sObjPropAndValue("name")
	End If
	
   Select Case sObjType
   
		   Case "image"
		   
		   		If BaseWindow.Image(sObjPropAndValue).Exist Then
					If BaseWindow.Image(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","image: "&objName&" Verification In Browser","Pass:  '"&"image: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.Image(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","image: "&objName&" Verification In Browser","Fail:  '"&"image: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","image: "&objName&" Verification In Browser","Fail:  '" &"Image: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if
				
	
		   Case "webarea"
		   
		   		 If BaseWindow.webarea(sObjPropAndValue).Exist Then
		   			If BaseWindow.webarea(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","webarea: "&objName&" Verification In Browser","Pass:  '"&"image: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.webarea(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","webarea: "&objName&" Verification In Browser","Fail:  '"&"image: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","webarea: "&objName&" Verification In Browser","Fail:  '" &"Image: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if
				   
		   Case "webedit"
		   
		   		 If BaseWindow.webedit(sObjPropAndValue).Exist Then
		   		 	If BaseWindow.webedit(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","webedit: "&objName&" Verification In Browser","Pass:  '"&"image: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.webedit(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","webedit: "&objName&" Verification In Browser","Fail:  '"&"image: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","webedit: "&objName&" Verification In Browser","Fail:  '" &"Image: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if
						
			Case "webbutton"
			
				  If BaseWindow.webbutton(sObjPropAndValue).Exist Then
					If BaseWindow.webbutton(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","webbutton: "&objName&" Verification In Browser","Pass:  '"&"image: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.webbutton(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","webbutton: "&objName&" Verification In Browser","Fail:  '"&"image: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","webbutton: "&objName&" Verification In Browser","Fail:  '" &"Image: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if
					
			Case "link"
			
				 If BaseWindow.link(sObjPropAndValue).Exist Then
					If BaseWindow.link(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","link: "&objName&" Verification In Browser","Pass:  '"&"weblist: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.link(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","link: "&objName&" Verification In Browser","Fail:  '"&"weblist: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","link: "&objName&" Verification In Browser","Fail:  '" &"weblist: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if 
			
			Case "weblist"
			
				  If BaseWindow.weblist(sObjPropAndValue).Exist Then
					If BaseWindow.weblist(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","weblist: "&objName&" Verification In Browser","Pass:  '"&"weblist: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.weblist(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","weblist: "&objName&" Verification In Browser","Fail:  '"&"weblist: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","weblist: "&objName&" Verification In Browser","Fail:  '" &"weblist: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if						
	
			Case "webradiogroup"
			
				  If BaseWindow.webradiogroup(sObjPropAndValue).Exist Then
					If BaseWindow.webradiogroup(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","webradiogroup: "&objName&" Verification In Browser","Pass:  '"&"image: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.webradiogroup(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","webradiogroup: "&objName&" Verification In Browser","Fail:  '"&"image: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","webradiogroup: "&objName&" Verification In Browser","Fail:  '" &"Image: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if
				
			Case "webcheckbox"
			
				  If BaseWindow.Webcheckbox(sObjPropAndValue).Exist Then
				  	If BaseWindow.Webcheckbox(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","Webcheckbox: "&objName&" Verification In Browser","Pass:  '"&"Webcheckbox: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.Webcheckbox(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","Webcheckbox: "&objName&" Verification In Browser","Fail:  '"&"Webcheckbox: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","Webcheckbox: "&objName&" Verification In Browser","Fail:  '" &"Webcheckbox: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End If
			Case "webelement"
					If BaseWindow.webelement(sObjPropAndValue).Exist Then
				  		If BaseWindow.webelement(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
				  			ReportWriter "PASS","webelement: "&objName&" Verification In Browser","Pass:  '"&"webelement: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
				  			CheckObjProperty=True
						Else
							chkActValue=BaseWindow.webelement(sObjPropAndValue).GetROproperty(chkPropName)
							ReportWriter "FAIL","webelement: "&objName&" Verification In Browser","Fail:  '"&"webelement: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
							CheckObjProperty=False
						End If
					Else
						ReportWriter "FAIL","webelement: "&objName&" Verification In Browser","Fail:  '" &"webelement: " &objName&"'  Not Existing In Browser: ",0
						CheckObjProperty=False
					End if
				
			Case "webfile"

			  If BaseWindow.webfile(sObjPropAndValue).Exist Then
					If BaseWindow.webfile(sObjPropAndValue).CheckProperty (chkPropName,chkPropValue,wTime) Then
						ReportWriter "PASS","webfile: "&objName&" Verification In Browser","Pass:  '"&"webfile: " &objName&" Property "&chkPropName&" has the expected value: "&chkPropValue&".",0
						CheckObjProperty=True
					Else
						chkActValue=BaseWindow.webfile(sObjPropAndValue).GetROproperty(chkPropName)
						ReportWriter "FAIL","webfile: "&objName&" Verification In Browser","Fail:  '"&"webfile: " &objName&" Property "&chkPropName&" has the Actual Value "&chkActValue&"."&" expected value: "&chkPropValue&".",0
						CheckObjProperty=False
					End If
				Else
					ReportWriter "FAIL","webfile: "&objName&" Verification In Browser","Fail:  '" &"webfile: " &objName&"'  Not Existing In Browser: ",0
					CheckObjProperty=False
				End if
			
		   Case Else
		   
				ReportWriter "FAIL","UnKnown Object : "&sObjType,"Fail:  '"&sObjType&"'  Not Found In Browser ",0
				CheckObjProperty=False
   End Select
   
   If Err <> 0 Then
		ReportWriter "FAIL","'"&sObjType&"' >> " & "UnKnown Error","'" & objName &"' >> " & "'"&Err.Description&"'.",0
		CheckObjProperty=False
   End If 
   Err.Clear		   
End Function

'*********************************************************************************************************************
Function ClearAllCookies()
	' @HELP
	' @group	: Webcontrols
	' @funcion	: ClearAllCookies()
	' @returns	: None
	' @parameter: None
	' @notes	: This is used for clearing all cookies in the Browser
	' @example	: ClearAllCookies()
	' @END 
	
 	WebUtil.DeleteCookies
 	Wait 3
 
End Function

'*********************************************************************************************************************
'ClickOnTableLink ("html tag := TABLE;name:=test","OrderId")
Sub ClickOnTableLink(sObjPropAndValue, sVal2Search,iCreationTime)

	    ' @HELP
		' @group	: WebControls
		' @method	: ClickOnTableLink(sObjPropAndValue,sVal2Search)
		' @returns	: None
		' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
		' @parameter: sVal2Search : Name of the LINK 
		' @notes	: Clicks on a link located in Table
		' @example	: ClickOnTableLink("name:=result", "order",0)
		' @END
		
		On Error Resume Next 
		iFound = 0
						
		 'Get Base web page object
		BaseWebPage iCreationTime
		
		Set sObjPropAndValue=CreateObjectDescription(cObjPropArray)
		
		Set ActTbl=BaseWindow.ChildObjects(sObjPropAndValue)
		
		For i=0 to (ActTbl.Count)-1
			Rows=ActTbl(i).GetRoProperty("Rows")
			Cols=ActTbl(i).GetRoProperty("Cols")
			For Row=1 to Rows
				For Col=1 to Cols
					CellData=ActTbl(i).GetCellData(Row,Col)
                    If Len(Trim(CellData))>0 Then
                        If Trim(LCase(CellData))=Trim(LCase(sVal2Search)) Then
							Set Mylink=ActTbl(i).ChildItem(Row,Col,"Link",INDEX_VALUE_ZERO)
							If MLP_HIGH_LIGHT Then  Mylink.HighLight
							Mylink.Click
							iFound=1
							Exit For
						End If
					End If
				Next
			Next						
		Next 
			
		If iFound<>0 Then
			ReportWriter "PASS","Click Link:  '"&sVal2Search&"'","Link:  '"&sVal2Search &"' Clicked",0
		Else
			ReportWriter "FAIL","Click Link:  '"&sVal2Search&"'","Link:  '"&sVal2Search &"' Not Found In Page",0
		End If
		
	If Err <> 0 Then
		ReportWriter "FAIL",,"UnKnown Error"," '"&Err.Description&"'.",0
   	End If 
   	Err.Clear
   	
End Sub

'*********************************************************************************************************************
Function ClickOnWebObjects(sObjPropAndValue,iCreationTime)
		' @HELP
		' @group    : WebControls
		' @method   : ClickOnWebObjects(sObjPropAndValue,iCreationTime)
		' @returns  : None
		' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
		' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
		' @notes    : To get Default Value of any Object (Webedit,Weblist,Webradiogroup,...)
		' @example	: ClickOnWebObjects("link;name:=logout",0)
		' @END
	

		'Get Base web page object
		BaseWebPage iCreationTime
		cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
		sObjType = LCase(Trim(cObjPropArray(0)))
		Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
		
		If sObjPropAndValue("name")="" Then
			objName=sObjPropAndValue(0)
		Else
			objName=sObjPropAndValue("name")
		End If
		 
		
		Select Case sObjType
		
		Case "webbutton"
		
				If BaseWindow.WebButton(sObjPropAndValue).Exist Then
					If MLP_HIGH_LIGHT Then  BaseWindow.WebButton(sObjPropAndValue).HighLight
					BaseWindow.WebButton(sObjPropAndValue).Click
					ReportWriter "PASS","Click Button: '"&objName&"'","Button:  '"&objName&"'  Is Clicked",0
				Else
					ReportWriter "FAIL","Click Button: '"&objName&"'","Button:  '"&objName&"'  Is Not Clicked",0
				End If
				
		Case "link"
		
			    If BaseWindow.Link(sObjPropAndValue).Exist Then
					If MLP_HIGH_LIGHT Then  BaseWindow.Link(sObjPropAndValue).HighLight
					BaseWindow.Link(sObjPropAndValue).Click
					ReportWriter "PASS" ,"Click Link: '"&objName&"'","Link:  '"&objName&"'  Is Clicked",0
				Else
					ReportWriter "FAIL","Click Link: '"&objName&"'","Link:  '"&objName&"'  Is Not Clicked",0
				End If	
				
		Case "image"
		
				If BaseWindow.Image(sObjPropAndValue).Exist Then
					If MLP_HIGH_LIGHT Then  BaseWindow.Image(sObjPropAndValue).HighLight
					BaseWindow.Image(sObjPropAndValue).Click
					ReportWriter "PASS" ,"Click Image: '"&objName&"'","Image:  '"&objName&"'  Is Clicked",0
				Else
					ReportWriter "FAIL" ,"Click Image: '"&objName&"'","Image:  '"&objName&"'  Is Not Clicked",0
				End If
				
		Case "webcheckbox"
		
				If BaseWindow.WebCheckBox(sObjPropAndValue).GetROProperty("checked") = 0 Then
				   If MLP_HIGH_LIGHT Then  BaseWindow.WebCheckBox(sObjPropAndValue).HighLight
				   BaseWindow.WebCheckBox(sObjPropAndValue).Click
				   ReportWriter "PASS" ,"Select CheckBox:  '"& objName ,"CheckBox:  '"&objName&"' Is Selected.",0
				 Else
				   ReportWriter "FAIL","Select CheckBox:  '"& objName ,"CheckBox:  '"&objName&"' Is Not Selected.",0
				 End If
		Case "webradiogroup"
		
				If BaseWindow.webradiogroup(sObjPropAndValue).Exist Then
				   If MLP_HIGH_LIGHT Then  BaseWindow.webradiogroup(sObjPropAndValue).HighLight
				   BaseWindow.webradiogroup(sObjPropAndValue).Click
				   ReportWriter "PASS" ,"Click RadioButton:  '"& objName ,"RadioButton:  '"&objName&"' Is Clicked.",0
				 Else
				   ReportWriter "FAIL" ,"Click RadioButton:  '"& objName ,"RadioButton:  '"&objName&"' Is Not Clicked.",0
				 End If
				 
		Case "webelement"
		
			    If BaseWindow.WebElement(sObjPropAndValue).Exist Then
					If MLP_HIGH_LIGHT Then  BaseWindow.WebElement(sObjPropAndValue).HighLight
					BaseWindow.WebElement(sObjPropAndValue).Click
					ReportWriter "PASS" ,"Click WebElement: '"&objName&"'","WebElement:  '"&objName&"'  Is Clicked",0
				Else
					ReportWriter "FAIL" ,"Click WebElement: '"&objName&"'","WebElement:  '"&objName&"'  Is Not Clicked",0
				End If	
		Case "webfile"
              
	            If BaseWindow.WebFile(sObjPropAndValue).Exist Then
	            
			       If MLP_HIGH_LIGHT Then  BaseWindow.WebFile(sObjPropAndValue).HighLight
			       BaseWindow.WebFile(sObjPropAndValue).Click
			       ReportWriter "PASS","Click WebFile: '"&objName&"'","WebFile:  '"&objName&"'  Is Clicked",0
		        Else
		      
			       ReportWriter "FAIL" ,"Click WebFile: '"&objName&"'","WebFile:  '"&objName&"'  Is Not Clicked",0
		        End If
				   
		Case Else
			 ReportWriter "FAIL", "Operate Object:  '"&sObjType&"'","'"&sObjType&"'  Not Found in page",0
		End Select
	
	
End Function


'*********************************************************************************************************************
Function CreateObjectDescription(StrProperties)
' @HELP
    ' @group    : WebControls
    ' @method   : CreateObjectDescription
    ' @returns  : Description Object
    ' @parameter: StrProperties: List of properties seperated with ';'
    ' @notes    : To create and return descriptive object
    ' @example  : CreateObjectDescription("webedit;name:=login;index:=0")
    ' @END

    Dim objDescription
    Dim ObjArr
    Dim PropCount
    Dim ObjProperty
    Dim objHtmlId
    Dim LabelName  

    Set objDescription=Description.Create
	If instr(1,split(StrProperties,";")(0),":=")=0 Then
		'objDescription "micclass",split(StrProperties,";")(0)
		objDescription.Add "micclass",split(StrProperties,";")(0) 'modified by nagarjun on 11th sep,2012
		StrProperties=split(StrProperties,";",2)(1)

	End If

    ObjArr=split(StrProperties,";")
               
    For PropCount=0 to ubound(ObjArr)
    	ObjProperty=split(ObjArr(PropCount),":=")
	    objDescription(ObjProperty(0)).value=ObjProperty(1)
    Next
	
    If instr(1,StrProperties,"Label:=")<>0 Then
    	LabelName=objDescription("Label").value
        objDescription.Remove "Label"
        objHtmlId=BaseWindow.WebElement("html tag:=LABEL","outertext:="&LabelName).GetROProperty("attribute/htmlfor")
        objDescription.Add "html Id",objHtmlId
    End If

    If lCase(objDescription("RegularExpression").value)="false" Then
		Set objDescription=DisableRegularExpression(objDescription)
		objDescription.Remove "RegularExpression"
	End If

    Set CreateObjectDescription=objDescription


End Function
'*********************************************************************************************************************
Function DisableRegularExpression(dpObject)
	
	' @HELP
    ' @group    : WebControls
    ' @method   : DisableRegularExpression
    ' @returns  : Description Object with regular expression property as false
    ' @parameter: dpObject: Description Object 
    ' @notes    : To make regular expression as False in descriptive object
    ' @example  : CreateObjectDescription(objDescription)
    ' @END

Dim dPropertyIndex

	For dPropertyIndex=0 to dpObject.Count-1
		dpObject(dPropertyIndex).RegularExpression=False
	Next

Set DisableRegularExpression=dpObject

End Function

'*********************************************************************************************************************
Function GetColumn(ByVal TObject)
	' @HELP
	' @group    : 
	' @method   : GetColumn(ByVal TObject)
	' @returns  : Column No
	' @parameter: TObject : Webobject (Link,Button,Webelement)
	' @notes    : To Get The Column no
	' @example	: GetColumn(Basewebpage(0).link("name"))
	' @END
	
    GetColumn = -1
    Set TObject = TObject.Object
 
    Do
        If TObject.nodeName = "TD" Or TObject.nodeName = "TH" Then
            GetColumn = TObject.cellIndex + 1
            Exit Function
        End If
 
        Set TObject = TObject.parentNode
    Loop Until TObject.nodeName = "HTML"
End Function
'*********************************************************************************************************************
Function GetDefaultValueOfWebObject(sObjPropAndValue,iCreationTime)

		' @HELP
		' @group	: Webcontrols
		' @method	: GetDefaultValueOfWebObject(sObjType,sObjPropAndValue,iCreationTime)
		' @returns	: Default value of webobject
		' @parameter: sObjType : Type of Object(webedit,weblist,webradiogroup)
		' @parameter: sObjPropAndValue : Name of Object
		' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
		' @notes	: To get Default Value of any Object (Webedit,Weblist,Webradiogroup,...)
		' @example	: GetDefaultValueOfWebObject("weblist;name:=department",0)
        ' @END

		On Error Resume Next 

		'Get Base web page object
		BaseWebPage iCreationTime
		
		cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
		sObjType = LCase(Trim(cObjPropArray(0)))
		Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
		
		If sObjPropAndValue("name")="" Then
			objName=sObjPropAndValue(0)
		Else
			objName=sObjPropAndValue("name")
		End If
		
		Select Case sObjType
		
			Case "webedit"
			
				If BaseWindow.WebEdit(sObjPropAndValue).Exist Then
					sDefaultValue = BaseWindow.WebEdit(sObjPropAndValue).GetRoProperty("value")
					GetDefaultValueOfWebObject = sDefaultValue				
					ReportWriter "PASS","Verify Text Box: '"&objName&"'","Text Box: '" &objName&"' Exist In Page and Default value is: " &sDefaultValue,0
				Else 
					ReportWriter "FAIL","Verify Text Box: '"&objName&"'","Text Box: '" &objName&"' Does Not Exist In Page ",0
				End If				
				
			Case "weblist"
						
				If BaseWindow.WebList(sObjPropAndValue).Exist Then
					sDefaultValue = BaseWindow.WebList(sObjPropAndValue).GetRoProperty("value")
					'MsgBox "Added for the purpose of debugging to output the selection of the list box "&sDefaultValue
					GetDefaultValueOfWebObject = sDefaultValue
					ReportWriter "PASS","Verify List Box:  '"&objName&"'","List Box: '" &objName&"' Exist In Page And Default Value is: " &sDefaultValue,0
				Else
					ReportWriter "FAIL","Verify List Box:  '"&objName&"'","List Box: '" &objName&"' Does Not Exist In Page ",0
				End If	
			
			Case "webelement"
						
				If BaseWindow.WebElement(sObjPropAndValue).Exist Then
					sDefaultValue = BaseWindow.WebElement(sObjPropAndValue).GetRoProperty("innertext")
					GetDefaultValueOfWebObject = sDefaultValue
					ReportWriter "PASS","Verify Element:  '"&objName&"'","Element: '" &objName&"' Exist In Page And Default Value is: " &sDefaultValue,0
				Else
					ReportWriter "FAIL","Verify Element:  '"&objName&"'","Element: '" &objName&"' Does Not Exist In Page ",0
				End If	
			    			    
			Case "webcheckbox"
			
			    If BaseWindow.WebCheckBox(sObjPropAndValue).Exist Then				   
				   sDefaultValue = BaseWindow.WebCheckBox(sObjPropAndValue).GetRoProperty("checked")
				   If CInt(sDefaultValue)  = CInt(1) Then 
				   	GetDefaultValueOfWebObject = sDefaultValue
				   	ReportWriter "PASS" ,"Verify CheckBox:  '"&objName&"'","CheckBox: '" &objName&"' Exist In Page And Default Value is: " &sDefaultValue,0
		       	   Else
					ReportWriter "FAIL","Verify CheckBox:  '"&objName&"'","CheckBox: '" &objName&"' Does Not Exist In Page ",0
				   End If		
				   Else
				   	 ReportWriter "FAIL","Verify CheckBox:  '"&objName&"'","CheckBox: '" &objName&"' Does Not Exist In Page ",0
				   End If	 
    						
			Case Else
				ReportWriter "FAIL", "Default Value of "&sObjType&": ","Fail:  'Default Value of "&sObjType&": not in case selection'",0
	   End Select
	   
		If Err <> 0 Then
			ReportWriter "FAIL","'"&sObjType&"' >> " & "UnKnown Error","'" & objName &"' >> " & "'"&Err.Description&"'.",0
		End If 
		Err.Clear
End Function

'*********************************************************************************************************************
Function GetObjFromParent(sChildObjProperties,sParentObjProperties,iCreationTime)

	' @HELP
	' @group    : 
	' @method   : GetObjFromParent
	' @returns  : Identified Child Object
	' @parameter: ChildObject Properties,ParentObject Properties,CreationTime
	' @notes    : To get a child object from a parent object
	' @example	: GetObjFromParent(CHILDPROPS,PARENT PROPS,0)
	' @END
	
	Dim sParentObject
	Dim sChildObject
	Dim sChildObjectDesc
	Dim sParentObjectDesc
	Dim sParentObjectClass
	Dim sChildObjectClass
	Dim sChildObjList
	
	Set sChildObjectDesc=CreateObjectDescription(sChildObjProperties)
	Set sParentObjectDesc=CreateObjectDescription(sParentObjProperties)
	
	sParentObjectClass=sParentObjectDesc("micclass")
	sChildObjectClass=sChildObjectDesc("micclass")

	BaseWebPage iCreationTime

	Execute("set sParentObject=BaseWindow."&sParentObjectClass&"(sParentObjectDesc)")

	Set sChildObjList=sParentObject.ChildObjects(sChildObjectClass)

	If  sChildObjList.count=1 Then
			Set GetObjFromParent=sChildObjList(0)

	ElseIf  sChildObjList.count=0 Then
		ReportWriter "FAIL","GetObjFromParent","Object Not Exist in The Parent",0

	Else
		ReportWriter "FAIL","GetObjFromParent","Multiple Objects found with the specified description",0
	End If

End Function
'*********************************************************************************************************************
Function GetRow(ByVal TObject)
	' @HELP
	' @group    : 
	' @method   : GetRow(ByVal TObject)
	' @returns  : Row No
	' @parameter: TObject : Webobject (Link,Button,Webelement)
	' @notes    : To Get The Row no
	' @example	: GetRow(Basewebpage(0).link("name"))
	' @END
	
    GetRow = -1
    Set TObject = TObject.Object
 	Do
    	If TObject.nodeName = "TR" Then
         	GetRow = TObject.rowIndex + 1
        	Exit Function
        End If
 
        Set TObject = TObject.parentNode
    Loop Until TObject.nodeName = "HTML"
End Function
'*********************************************************************************************************************
Function SetValuesForWebobjects(sObjPropAndValue,sValue,iCreationTime)

	' @HELP
	' @group    : WebControls
	' @method   : SetValuesForWebobjects(sObjPropAndValue,sValue,iCreationTime)
	' @returns  : None
	' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
	' @parameter: sValue : Value to be Set for respective object
	' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
	' @notes    : To Set Or Select (Webedit,Weblist,Webradiogroup,...)
	' @example	: SetValuesForWebobjects("webedit;name:=username","vmware",0)
	' @END

                                     
	sObjType = LCase(Trim(sObjType))
	
	'Get Base web page object
	BaseWebPage iCreationTime
	
	cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
	sObjType = LCase(Trim(cObjPropArray(0)))
	Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
	
	If sObjPropAndValue("name")="" Then
		objName=sObjPropAndValue(0)
	Else
		objName=sObjPropAndValue("name")
	End If
	
	Select Case sObjType
	
		Case "webedit"
		
			If BaseWindow.WebEdit(sObjPropAndValue).Exist Then
	
				If MLP_HIGH_LIGHT Then  BaseWindow.WebEdit(sObjPropAndValue).HighLight
			
				If sValue = "" Then
					BaseWindow.WebEdit(sObjPropAndValue).Set ""
					ReportWriter "PASS" ,"Clear Text Box:  '" & objName ,"' Text Box:  '"& objName&"'  Cleared ",0
				Else
					BaseWindow.WebEdit(sObjPropAndValue).Set sValue
		
					ReportWriter "PASS" ,"Enter Data In Text Box:  '" & objName ,"Data  '"&sValue&"'  Entered In Text Box:  '"& objName&"' ",0
				End If    		
			Else
				ReportWriter "FAIL","Enter Data In Text Box: '"&objName&"'","Text Box: '" &objName&"' Does Not Exist In Page ",0
			End If          
		
		Case "weblist"
		
			If BaseWindow.WebList(sObjPropAndValue).Exist Then
				If MLP_HIGH_LIGHT Then  BaseWindow.WebList(sObjPropAndValue).HighLight
				BaseWindow.WebList(sObjPropAndValue).Select sValue
				sVal = BaseWindow.WebList(sObjPropAndValue).GetROProperty("Selection")
				ReportWriter "PASS" ,"Select From List Box: '" & objName ,"Option  '"&sVal&"' Is Selected From List Box:  '"& objName &"' ",0
			Else
			    ReportWriter "FAIL" ,"Select From List Box: '" & sobjName ,"Option  '"&sVal&"' Is Not Selected From List Box:  '"& objName &"' ",0
			End If
			SetValuesForWebobjects = sVal
				
		Case "webradiogroup"
		     
			If (IsNumeric(sValue)) Then
'				sValue = "#"&sValue
				sValue = "#"&sValue
			End If
			If BaseWindow.WebRadioGroup(sObjPropAndValue).Exist Then
		    	If MLP_HIGH_LIGHT Then  BaseWindow.WebRadioGroup(sObjPropAndValue).HighLight
				BaseWindow.WebRadioGroup(sObjPropAndValue).Select sValue  
'				BaseWindow.WebRadioGroup(sObjPropAndValue).Select sValue1  
				ReportWriter "PASS","Select Radio Button: "&objName,"'  Radio Button:  '"&objName&"'  Is Selected " ,0
			Else
				ReportWriter "FAIL","Select Radio Button: "&objName,"'  Radio Button:  '"&objName&"'  Is Not Selected ",0
			End If
		Case "webcheckbox"
	        If BaseWindow.webcheckbox(sObjPropAndValue).Exist Then
	        	If MLP_HIGH_LIGHT Then  BaseWindow.webcheckbox(sObjPropAndValue).HighLight
	            BaseWindow.webcheckbox(sObjPropAndValue).Set sValue
	            ReportWriter "PASS" ,"WebCheckBox: " & objName ," '"&objName &"' CheckBox is Selected ",0
	        Else
	            ReportWriter "FAIL" ,"WebCheckBox: " & objName ,"'"&objName&"' CheckBox is not Selected",0
	        End If
				
		Case Else
				ReportWriter "FAIL", "operate:  '"&sObjType&"'","'"&sObjType&"'  Is Not Identified",0
	End Select
	

End Function
'*********************************************************************************************************************
 Function VerifyColumnValueInWebTable(sDataColum,sTabName, iCreationTime)
	' @HELP
	' @group    : 
	' @method   : VerifyColumnValueInWebTable()
	' @returns  : None
	' @parameter: sDataColum : Column value of the Table of which the Data is to be read
	' @parameter: sTabName : Table name of which the Data is to be read
	' @parameter: iBrowserIndex	: Creation Time of the Browser
	' @notes    : To Verify whether a cell corresponding to a specified cell contains valid data or not
	' @example	: VerifyColumnValueInWebTable("order id","result", 0)
	' @END
	
	Set oDesc = Description.Create() 
	oDesc("micclass").Value = sTabName 
	Dim sTag		
	'Get Base web page object
	BaseWebPage iCreationTime
		
	Set Lists = BaseWindow.ChildObjects(oDesc) 
	sTag = 1
	
		NumberOfLists = Lists.Count()
		For i = 0 to NumberOfLists - 1
			rows = Lists(i).rowcount()
			For r = 1 to rows
				cols = Lists(i).ColumnCount(r)
				For C = 1 to cols
					sText = Lists(i).GetCellData(r,c)
					If (UCase(trim(sText))=UCase(trim(sDataColum))) Then
						sTag = 0
                        sWebElementVal = Lists(i).GetCellData(r+1,c)
                        If IsNull(sWebElementVal) Or sWebElementVal = "" Or sWebElementVal = " " Then
'                        	Reporter.ReporterEvent "Fail", ""&sText&" Verification", ""&sText&" Verification Failed."&sText& " = " &sWebElementVal,INDEX_VALUE_ZERO
							'Modified by sudheer
                        	ReportWriter "FAIL", ""&sText&" Verification", "No data exist in "&sText&" column",0
							Exit Function
						Else
							'ReportWriter "PASS", ""&sText&" Verification", ""&sText&" Verification Failed. Expected : "&sText&" Actual Value := " &sWebElementVal,INDEX_VALUE_ZERO,0
							'Modified by sudheer
							ReportWriter "PASS", ""&sText&" Verification", ""&sText&" column Verification passed. "&sText&" column contains :="&sWebElementVal,0
							VerifyColumnValueInWebTable = sWebElementVal
							Exit Function
						End If
						
                    End If
				Next
			Next
		Next
		
		If 	sTag = 1 Then
'			ReportWriter "FAIL", ""&sText&" Verification", sText&" filed does not exist",INDEX_VALUE_ZERO,1
			'Modified by sudheer
			ReportWriter "FAIL", ""&sDataColum&" Verification", sDataColum&" field does not exist",0
		End If 
		
End Function

'*********************************************************************************************************************
Function VerifyDefaultValueOfWebObject(sObjPropAndValue,iCreationTime,sData2Verify)

		' @HELP
		' @group	: Webcontrols
		' @method	: VerifyDefaultValueOfWebObject(sObjType,sObjPropAndValue,iCreationTime,sData2Verify)
		' @returns	: Default value of webobject
		' @parameter: sObjPropAndValue : Defined Constant for the Object
		' @parameter: iCreationTime	: Browser creation time (ex: 1,2 ...)
		' @notes	: To Verify Default Value of any Object (Webedit,Weblist,Webradiogroup,...)
		' @example	: VerifyDefaultValueOfWebObject("weblist;name:=company",0,"vmware")
		' @example	: VerifyDefaultValueOfWebObject(ObjConstName,0,"vmware")
        ' @END
		
		On Error Resume Next 
		
		'Get Base web page object
		BaseWebPage iCreationTime
		
		cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
		sObjType = LCase(Trim(cObjPropArray(0)))
		Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
		
		If sObjPropAndValue("name")="" Then
			objName=sObjPropAndValue(0)
		Else
			objName=sObjPropAndValue("name")
		End If
		
		Select Case sObjType
		
			Case "webedit"
			
				If BaseWindow.WebEdit(sObjPropAndValue).Exist Then
					sDefaultValue = BaseWindow.WebEdit(sObjPropAndValue).GetRoProperty("value")
					If LCase(Trim(sDefaultValue)) = LCase(Trim(sData2Verify)) Then
						ReportWriter "PASS","webedit:"&objName&" 's Default value is:", "Expected: "&sData2Verify &VbNewLine &"Actual: "&sDefaultValue,0
						VerifyDefaultValueOfWebObject = True
					End If
				Else 
					ReportWriter "FAIL","webedit:"&objName&" 's Default value is:", "Expected: "&sData2Verify &VbNewLine &"Actual: "&sDefaultValue,0
				End If				
				
			Case "weblist"
									
				If BaseWindow.WebList(sObjPropAndValue).Exist Then
					sDefaultValue = BaseWindow.WebList(sObjPropAndValue).GetRoProperty("value")
					If LCase(Trim(sDefaultValue)) = LCase(Trim(sData2Verify)) Then
						VerifyDefaultValueOfWebObject = sDefaultValue
						ReportWriter "PASS","weblist: "&objName&" 's Default value is: ","Expected: "&sData2Verify &VbNewLine &"Actual: "&sDefaultValue,0
					End If
				Else
					ReportWriter "FAIL","weblist: "&objName&" 's Default value is: ","Expected: "&sData2Verify &VbNewLine &"Actual: "&sDefaultValue,0
				End If	
				
			    			    
			Case "webcheckbox"
			
				If BaseWindow.WebCheckBox(sObjPropAndValue).GetROProperty("checked") = 0 Then				   
				   sDefaultValue = BaseWindow.WebCheckBox(sObjPropAndValue).GetRoProperty("checked")
				   If LCase(Trim(sDefaultValue)) = LCase(Trim(sData2Verify)) Then
				   		VerifyDefaultValueOfWebObject = sDefaultValue
				   		ReportWriter "PASS" ,"CheckBox: "& objName&" 's Default value is: " ,"Expected: "&sData2Verify &VbNewLine &"Actual: "&sDefaultValue,0
				   End If 
				 Else
				   ReportWriter "FAIL" ,"CheckBox: "& objName&" 's Default value is: " ,"Expected: "&sData2Verify &VbNewLine &"Actual: "&sDefaultValue,0
				 End If			 
    						
			Case Else
				ReportWriter "FAIL", "Default Value of "&sObjType&": ","Fail:  'Default Value of "&sObjType&": not in case selection'",0			
	   End Select	   
	   
End Function
'*********************************************************************************************************************
Function VerifyObjExistInParent(sChildObjProperties,sParentObjProperties,iCreationTime)

	' @HELP
	' @group    : 
	' @method   : VerifyObjExistInParent
	' @returns  : true or False
	' @parameter: ChildObject Properties,ParentObject Properties,CreationTime
	' @notes    : To verify a child object exist in a parent object
	' @example	: VerifyObjExistInParent(CHILDPROPS,PARENT PROPS,0)
	' @END
	
	Dim sParentObject
	Dim sChildObject
	Dim sChildObjectDesc
	Dim sParentObjectDesc
	Dim sParentObjectClass
	Dim sChildObjectClass
	Dim sChildObjList
	
	Set sChildObjectDesc=CreateObjectDescription(sChildObjProperties)
	Set sParentObjectDesc=CreateObjectDescription(sParentObjProperties)
	
	sParentObjectClass=sParentObjectDesc("micclass")
	sChildObjectClass=sChildObjectDesc("micclass")

	BaseWebPage iCreationTime

	Execute("set sParentObject=BaseWindow."&sParentObjectClass&"(sParentObjectDesc)")

	Set sChildObjList=sParentObject.ChildObjects(sChildObjectClass)

	If  sChildObjList.count=1 Then
		ReportWriter "PASS","VerifyObjExistInParent","Object Exitst",0
	ElseIf  sChildObjList.count=0 Then
		ReportWriter "FAIL","VerifyObjExistInParent","Object Not Exist in The Parent",0
	Else
		ReportWriter "FAIL","VerifyObjExistInParent","Multiple Objects found with the specified description",0
	End If

End Function

'*********************************************************************************************************************
Function VerifyValueByColumnNameinWebTable(sColumnName,sColumnValue,iCreationTime)
	' @HELP
	' @group    : 
	' @method   : VerifyValueByColumnNameinWebTable()
	' @returns  : None
	' @parameter: sFieldName : Column Name of the Table 
	' @parameter: sFieldValue : Column value of the Table 
	' @parameter: iCreationTime : Creation Time of the Browser
	' @notes    : To Verify whether a Column value corresponding to a specified  Column name
	' @example	: VerifyValueByColumnNameinWebTable("Contract ID","3234516",0)
	' @END
		
		
		'Get Base web page object
		BaseWebPage iCreationTime
		
		Set oDesc = Description.Create() 
		oDesc("micclass").Value = "WebTable" 
		
		Set Lists = BaseWindow.ChildObjects(oDesc) 
		
		NumberOfLists = Lists.Count()
		sTag = 1
		For i = 0 to NumberOfLists - 1
			rows = Lists(i).rowcount()
			For r = 1 to rows
				cols = Lists(i).ColumnCount(r)
				For C = 1 to cols
					sTextToSearch = Lists(i).GetCellData(r,c)
					If (trim(Lcase(sTextToSearch))= Trim(Lcase(sColumnName))) Then
						sTag = 0
                        sWebElementVal = Lists(i).GetCellData(r+1,c)
                        If InStr(Trim(Lcase(sWebElementVal)),Trim(Lcase(sColumnValue))) > 0 Or InStr(Trim(Lcase(sColumnValue)),Trim(Lcase(sWebElementVal))) > 0 Then 
'						 Modiify by sudheer
'                        If LCase(Replace(sWebElementVal," ",""))= LCase(Replace(sColumnValue," ","")) Then 
                    		ReportWriter "PASS", ""&sColumnName&"Verification", ""&sColumnName& " Verification Passed. Expected : "&sColumnValue&" Actual Value : = " &sWebElementVal,0
							VerifyValueByColumnNameinWebTable = sWebElementVal
							Exit Function
						Else 
							ReportWriter "FAIL", ""&sColumnName&" Verification", ""&sColumnName&" Verification Failed. Expected : "&sColumnValue&" Actual Value := " &sWebElementVal,0
							Exit Function
						End If
                    End If
				Next
			Next
		Next
		If sTag = 1 Then
			ReportWriter "FAIL", ""&sColumnName&" Verification", sColumnName&" field does not exist",0
		End if
End Function
'*********************************************************************************************************************
Function VerifyValueByColumnNameinWebTableWithTblName(sTbl,sColumnName,sColumnValue,iCreationTime)
	' @HELP
	' @group    : 
	' @method   : VerifyValueByColumnNameinWebTable()
	' @returns  : None
	' @parameter: sTblName : name of the Table 
	' @parameter: sFieldName : Column Name of the Table 
	' @parameter: sFieldValue : Column value of the Table 
	' @parameter: iCreationTime : Creation Time of the Browser
	' @notes    : To Verify whether a Column value corresponding to a specified  Column name
	' @example	: VerifyValueByColumnNameinWebTable("Contract ID","3234516",0)
	' @END
		
		On Error Resume Next 
		
		'Get Base web page object
		BaseWebPage iCreationTime
			sTag=1
			Set sTblName=CreateObjectDescription(sTbl)
			rows = BaseWindow.Webtable(sTblName).rowcount()
			For r = 1 to rows
				cols = BaseWindow.Webtable(sTblName).ColumnCount(r)
				For C = 1 to cols
					sTextToSearch = BaseWindow.Webtable(sTblName).GetCellData(r,c)
					'msgbox sTextToSearch
					If (trim(Lcase(sTextToSearch))= trim(Lcase(sColumnName))) Then
						'msgbox "True"
						sTag = 0
                        sWebElementVal = BaseWindow.Webtable(sTblName).GetCellData(r+1,c)
                        If InStr(Trim(Lcase(sWebElementVal)),Trim(Lcase(sColumnValue))) > 0 Or InStr(Trim(Lcase(sColumnValue)),Trim(Lcase(sWebElementVal))) > 0 Then 
                    		ReportWriter "PASS", ""&sColumnName&"Verification", ""&sColumnName& " Verification Passed. Expected : "&sColumnValue&" Actual Value : " &sWebElementVal,0
							VerifyValueByColumnNameinWebTableWithTblName = sWebElementVal
							Exit Function
						Else 
							ReportWriter "FAIL", ""&sColumnName&" Verification", ""&sColumnName&" Verification Failed. Expected : "&sColumnValue&" Actual Value : " &sWebElementVal,0
							Exit Function
						End If
                    End If
				Next
			Next

		If sTag = 1 Then
			ReportWriter "FAIL", ""&sColumnName&" Verification", sColumnName&" field does not exist",0
		End if
End Function
'*********************************************************************************************************************
Function VerifyValueByFieldNameinWebTableWithTblName(sTbl,sColumnName,sColumnValue,iCreationTime)
	' @HELP
	' @group    : 
	' @method   : VerifyValueByColumnNameinWebTable()
	' @returns  : None
	' @parameter: sTblName : name of the Table 
	' @parameter: sFieldName : Column Name of the Table 
	' @parameter: sFieldValue : Column value of the Table 
	' @parameter: iCreationTime : Creation Time of the Browser
	' @notes    : To Verify whether a Column value corresponding to a specified  Column name
	' @example	: VerifyValueByColumnNameinWebTable("Contract ID","3234516",0)
	' @END
		
		On Error Resume Next 
		
		'Get Base web page object
		BaseWebPage iCreationTime

			Set sTblName=CreateObjectDescription(sTbl)
			rows = BaseWindow.Webtable(sTblName).rowcount()
			For r = 1 to rows
				cols = BaseWindow.Webtable(sTblName).ColumnCount(r)
				For C = 1 to cols
					sTextToSearch = BaseWindow.Webtable(sTblName).GetCellData(r,c)
					If (trim(Lcase(sTextToSearch))= Trim(Lcase(sColumnName))) Then
						sTag = 0
                        sWebElementVal = BaseWindow.Webtable(sTblName).GetCellData(r,c+1)
                        If InStr(Trim(Lcase(sWebElementVal)),Trim(Lcase(sColumnValue))) > 0 Or InStr(Trim(Lcase(sColumnValue)),Trim(Lcase(sWebElementVal))) > 0 Then 
                    		ReportWriter "PASS", ""&sColumnName&"Verification", ""&sColumnName& " Verification Passed. Expected : "&sColumnValue&" Actual Value : = " &sWebElementVal,0
							VerifyValueByColumnNameinWebTable = sWebElementVal
							Exit Function
						Else 
							ReportWriter "FAIL", ""&sColumnName&" Verification", ""&sColumnName&" Verification Failed. Expected : "&sColumnValue&" Actual Value := " &sWebElementVal,0
							Exit Function
						End If
                    End If
				Next
			Next

		If sTag = 1 Then
			ReportWriter "FAIL", ""&sColumnName&" Verification", sColumnName&" filed does not exist",0
		End if
End Function
'*********************************************************************************************************************
Function VerifyValueInWebTable(sDataColum,sObjPropAndValue, iCreationTime)
	' @HELP
	' @group    : 
	' @method   : VerifyValueInWebTable()
	' @returns  : None
	' @parameter: sDataColum : Column value of the Table of which the Data is to be read
	' @parameter: sObjPropAndValue	: Table name of which the Data is to be read
	' @parameter: iBrowserIndex	: Creation Time of the Browser
	' @notes    : To Verify whether a cell corresponding to a specified cell contains valid data or not
	' @example	: VerifyValueInWebTable("Contract","webtable;name:=history", 0)
	' @END
		
		On Error Resume Next 
		
		'Get Base web page object
		BaseWebPage iCreationTime
		Dim sTag
		cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
		sObjType = LCase(Trim(cObjPropArray(0)))
		Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
		
		If sObjPropAndValue("name")="" Then
			objName=sObjPropAndValue(0)
		Else
			objName=sObjPropAndValue("name")
		End If
	
'        Set Lists = BaseWindow.ChildObjects(oDesc) 
        Set Lists = BaseWindow.ChildObjects(sObjPropAndValue) 
        
		
		NumberOfLists = Lists.Count()
		sTag = 1
		For i = 0 to NumberOfLists - 1
			
			rows = Lists(i).rowcount()
			For r = 1 to rows
				cols = Lists(i).ColumnCount(r)
				For C = 1 to cols
					sText = Lists(i).GetCellData(r,c)
					If (trim(sText)=trim(sDataColum)) Then
						sTag = 0
						sWebElementVal = Lists(i).GetCellData(r,c+1)
'                        If IsNull(sWebElementVal) Or sWebElementVal = "" Or sWebElementVal = " " Then
'							Modified by sudheer
                        If IsNull(sWebElementVal) Or sWebElementVal = "" Or sWebElementVal = " " Or sWebElementVal="ERROR: The specified cell does not exist."Then
'                        	ReportWriter "FAIL", ""&sText&" Verification", ""&sText&" Verification Failed."&sText& " = " &sWebElementVal,1
'							Modified by sudheer
							ReportWriter "FAIL", ""&sText&" Verification", "No data exist in "&sText&" column",0
                        	Exit Function
						Else
							ReportWriter "PASS", ""&sText&"Verification", ""&sText& " Verification Passed." &sText& " = " &sWebElementVal,0
							VerifyValueInWebTable = sWebElementVal
							Exit Function
						End If
                    End If
				Next
			Next
		Next
		If 	sTag = 1 Then
'			ReportWriter "FAIL", ""&sFieldName&" Verification", sFieldName&" field does not exist",0
			'Modifed by sudheer
			ReportWriter "FAIL", ""&sDataColum&" Verification", sDataColum&" field does not exist",0
		End If
End Function

'*********************************************************************************************************************
'Usage: WaitForWebObjects "name:=Feedback", INDEX_VALUE_ZERO, 1
Function WaitForWebObjects(sObjPropAndValue, iCreationTime, iNeedsToWait)
   	
   	' @HELP
	' @group   	: WebControls
	' @method	: WaitForWebObjects(sObjPropAndValue, iCreationTime, iNeedsToWait)
	' @returns	: None
	' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
	' @parameter: iCreationTime : Browser creation time (ex: 1,2 ...)
	' @parameter: iNeedsToWait	: Number of times to wait for identify an Object
	' @notes	: Wait for an Object until identify the object otherwise exit after the specified time
	' @example	:WaitForWebObjects("name:=login", 0, 30000)
	' @END

	On Error Resume Next
		
	iWait = 0	
	sBrowserStatus = False	
	
	While (sBrowserStatus = False) And (iWait < iNeedsToWait)
		sBrowserStatus = Browser("micClass:=Browser", "CreationTime:=" &iCreationTime).Page("micClass:=Page").Exist
		iWait = iWait + 1
	Wend
	
	iWait = 0
	
 	'Get Base web page object
	BaseWebPage iCreationTime
	
	cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
	sObjType = LCase(Trim(cObjPropArray(0)))
	Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
	
	If sObjPropAndValue("name")="" Then
		objName=sObjPropAndValue(0)
	Else
		objName=sObjPropAndValue("name")
	End If
	
	
	'While (Trim(UCase(sBrowserStatus)) <> "DONE" And Trim(sBrowserStatus) <> "") And (iWait < iNeedsToWait)
		'sBrowserStatus = Browser("micClass:=Browser", "CreationTime:=" &iCreationTime).WinStatusBar("micclass:=WinStatusBar").GetROProperty("regexpwndtitle")
		'If inStr(LCase(sBrowserStatus), "error") >0 Then
		'	iWait = iWait + 1
		'End If
	'Wend

	iWait = 0
	
	bSpecifiedObjectExist = False
	
   Select Case sObjType
		   Case "image"
               
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.Image(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for image: "& objName &" Verification In Browser"," '" &"Image: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for image: "& objName &" Verification In Browser","'" &"Image: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "webarea"
               
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.WebArea(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for webarea: "& objName &" Verification In Browser"," '" &"WebArea: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for webarea: "& objName &" Verification In Browser","'" &"WebArea: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "webedit"
               
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.WebEdit(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for WebEdit: " & objName & " Verification In Browser"," '" &"WebEdit: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for WebEdit: " & objName & " Verification In Browser","'" &"WebEdit: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "webbutton"
               
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.WebButton(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for WebButton: " & objName & " Verification In Browser"," '" &"WebButton: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for WebButton: " & objName & " Verification In Browser","'" &"WebButton: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "link"
               
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.Link(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS", "Waiting for Link: " & objName & " Verification In Browser"," '" &"Link: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for Link: " & objName & " Verification In Browser","'" &"Link: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "weblist"
               
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.WebList(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for WebList: " & objName & " Verification In Browser"," '" &"WebList: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for WebList: " & objName & " Verification In Browser","'" &"WebList: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "webradiogroup"
               
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.WebRadioGroup(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for WebRadioGroup: " & objName & " Verification In Browser"," '" &"WebRadioGroup: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for WebRadioGroup: " & objName & " Verification In Browser","'" &"WebRadioGroup: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "webcheckbox"
                
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.WebCheckBox(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for WebCheckBox: " & objName & " Verification In Browser"," '" &"WebCheckBox: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for WebCheckBox: " & objName & " Verification In Browser","'" &"WebCheckBox: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "webtable"
                
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.WebTable(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for WebTable: " & objName & " Verification In Browser"," '" &"WebTable: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for WebTable: " & objName & " Verification In Browser","'" &"WebTable: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "frame"
                
			   While (bSpecifiedObjectExist = False) And (iWait < iNeedsToWait)
					bSpecifiedObjectExist = BaseWindow.Frame(sObjPropAndValue).Exist(iNeedsToWait)
					iWait = iWait + 1
			   Wend
			   
				If bSpecifiedObjectExist Then
					ReportWriter "PASS","Waiting for Frame: " & objName & " Verification In Browser"," '" &"Frame: " & objName & "'  Existing In Browser: ",0
				Else
					ReportWriter "FAIL","Waiting for Frame: " & objName & " Verification In Browser","'" &"Frame: " & objName & "'  Not Existing In Browser: ",0
				End if

			Case "webelement"
				
				Found=0
				While (iWait < iNeedsToWait) And (Found=0)
					For Each objElement In Browser("CreationTime:=" &iCreationTime).Object.document.all
						If Trim(LCase(objElement.innerText))=Trim(LCase(sWebElement(1))) Then
							Found=1 
							Exit For
						End If
					Next
					iWait = iWait + 1
				Wend
           		If Found=1 Then
           			ReportWriter "PASS","Waiting for webelement: "&objName&" Verification In Browser","'"&"WebElement: " &objName&"' WebElement Existing In Browser: ",0
           		Else	
					ReportWriter "FAIL","Waiting for webelement: "&objName&" Verification In Browser","'"&"WebElement: " &objName&"' WebElement Not Existing In Browser: ",0
				End If

			Case Else
				ReportWriter "FAIL","UnKnown Object : " & sObjType, "'" & sObjType & "'  Not Found In Browser ",0
		End select
		   
		On Error GoTo 0
End Function

'*********************************************************************************************************************

Function GetDateInMonthNameFormat(iDate,sDateFormat)
		' @HELP
		' @group	: 	WebControls
		' @method	:	GetDateInMonthNameFormat(iDate,sDateFormat)
   		' @returns	:   Date in Required format     
   		' @parameter:   iDate      : Required date in "MM/DD/YYYY" Format
   		' @parameter:   sDateFormat: Required date format  (Ex:"DD-MMM-YYYY ->05-May-2008","MMM-DDD-YY ->May-Wed-2008")
		' @notes	:	Used to display required date in the required format
    	' @END 
	    Set sDateVal = GetObject("","Excel.Application")
		GetDateInMonthNameFormat = sDateVal.Text(iDate,sDateFormat)
		Set xDate = Nothing
 End Function
 Function SelectValueFromWeblist(sObjPropAndValue,sValue,iCreationTime)
'' @HELP
'		' @group	: WebControls
'		' @method	: SelectValueFromWeblist(sObjPropAndValue,sValue,iCreationTime)
'   		' @returns	: None
'   		' @parameter: iCreationtime  : Creationtime of Browser
'		' @parameter: sObjPropAndValue : Property name and value of WebTable
'		' @parameter: sValue	:  index of WebTable
'		' @notes	: Selects Value from WebList
'		' @example	: SelectValueFromWeblist("WebList;html id:=toFundsList","723314 CC_11537452_1 (100000 Credits)","0")
'		' @END
''   On Error Resume Next                                      
	sObjType = LCase(Trim(sObjType))
	
	cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
	sObjType = LCase(Trim(cObjPropArray(0)))
	Set sObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
	
	If sObjPropAndValue("name")="" Then
		objName=sObjPropAndValue(0)
	Else
		objName=sObjPropAndValue("name")
	End If
	'MsgBox BaseWindow.WebList(sObjPropAndValue).Exist(2)

	If BaseWindow.WebList(sObjPropAndValue).Exist Then
		iIndex=1
		
	     Setting.WebPackage("ReplayType") = 2 
	     BaseWindow.WebList(sObjPropAndValue).Click
		 Setting.WebPackage("ReplayType") = 1
		 
		
		Do Until StrComp (sVAL,sValue)=0 Or iIndex=10
        sVAL = BaseWindow.WebList(sObjPropAndValue).GetROProperty("Selection")
        Print sVAL&" : "&sValue
        Set oShell = CreateObject("WScript.Shell")
           
        
         If  StrComp (sVAL,sValue)<>0 Then
                 BaseWindow.WebList(sObjPropAndValue).Click
		    	 oShell.SendKeys "{DOWN}"
		    	 oShell.SendKeys "{ENTER}"
		    	wait 2
		 Else
		      oShell.SendKeys "{ENTER}"
		      Exit Function
		 End if
			
			
			
			iIndex=iIndex+1
			Set oShell = Nothing
		Loop
		If iIndex=5 Then
			ReportWriter  "FAIL" ,"Select From List Box: '" & sobjName ,"Option  '"&sValue&"' Is Not Selected From List Box:  '"& objName,0

		End If

		 ReportWriter  "PASS" ,"Select From List Box: '" & objName ,"Option  '"&sVAL&"' Is Selected From List Box:  '"& objName &"' ",0
	Else
  		 ReportWriter  "FAIL" ,"Select From List Box: '" & sobjName ,"Option  '"&sVAL&"' Is Not Selected From List Box:  '"& objName &"' ",0
	End If
			SelectValueFromWeblist = sVAL

End Function
'*************************************************************************************
Sub ClickWebElementInTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowId,iColID)

		' @HELP
		' @group	: WebControls
		' @method	: ClickWebElementInTableByRowAndColID(iCreationtime,sTblPropAndVal,iTblIndex,iRowId,iColID)
   		' @returns	: None
   		' @parameter: iCreationtime  : Creationtime of Browser
		' @parameter: sTblPropAndVal : Property name and value of WebTable
		' @parameter: iTblIndex	:  index of WebTable
		' @parameter: iRowID : Row ID in which data is to be compared
		' @parameter: iColID : Column ID in which data is to be compared
		' @notes	: Clicks WebElement present in WebTable based on specifed row and column ID
		' @example	: ClickLinkInTableByRowAndColID(0,"name:=role",0,3,4)
		' @END
		
		'Get Base web page object
		BaseWebPage iCreationTime
		If BaseWindow.WebTable(sTblPropAndVal,"index:="&iTblIndex).Exist(.1) Then
			Set oDesc = BaseWindow.WebTable(sTblPropAndVal,"index:="&iTblIndex).ChildItem(iRowId,iColID,"WebElement",INDEX_VALUE_ZERO)
			oDesc.Highlight
			Setting.WebPackage("ReplayType") = 2 'Mouse
			oDesc.Click
			Setting.WebPackage("ReplayType") = 1 'Events
			ReportWriter "PASS","WebElement","Clicked on the WebElement  " & oDesc.GetROProperty("name"),0
		Else
			ReportWriter "FAIL","WebTable","WebTable Doesn't exist",0
		End If	
End Sub
'*************************************************************************************

Function Bookings_InvokeBrowser (sURL, sBrowser,iBrowIndex)
		' @HELP
		' @group	:	WebControls	
		' @method	:	InvokeBrowser (sURL, sBrowser, iBrowIndex)
		' @returns	:	None
		' @parameter:   sURL		: Name of the URL to invoke
		' @parameter:   sBrowser	: Name of the Browser
		' @parameter:   iBrowIndex	: Index of the Browser
		' @notes	: 	Invoke the URL for the given browser
		' @END
		
		GetObjectDescriptions()
		If UCase(sBrowser) = "IE" Then
			SystemUtil.Run "iexplore.exe", sURL,"","",3
'			wait 8
			Browser(oSesObjDescriptions,"Creationtime:="&iBrowIndex).Page(oObjectDescriptions).Sync
			If Browser(oSesObjDescriptions,"Creationtime:="&iBrowIndex).Page(oObjectDescriptions).Exist(3) Then
				Reporter.ReportEvent micPass,"Browser", "URL:"&sURL&" Is Successfully Opened in the Browser"
            Else
				Reporter.ReportEvent micFail,"Failed to Open Browser","Unable to Open the Browser"
				'ExitRun
				sOpen = False
			End If
		ElseIf UCase(sBrowser) = "FF" Then
			SystemUtil.Run "firefox.exe", sURL
			wait 8
			If Browser(oSesObjDescriptions,"Creationtime:="&iBrowIndex).Page(oObjectDescriptions).Exist(3) Then
				Reporter.ReportEvent micPass,"Browser", "URL:"&sURL&" Is Successfully Opened in the Browser"
            Else
				Reporter.ReportEvent micFail,"Failed to Open Browser","Unable to Open the Browser"
				'ExitRun
				sOpen = False
			End If
		ElseIf UCase(sBrowser) = "CH" Then
			SystemUtil.Run "chrome.exe", sURL
			wait 8
			If Browser(oSesObjDescriptions,"Creationtime:="&iBrowIndex).Page(oObjectDescriptions).Exist(3) Then
				Reporter.ReportEvent micPass,"Browser", "URL:"&sURL&" Is Successfully Opened in the Browser"
            Else
				Reporter.ReportEvent micFail,"Failed to Open Browser","Unable to Open the Browser"
				'ExitRun
				sOpen = False
			End If
		Else
			SystemUtil.Run "netscp6.exe", sURL
			If Browser(oSesObjDescriptions,"Creationtime:="&iBrowIndex).Page(oObjectDescriptions).Exist(3) Then
				Reporter.ReportEvent micPass,"Browser", "URL:"&sURL&" Is Successfully Opened in the Browser"
            Else
				Reporter.ReportEvent micFail,"Failed to Open Browser","Unable to Open the Browser"
				'ExitRun
				sOpen = False
			End If
		End If
End Function
'**************************************
Sub WaitForObeject(sObjPropAndValue)
   	' @group                		: Common
	' @method               		: WaitForObeject()
	' @returns              		: None
	' @parameter: sObjPropAndValue : Object Property & Value (ex: micClass:=Browser)
	' @notes                		: Wait for object
	' @END
	wait INDEX_VALUE_TWO
    Set obj=Description.Create
	obj("micclass").value = "Browser"
	Set aObj=Desktop.ChildObjects(obj)
	BrowserIndex=aObj.count-1
	If aObj.count>=0 Then		
		Set oPage = Browser("Creationtime:="&BrowserIndex).page("micclass:=page")
		oPage.Sync
		Else
        ReportWriter "Fail", "operate:  '"&sObjType&"'","WebPage  Is Not Identified",INDEX_VALUE_ZERO
		Exit Sub
	End If
   cObjPropArray=Split(sObjPropAndValue,";",INDEX_VALUE_TWO)
   sObjType = LCase(Trim(cObjPropArray(0)))
   Set oObjPropAndValue=CreateObjectDescription(cObjPropArray(1))
   WaitTime=0
   Select Case sObjType
		Case "webedit"
			Do 
					If oPage.WebEdit(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case "weblist"
			Do 
					If oPage.weblist(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case "webradiogroup"
			Do 
					If oPage.webradiogroup(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case "webcheckbox"
			Do 
					If oPage.webcheckbox(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case "webelement"
			Do 
					If oPage.webelement(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case "link"
			Do 
					If oPage.link(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case "image"
			Do 
					If oPage.image(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case "webtable"
			Do 
					If oPage.webtable(oObjPropAndValue).Exist Then
					Exit Do
					ElseIf WaitTime=5 Then
					Exit Do
					Else
					WaitTime=WaitTime+1
					End if
		    Loop
		Case Else
				ReportWriter "Fail", "operate:  '"&sObjType&"'","'"&sObjType&"'  Is Not Identified",INDEX_VALUE_ZERO
	End Select
End Sub
'*************************************************