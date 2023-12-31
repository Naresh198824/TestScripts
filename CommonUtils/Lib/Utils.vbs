'*********************************************************************************************************************
'**  Project            : 
'**  File Name          : Utils.vbs
'**  Version            : 
'**  Created on         : 
'**  Updated on         : 
'**  Description        : Class and methods related to different functions
'**  Copyright          : Arsin Corporation.
'**  Author             :       
'*********************************************************************************************************************

'Class clsUtils

Function TimeToSeconds(dTime)
	
		' @HELP
		' @class	: clsUtils
		' @method	: TimeToSeconds(dTime)
		' @returns	: TimeToSeconds: Returns seconds
		' @parameter: dTime: Any date data type or string expression in hh:mm:ss format
		' @notes	: Converts time to secods.
		' @END
	
		Dim iSec
		Dim iMin
		Dim iHr
		Dim iTArr
		Dim bIsNum
		
		If isdate(dTime) then
			iSec=Second(dTime)
			iMin=Minute(dTime)
			iHr=Hour(dtime)
			TimeToSeconds=iSec+iMin*60+iHr*60*60
		else
			iTArr=split(dTime,":",3)
			bIsNum=true
			for i=0 to ubound(iTArr) 
				if not IsNumeric(iTArr(i)) then
					bIsNum=False
					exit for 
				end if
			Next
			if ubound(iTarr)>=2  and bIsNum then
				iSec=clng(iTArr(2))
				iMin=clng(iTArr(1))
				iHr=clng(iTArr(0))
			elseif ubound(iTarr)>=1 and bIsNum then
				iSec=0
				iMin=clng(iTArr(1))
				iHr=clng(iTArr(0))
			End IF	
			If (iSec>=0 and iSec<=60) and (iMin>=0 and iMin<=60) and (iHr>=0) and (ubound(iTarr)<=2 and ubound(iTarr)>=1 and bIsNum) then
				TimeToSeconds=iSec+iMin*60+iHr*60*60
			Else
				Reporter.ReportEvent micFail, "Wrong Usage: SecondsToTime", "Time Expression Not In format."
				ExitActionIteration 
			End If
		End If
End Function
'--------------------------------------------------------------------------------------------------------------------
Function SecondsToTime(iSecs)
	
		' @HELP
		' @class	: clsUtils
		' @method	: SecondsToTime(iSecs)
		' @returns	: SecondsToTime: Returns a string expression in hh:mm:ss format.
		' @parameter: iSecs: Seconds in integer data type
		' @notes	: Converts seconds to time expression.
		' @END
		
		Dim iSec
		Dim iMin
		Dim iHr
		
		SecondsToTime=""
		If Not isNumeric(iSecs) Then
			Reporter.ReportEvent micFail, "Wrong Usage: SecondsToTime", "This function allows only numerics." 
			ExitActionIteration 
			Exit Function
		End If	
		
		SecondsToTime=""
		iSecs=Clng(ISecs)
		iSec=iSecs Mod 60
		iSecs=iSecs\60
		iMin=iSecs Mod 60
		iSecs=iSecs\60
		iHr=iSecs
		
		If len(Cstr(iHr))<2 Then
			SecondsToTime=SecondsToTime & "0" & Cstr(iHr)
		Else
			SecondsToTime=SecondsToTime & Cstr(iHr)
		End If
		If len(Cstr(iMin))<2 Then
			SecondsToTime=SecondsToTime & ":0" & Cstr(iMin)
		Else
			SecondsToTime=SecondsToTime & ":" & Cstr(iMin)
		End If
		If len(Cstr(iSec))<2 Then
			SecondsToTime=SecondsToTime & ":0" & Cstr(iSec)
		Else
			SecondsToTime=SecondsToTime & ":" & Cstr(iSec)
		End If
			
End Function
'--------------------------------------------------------------------------------------------------------------------
Function RandomString(iLength)
		
		' @HELP
		' @class	: clsUtils
		' @method	: RandomString(iLength)
		' @returns	: RandomString: Returns a string.
		' @parameter: iLength: Length of string, integer data type
		' @notes	: Returns a random string , length of iLength.
		' @END
		
		Dim iIndex
		Dim sTemp
		
		sTemp=""
		If not isNumeric(iLength) then
			Reporter.ReportEvent micFail, "Wrong Usage: RandomString", "This Function allows only numarics."
			ExitActionIteration 
			Exit Function
		End If
		
		Randomize
		For iIndex=1 To iLength 
			sTemp=sTemp & Chr(Int((62 * Rnd) + 65) )
		Next
		
		RandomString=sTemp
End Function
'--------------------------------------------------------------------------------------------------------------------
Function RandomNumber(iLength)
	
		' @HELP
		' @class	: clsUtils
		' @method	: RandomNumber(iLength)
		' @returns	: RandomNumber: Returns a string of numbers.
		' @parameter: iLength: Length of string, integer data type
		' @notes	: Returns a random string numbers , length of iLength.
		' @END
	
		Dim iIndex
		Dim sTemp
		
		sTemp=""
		If not isNumeric(iLength) then
			Reporter.ReportEvent micFail, "Wrong Usage: RandomNumber", "This Function allows only numarics."
			ExitActionIteration 
			Exit Function
		End If
		
		Randomize
		For iIndex=1 To iLength 
			sTemp=sTemp & Chr(Int((10 * Rnd) + 48) )
		Next
		RandomNumber=sTemp
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function VerifyItemInList(oWeblst,sItem)
	
		' @HELP
		' @class	: clsUtils
		' @method	: VerifyItemInList(oWeblst,sItem)
		' @returns	: VerifyItemInList: Boolean.
		' @parameter: oWeblst: HtmlList object
		' @parameter: sItem: String Item to be verified in the list
		' @notes	: Reports pass if item found in the list otherwise reports fail.
		' @END
	
		Dim iIndex
		Dim iLstCount
		Dim sTypeName 
		
		iLstCount=oWeblst.GetROProperty("Items count")
		VerifyItemList=False
		sTypeName =TypeName (oWeblst)
		
		If sTypeName  <> "HtmlList" Then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyItemInList", "Cannot call this function on the object of type: " & sTypeName & "."
			ExitActionIteration
		End If
		
		For iIndex=1 To iLstCount
			If oWeblst.GetItem(iIndex)=sItem Then
				VerifyItemsList=True
				Reporter.ReportEvent micPass, "VerifyItemInList", "Item: " & sItem  & " found in the list."  
				exit for
			end if
		Next
		
		If iIndex>iLstCount Then
			Reporter.ReportEvent micFail, "VerifyItemInList", "Item: " & sItem  & " not found in the list." 
		End if
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function VerifyItemsInList(oWeblst,arItems)
	
		' @HELP
		' @class	: clsUtils
		' @method	: VerifyItemsInList(oWeblst,sItem)
		' @returns	: VerifyItemsInList: Boolean.
		' @parameter: oWeblst: HtmlList object
		' @parameter: arItems: Array of item to be verified in the list
		' @notes	: Reports pass if all the items in the array found in the list else reports fail.
		' @END
	
		Dim iOuter
		Dim iInner
		Dim sTypename
		
		IsArrayInList=False
		sTypename=TypeName (oWeblst)
		
		If  sTypename <> "HtmlList" Then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyItemsInList", "Cannot call this function on the object of type: " & sTypeName & "."
			ExitActionIteration
		End If
		
		sTypename=TypeName(arItems)
		If isArray(arItems)  Then
			For iOuter=Lbound(arItems) To Ubound(arItems)
				VerifyItemsInList=False
		
				For iInner=1 To oWeblst.GetROProperty("Items count")
					If oWeblst.GetItem(iInner)=arItems(iOuter) Then
						VerifyItemsInList=True
						Exit For
					End If
				Next
		
				If Not VerifyItemsInList Then
					Reporter.ReportEvent micFail, "VerifyItemsInList", "Item: " & arItems(iOuter) & " not found in the list." 
					ExitActionIteration
					Exit For
				End If
			Next
		
			If VerifyItemsInList Then
				Reporter.ReportEvent micPass, "VerifyItemsInList", "Items found in the list."  
	 		End If
		Else
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyItemsInList", "Cannot call this function on the object of type: " & sTypeName & "."
			ExitActionIteration
		End if
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DeleteCookie(oDocument,sCookiename)
	
		' @HELP
		' @class	: clsUtils
		' @method	: DeleteCookie(oDocument,sCookiename)
		' @returns	: DeleteCookie: None.
		' @parameter: oDocument: HtmlDocument object
		' @parameter: sCookiename: Name of Cookie
		' @notes	: Deletes the cookie, with respective to document for given name.
		' @END
		
		Dim sTypename
		
		sTypename=TypeName(oDocument)
		If Not (TypeName(oDocument) = "HTMLDocument") Then
			Reporter.ReportEvent micFail, "Wrong Usage: DeleteCookie", "Cannot call this function on the object of type: " & sTypename & "."
			ExitActionIteration
			Exit Function
		End If
		
		If (oDocument.Cookie = "") Then
			Reporter.ReportEvent micWarning, "DeleteCookie", "No Cookeis found."
			DeleteCookie=False
		Else

			If (InStr(oDocument.Cookie,sCookiename)>0) Then
				oDocument.Cookie =sCookiename & "=" &  "NULL;expires=Thursday, 29-Feb-96 12:00:00 GMT"
				DeleteCookie=True
			Else
				Reporter.ReportEvent micWarning, "DeleteCookie", "Cookie not found with the name:" & sCookiename
			 	DeleteCookie=False
			End If

		End If
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function VerifyURL(oWebpage,sUrlstr)
	
		' @HELP
		' @class	: clsUtils
		' @method	: VerifyURL(oWebpage,sUrlstr)
		' @returns	: VerifyURL: True/False.
		' @parameter: oWebpage: Page(step)object
		' @parameter: sUrlstr: URL string
		' @notes	: Reports pass if the URL matches with page, else reports fail.
		' @END
			
		Dim sTypename
		
		VerifyURL=False
		sTypename=TypeName(oWebpage)
		If Not TypeName(oWebpage)="Step" then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyURL", "Cannot call this function on the object of type: " & sTypename & "."
			ExitActionIteration
		End If
		
		sTypename=TypeName(sUrlstr)
		If Not sTypename="String" then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyURL", "Cannot call this function on the data type: " & sTypename & ", for sUrlstr."
			ExitActionIteration
		End If
			
		If oWebpage.GetROProperty("url")= sUrlstr then 
			VerifyURL=True
			Reporter.ReportEvent micPass, "VerifyURL", "Actual URL matched with expected url " & sUrlstr & "."
		Else
			Reporter.ReportEvent micFail, "VerifyURL", "Actual URL not matched with expected url " & sUrlstr & "."
		End If
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function IsPageNotFound(oWebpage)
		
		' @HELP
		' @class	: clsUtils
		' @method	: IsPageNotFound(oWebpage)
		' @returns	: IsPageNotFound: True/False
		' @parameter: oWebpage: Page(step)object
		' @notes	: Reports fail if page not found(HTTP 404 Not Found) occurs
		' @END
		
		Dim sTypename
		
		IsPageNotFound=False
		
		sTypename=TypeName(oWebpage)
		
		If sTypename<>"Step" then
			Reporter.ReportEvent micFail, "Wrong Usage: IsPageNotFound", "Cannot call this function on the object of type: " & sTypename & "."
			ExitActionIteration
			Exit Function
		End If
		
		If oWebpage.GetROProperty("title")="HTTP 404 Not Found" Then 
			Reporter.ReportEvent micFail, "IsPageNotFound", "Page not found : HTTP 404 Not Found."
			IsPageNotFound=True
		End If	
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function IsPageFailed(oWebpage,arFailCases)

		' @HELP
		' @class	: clsUtils
		' @method	: IsPageFailed(oWebpage,arFailCases)
		' @returns	: IsPageFailed: True/False
		' @parameter: oWebpage: Page(step)object
		' @parameter: arFailCases : Array of Fail cases(ex: 'HTTP 404 Not Found','HTTP 501 Internal Server Error')
		' @notes	: Reports fail if page failed and matched in fail cases else this function reports pass.
		' @END
		
		Dim sTitle
		Dim iIndex
		Dim sTypename
		
		IsPageFailed=False
		sTypename=TypeName(oWebpage)
		
		If sTypename<>"Step" then
			Reporter.ReportEvent micFail, "Wrong Usage: IsPageFailed", "Cannot call this function on the object of type: " & sTypename & "."
			ExitActionIteration
			Exit Function
		End If
		
		sTypename=TypeName(arFailCases)
		If not isArray(arFailCases) then
			Reporter.ReportEvent micFail, "Wrong Usage: IsPageFailed", "Cannot call this function on the object of type: " & sTypename & ",for arFailCases."
			ExitActionIteration
			Exit Function
		End If	
		
		sTitle=oWebpage.GetROProperty("title")
		For iIndex=lbound(arFailCases) to ubound(arFailCases)
			If sTitle=arFailCases(iIndex) then 
				IsPageFailed=True
				Reporter.ReportEvent micFail, "IsPageFailed", "Page failed :"  &  arFailCases(iIndex) & "."
				Exit For
			End if
		Next
		
		If IsPageFailed=False then
			Reporter.ReportEvent micPass, "IsPageFailed", "No page failures occured."
		End If
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function CompareValues(vVal1,vVal2)
		
		' @HELP
		' @class	: clsUtils
		' @method	: CompareValues(vVal1,vVal2)
		' @returns	: CompareValues: True/False
		' @parameter: vVal1: value of any premitive data type(String,Intger,double etc.)
		' @parameter: vVal2: value of any premitive data type(String,Intger,double etc.)
		' @notes	: Reports pass if string pattern found , else this function reports fail.
		' @END
	
		Dim sTypname1
		Dim sTypename2
		
		sTypname1=TypeName(vVal1)
		sTypname2=TypeName(vVal2)
	
		CompareValues=False
		If Not (isObject(vVal1) Or isObject(vVal2) Or isArray(vVal1) Or isArray(vVal2)) Then
			If CStr(vVal1)=CStr(vVal2) then
				Reporter.ReportEvent micPass, "CompareValues", "The two values  '" & CStr(vVal1) & "' and '" &  CStr(vVal2) & "' are equal."
				CompareValues=True
			Else
				Reporter.ReportEvent micFail, "CompareValues", "The two values  '" & CStr(vVal1) & "' and '" &  CStr(vVal2) & "' are not equal."
				CompareValues=False
			End if
		Else
			Reporter.ReportEvent micFail, "CompareValues","Cannot call this function on the object of types: " & sTypename1 & "and " & sTypename2 & "."
			ExitActionIteration
		End If
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function CompareRegularExp(sMainStr,sPattern)
	
		' @HELP
		' @class	: clsUtils
		' @method	: CompareRegularExp(sMainStr,sPattern)
		' @returns	: CompareRegularExp: True/False
		' @parameter: sMainStr: String being searched in
		' @parameter: sPattern: String pattern being searched for
		' @notes	: Reports pass if string pattern found , else this function reports fail.
		' @END
		
		Dim objRegExp
		Dim sTypname1
		Dim sTypname2
		Dim iPos
		
		sTypename1=TypeName(sMainStr)
		sTypename2=TypeName(sPattern)
		CompareRegularExp=False
		
		If sTypename1<>"String" Or sTypename2<>"String" Then
			Reporter.ReportEvent micFail, "Wrong Usage: CompareRegularExp","Cannot call this function on the object of types: " & sTypename1 & " and " & sTypename2  & "."
			ExitActionIteration
		End if
		
		Set objRegExp = New RegExp
		With objRegExp
			.Pattern = sPattern
			.IgnoreCase = True
			.Global = True
		End With
		
		iPos=objRegExp.Test(sMainStr)
		If iPos Then
			CompareRegularExp=True
			Reporter.ReportEvent micPass, "CompareRegularExp", "The pattern " & sPattern &  " found."
		Else
			CompareRegularExp=False
			Reporter.ReportEvent micFail, "CompareRegularExp", "The pattern " & sPattern &  " not found."
		End if
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function VerifyMaxLengthNumber(oWebedit,iInLen)
	
		' @HELP
		' @class	: clsUtils
		' @method	: VerifyMaxLengthNumber(oWebedit,iInLen)
		' @returns	: VerifyMaxLengthNumber: True/False
		' @parameter: oWebedit: HtmlEdit Object
		' @parameter: iInLen: Length to be verified, integer value
		' @notes	: Reports pass if object maxlength matched with iInLen, else this function reports fail. 
		'			  This function writes a number string length of iInLen.
		' @END
		
		Dim iIndex
		Dim sStrNum
		Dim iActMaxLen
		Dim sTypeName
		
		VerifyMaxLengthNumber=False
	
		sTypeName=TypeName(oWebedit) 
		If sTypeName<>"HtmlEdit" Then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyMaxLengthNumber","Cannot call this function on the object of types: " & sTypename & "."
			ExitActionIteration
		End If
		sTypeName=TypeName(iInLen) 
		
		If sTypeName<>"Integer" Then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyMaxLengthNumber","Cannot call this function on the object of types: " & sTypename & ", for iInLen."
			ExitActionIteration
		End If
	
		sStrNum=""
		iActMaxLen=oWebedit.Object.MAXLENGTH
		sStrNum = RandomNumber(iInLen+1)
		If iActMaxLen>=len(sStrNum) Then
			oWebedit.Set sStrNum
		Else
			oWebedit.Set Left(sStrNum,iActMaxLen)
		End If
	
		If iActMaxLen=iInLen then
			VerifyMaxLengthNumber=True
			Reporter.ReportEvent micPass, "VerifyMaxLengthNumber", "Actual maximum length of the object  '" & iActMaxLen & "', is equal to expected length '" & iInLen & "'."
		Else
			Reporter.ReportEvent micFail, "VerifyMaxLengthNumber", "Actual maximum length of the object is '" & iActMaxLen & "." 
		End If	
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function VerifyMaxLengthString(oWebedit,iInLen)
	
		' @HELP
		' @class	: clsUtils
		' @method	: VerifyMaxLengthString(oWebedit,iInLen)
		' @returns	: VerifyMaxLengthString: True/False
		' @parameter: oWebedit: HtmlEdit Object
		' @parameter: iInLen: Length to be verified, integer value
		' @notes	: Reports pass if object maxlength matched with iInLen, else this function reports fail.
		'			  This function writes a character string length of iInLen.
		' @END
		
		Dim sStrChar
		Dim iActMaxLen
		
		VerifyMaxLengthString=False
	
		sTypeName=TypeName(oWebedit) 
		If sTypeName<>"HtmlEdit" Then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyMaxLengthString","Cannot call this function on the object of types: " & sTypename & "."
			ExitActionIteration
			Exit Function
		End If
		
		sTypeName=TypeName(iInLen) 
		If sTypeName<>"Integer" Then
			Reporter.ReportEvent micFail, "Wrong Usage: VerifyMaxLengthString","Cannot call this function on the object of types: " & sTypename & ", for iInLen."
			ExitActionIteration
			Exit Function
		End If
	
		sStrChar=""
		iActMaxLen=oWebedit.Object.MAXLENGTH
		sStrChar=RandomString(iInLen+1)
		If iActMaxLen>=len(sStrChar) Then
			oWebedit.Set sStrChar
		else
			oWebedit.Set Left(sStrChar,iActMaxLen)
		End If
	
		If iActMaxLen=iInLen then
			VerifyMaxLengthString=True
			Reporter.ReportEvent micPass, "VerifyMaxLengthString","Actual maximum length of the Object  '" & iActMaxLen & "', is equal to expected length '" & iInLen & "'."
		Else
			Reporter.ReportEvent micFail, "VerifyMaxLengthString", "Actual maximum length of the object is '" & iActMaxLen & "." 
		End If	
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function CopyWebPageText(oWebpage,sFPath)
	
		' @HELP
		' @class	: clsUtils
		' @method	: CopyWebPageText(oWebpage,sFPath)
		' @returns	: CopyWebPageText: None
		' @parameter: oWebpage: Page(Step) Object.
		' @parameter: sFPath: Path of destination file.
		' @notes	: Creates a text file and wites to file from the page text.
		' @END
		
		Dim oFSO
		Dim oTFile
		Dim sTypename1
		Dim sTypename2
		
		sTypename1=Typename(oWebpage)
		sTypename2=Typename(sFPath)
		
		If sTypename1 <> "Step" Or sTypename2 <> "String" Then
			Reporter.ReportEvent micFail, "Wrong Usage: CopyWebPageText","Cannot call this function on the object of types: " & sTypename1 & " and " & sTypename2 & "."
			ExitActionIteration
			Exit function
		End If
		
		Set oFSO=CreateObject("Scripting.FileSystemObject")
		Set	oTFile=oFSO.CreateTextFile(sFPath,True)
		oTFile.Write oWebpage.GetROProperty("innertext")
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function SearchStrInPage(oWebpage,sInput)
	
		' @HELP
		' @class	: clsUtils
		' @method	: SearchStrInPage(oWebpage,sInput)
		' @returns	: SearchStrInPage: True/False
		' @parameter: oWebpage: Page(Step) Object.
		' @parameter: sInput: String being searched for.
		' @notes	: Reports pass if sring found in page text else this function reports fail.
		' @END
		
		Dim sTypename1
		Dim sTypename2
		Dim iPos 
		
		sTypename1=Typename(oWebpage)
		sTypename2=Typename(sInput)
		
		If sTypename1 <> "Step" Or sTypename2 <> "String" Then
			Reporter.ReportEvent micFail, "Wrong Usage: SearchStrInPage","Cannot call this function on the object of types: " & sTypename1 & " and " & sTypename2 & "."
			ExitActionIteration
			Exit function
		End If
		
		SearchStrInPage=False
		iPos=instr(oWebpage.GetROProperty("innertext"),sInput)
		
		If ipos  Then
			SearchStrInPage=True
			Reporter.ReportEvent micPass, "SearchStrInPage","String '" & sInput & "' found at position " & iPos & " ."
		Else
			SearchStrInPage=True
			Reporter.ReportEvent micFail, "SearchStrInPage","String '" & sInput & "' not found." 
		End If	
		
End Function
'--------------------------------------------------------------------------------------------------------------------
Function DeleteDomainCookies(sDomainName,sBrowserType)
	
		' @HELP
		' @class	: clsUtils
		' @method	: DeleteDomainCookies(sDomainName,sBrowserType)
		' @returns	: DeleteDomainCookies: none
		' @parameter: sDomainName: Name of cookie domain Ex: yahoo.com.
		' @parameter: sBrowserType: Name of the Browser Ex:'IE 6.0','NetScape 7.x'
		' @notes	: Delete cookies from cookie folder/file with given domain name.
		' @END
		
		Dim objFSO
		Dim objFolder
		Dim collFile
		Dim objFileR
		Dim objFileW
		Dim sCookiePath
		Dim sCookie
		Dim sCookies
		Dim iCookieEnd
		Dim sLine
		Dim bFound
		Dim p1
		Dim iNoFounds
		Dim sTypename1
		Dim sTypename2
		
		sTypename1=Typename(sDomainName)
		sTypename2=Typename(sBrowserType)
		If sTypename1 <> "String" Or sTypename2 <> "String" Then
			Reporter.ReportEvent micFail, "Wrong Usage: DeleteDomainCookies" ,"Cannot call this function on the object of types: " & sTypename1 & " and " & sTypename2 & "."
			ExitActionIteration
			Exit function
		End If
		
		Set objFSO=CreateObject("Scripting.FileSystemObject")
	
		'Getting Cookies Folder where IE stores the cookies
		sCookiePath= objFSO.GetAbsolutePathName(objFSO.getSpecialFolder(2))
	
		p1=0
		p1=Instr(sCookiePath,"\")
		if p1 then
			p1=Instr(p1+1,sCookiePath,"\")
			if p1 then
				p1=Instr(p1+1,sCookiePath,"\")
			end If
		End If	
		If sBrowserType="IE 5.5" Or  sBrowserType="IE 6.0" Then
	  		if p1 then 
				sCookiePath=left(sCookiePath,p1)& "Cookies\"
			else
				msgbox "Path not found"
			end if
	
		'Getting the all the text file & and search for cookies
			Set objFolder=objFSO.GetFolder(sCookiePath)
			Set collFile=objFolder.Files
			For each objFile in collFile
				If lcase(objFile.Name)<>"index.dat" then
					Set objFileR=objFSO.OpenTextFile(objFile.Path ,1)
					iCookieEnd=1
					bFound=False
					sCookies=""
					iNoFounds=0
					do while not objFileR.AtEndOfStream 
						sLine=objFileR.ReadLine
						If iCookieEnd=9  Then
							If not bFound Then
								sCookie=sCookie & sLine  & vbcrlf
								sCookies=sCookies+sCookie
							End If	
							iCookieEnd=0
							sCookie=""
							bFound=False
						ElseIf iCookieEnd=3 Then 
							If InStr(sLine,sDomainName) Then
								bFound=True
								iNoFounds=iNoFounds+1
							end if
							sCookie=sCookie & sLine  & vbcrlf
						Else 
							sCookie=sCookie & sLine & vbcrlf
						End IF
						iCookieEnd=iCookieEnd+1
					loop
					objFileR.Close
	
		'If any Cookies found modifying the text file.
					If iNoFounds then
						set objFileW=objFSO.OpenTextFile(objFile.Path,2)
						objFileW.Write sCookies
						objFileW.Close
					End If	
	
				End If
			Next
		ElseIf  sBrowserType="NS 6.x" Or sBrowserType="NS 7.x" Then
			sCookiePath=left(sCookiePath,p1) & "Application Data\Mozilla\Profiles\default\ftxmhsog.slt\cookies.txt"
			Set objFileR=objFSO.OpenTextFile(sCookiePath,1)
			sCookies=""
			Do While not objFileR.AtEndOfStream 
				sLine=objFileR.ReadLine
				If not(Instr(sLine,sDomainName)=2) Then
					sCookies=SCookies & sLine & vbcrlf
				End If
			Loop
			objFileR.Close
			Set objFileW=objFSO.OpenTextFile(sCookiePath,2)
			objFileW.Write sCookies
		End If
		
End Function
'--------------------------------------------------------------------------------------------------------------------
'End Class
