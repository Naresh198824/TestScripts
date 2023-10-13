




'**********************************************Script Description*************************************************************
'Script Name							   			- 
Rem Purpose / Description    		  - 
Rem ********************************************************************************************************************************
Rem **********************************************Start of Framework initialization**********************************************

 @@ script infofile_;_ZIP::ssf507.xml_;_

Set fso=CreateObject("scripting.filesystemobject")
pFolderPath=Environment("TestDir")
pCaseName=Environment("TestName")
pPath= split (pFolderPath, "IPTool\Scripts\"&pCaseName)(0)
pFolderName=fso.GetParentFolderName(pFolderPath)


ErrComponentCreation sTestCaseID

''********************************************** Initialize Required Classes **************************************************
sSQL = "SELECT  * from [Scenarios$] where ExecutionTag='T' order by SNO Asc"
oData  = oDataAccess.GetSingleRowValuefromAccessDB (sSQL)

''********************************************** Add Common Variables **************************************************

If UBound(oData)>=0 Then

	For FCount = 0 to Ubound(oData)
		'Fetch data from test data table
		sSQL_GETRECORD = "SELECT  * from [IOData$] where TestDataTableLinkID='"&oData(FCount, 3)&"'"
		Set InputData = oDataAccess.ExecSQLStatementWithWhereClass (sSQL_GETRECORD)
		FunName = oData(FCount, 2)
		TestName = oData(FCount, 4)
		fnName = FunName & "(InputData)"
	'	MsgBox TestName
	'	MsgBox fnName
		ReportStart TestName
	
		Execute fnName
	
		ReportEnd	
	Next
Else
	ReportWriter "Fail","Test Data table", "No T Flag was set to execute ",0
End IF
' **********************************************Release Object Variables****************************************************

Set oDataAccess  = Nothing @@ hightlight id_;_1052360_;_script infofile_;_ZIP::ssf489.xml_;_



'Browser("name:=Logon").Page("title:=Logon").WebEdit("name:=sap-user").SetSecure EXECUTION_ENVIRONMENT_3

'
''Browser("name:=Home").Page("title:=Home").WebButton("acc_name:=Home - Show All My Apps").Highlight
'
'Browser("name:=Home").Page("title:=Home").WebButton("acc_name:=Home - Show All My Apps").Click
'Wait 5
'
'
'value = "Sales - Sales Order Processing"
'
'
''Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesList-listUl").WebElement("innerhtml:="&MyApp).Highlight
'
'Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesList-listUl").WebElement("innerhtml:="&MyApp).Click
'
'
'
'''msgbox Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesL.*").GetAllROProperties  
'
'
'
''Browser("Home").Page("Home").WebElement("Sales - Sales Order Processing").Click
''v = Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesL.*").GetItem(3)
''print v
''
''
''set gv = Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesL.*").ChildObjects
''
''For i = 0 To gv.count
''	val = Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesL.*").GetItem(i)
'''	print val
''	If val="Sales - Sales Order Processing" Then
''		Browser("name:=Home").Page("title:=Home").WebList("html id:=sapUshellAllMyAppsDataSourcesL.*").Select i-1
''		Exit for
''	End If
''	
''	
''Next
'
'
'
'
'Wait 3
'Browser("name:=Home").Page("title:=Home").WebList("html id:=oItemsContainerlist-listUl").Select "Create Sales OrdersVA01"
'
'wait 3
''
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPEdit("logical name:=Order Type").Highlight
'Wait 1
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPEdit("logical name:=Order Type").Set "OR"
'
'
'
'
'
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Sales Organization").SAPEdit("logical name:=Sales Organization").Highlight
'
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Sales Organization").SAPEdit("logical name:=Sales Organization").Set "ABCB"
'
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Distribution Channel").SAPEdit("logical name:=Distribution Channel").Set "A1"
'
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Division").SAPEdit("logical name:=Division").Set "B1"

'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Sales office").SAPEdit("logical name:=Sales office").Highlight
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Sales office").SAPEdit("logical name:=Sales office").Set "B1"
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Sales group").SAPEdit("logical name:=Sales group").Highlight
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("innertext:=Sales group").SAPEdit("logical name:=Sales group").Set "GSR"
'
'
'
'
'
''
''
''Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Sold-To Party").SAPEdit("logical name:=Sold-To Party").Highlight
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Sold-To Party").SAPEdit("logical name:=Sold-To Party").Set "1000014"
''Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Highlight
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Set "1000014"
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Customer Reference").SAPEdit("logical name:=Cust. Reference").Highlight @@ script infofile_;_ZIP::ssf525.xml_;_
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Customer Reference").SAPEdit("logical name:=Cust. Reference").Set "Test"
''Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPEdit("logical name:=Order Type").Click
''Wait 1
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPEdit("logical name:=Order Type").Set "OR"
'
'
'
'abc = Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("name:=Insert in Personal List").RowCount 
'MsgBox abc
'val = Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("name:=Insert in Personal List").ColumnCount
'MsgBox val
''For i = 0 To abc Step 1
''	Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("name:=Insert in Personal List").ColumnCount
''Next
''
''xyz = Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").WebTable("name:=Insert in Personal List").GetCellData (1,1)
''MsgBox xyz
'
''
''Browser("Create Sales Documents").Page("Create Sales Documents").SAPFrame("Create Sales Documents").SAPEdit("Order Type").o
''Browser("Create Sales Documents").Page("Create Sales Documents").SAPFrame("Create Sales Documents").SAPEdit("Order Type").OpenPossibleEntries @@ script infofile_;_ZIP::ssf512.xml_;_
''Browser("Create Sales Documents").Page("Create Sales Documents").SAPFrame("Create Sales Documents").SAPTable("SAPTable").SelectCell 19,1 @@ script infofile_;_ZIP::ssf513.xml_;_
''Browser("Create Sales Documents").Page("Create Sales Documents").SAPFrame("Create Sales Documents").SAPButton("Copy").Click @@ script infofile_;_ZIP::ssf514.xml_;_
''Browser("Create Sales Documents").Page("Create Sales Documents").SAPUIButton("Back").Click @@ script infofile_;_ZIP::ssf515.xml_;_
''Browser("Create Sales Documents").Page("Create Sales Documents").SAPFrame("Create Sales Documents").SAPTable("SAPTable").SetCellData 19,1,"ON" @@ script infofile_;_ZIP::ssf516.xml_;_
'
'
'
'
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPEdit("logical name:=Order Type").Click
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPEdit("logical name:=Order Type").OpenPossibleEntries
'wait 3
'MyApp = "RK"
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPTable("Class Name:=SAPTable").WebElement("outertext:="&MyApp).Highlight
''MsgBox xyz
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPTable("Class Name:=SAPTable").WebElement("outertext:="&MyApp).Click
'
'MyAp = "OR"
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPTable("Class Name:=SAPTable").WebElement("outertext:="&MyApp).Highlight
''MsgBox xyz
'Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPTable("Class Name:=SAPTable").WebElement("outertext:="&MyAp).Click

'z = Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPTable("Class Name:=SAPTable").GetCellData (8,2)
'MsgBox z
'
'y = Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPTable("Class Name:=SAPTable").GetCellData (10,1)
'MsgBox y
'
'x = Browser("name:=Create Sales Documents").Page("title:=Create Sales Documents").SAPTable("Class Name:=SAPTable").GetCellData (10,2)
'MsgBox x

 @@ script infofile_;_ZIP::ssf523.xml_;_
'Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebTable("innertext:=Terms of payment key").SAPEdit("logical name:=Pyt Terms").Highlight
'Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebTable("innertext:=Terms of payment key").SAPEdit("logical name:=Pyt Terms").Set "001"
'
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Customer Reference").SAPEdit("logical name:=Cust. Reference").Set "Test"
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Sold-To Party").SAPEdit("logical name:=Sold-To Party").Set "1000014"
''Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Highlight
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Set "1000014"
'
'Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebElement("html id:=wnd[0]/sbar_msg-txt").Highlight
'value = Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebElement("html id:=wnd[0]/sbar_msg-txt").getvisibletext("text")'                        GetROProperty("Text") 
'MsgBox value


'Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebTable("innertext:=Terms of payment key").SAPEdit("logical name:=Pyt Terms").Highlight
''Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebTable("innertext:=Terms of payment key").SAPEdit("logical name:=Pyt Terms").Set "001"
'Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebElement("outertext:=Please enter.*").Highlight
'abc =Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebElement("outertext:=Please enter.*").GetROProperty("outertext") 
'MsgBox abc

















	
