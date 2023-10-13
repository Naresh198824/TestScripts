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


Function EnterSalesOrderDetails(InputData)
	
	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Sold-To Party").SAPEdit("logical name:=Sold-To Party").Set InputData("SoldToParty") ' "1000014"
	Set Keys = CreateObject("WScript.Shell")
		Keys.SendKeys("{ENTER}")
		Wait 1
	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Customer Reference").SAPEdit("logical name:=Cust. Reference").Set InputData("CustReference") '"Test"
	Set Keys = CreateObject("WScript.Shell")
	Keys.SendKeys("{ENTER}")
	Wait 1
	Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebTable("innertext:=Terms of payment key").SAPEdit("logical name:=Pyt Terms").Set InputData("Payment")   ' "0001"
	Set Keys = CreateObject("WScript.Shell")
		Keys.SendKeys("{ENTER}")
	Wait 1
	Dim oDesc,iCounter
 
	Set oDesc = Description.Create
 
	oDesc("micclass").value = "SAPTable"
 
	'Find all the SAPTables in a Page
 
	Set objChkBox = Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").ChildObjects(oDesc)
	'objChkBox(2).highlight 
	objChkBox(2).SetCellData 2,3,InputData("Material") '"102"
	objChkBox(3).SetCellData 2,5,InputData("Units") '"1"
	objChkBox(3).SetCellData 2,14,InputData("Plant") '"1710"	
	
	
End Function
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Sold-To Party").SAPEdit("logical name:=Sold-To Party").Set "1000014"
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Highlight
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Ship-to party").SAPEdit("logical name:=Ship-To Party").Set "1000014"
'	Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").WebTable("innertext:=Customer Reference").SAPEdit("logical name:=Cust. Reference").Set "Test"
'Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebTable("innertext:=Terms of payment key").SAPEdit("logical name:=Pyt Terms").Set "0001"
'Set Keys = CreateObject("WScript.Shell")
'			Keys.SendKeys("{ENTER}")
'			
'
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("logical name:=All Items").SelectCell 2,3 @@ script infofile_;_ZIP::ssf1.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items").SetCellData 2,3,"102" @@ script infofile_;_ZIP::ssf2.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items_2").SelectCell 2,6 @@ script infofile_;_ZIP::ssf3.xml_;_
 @@ script infofile_;_ZIP::ssf12.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items_2").SetCellData 2,6,"1" @@ script infofile_;_ZIP::ssf4.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPButton("SAPButton").Click @@ script infofile_;_ZIP::ssf5.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items_2").SelectCell 2,14 @@ script infofile_;_ZIP::ssf6.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items_2").SetCellData 2,14,"1710" @@ script infofile_;_ZIP::ssf7.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items_2").SelectCell 2,6 @@ script infofile_;_ZIP::ssf8.xml_;_
'Browser("Create Standard Order:").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items_2").SetCellData 2,6,"2" @@ script infofile_;_ZIP::ssf9.xml_;_
''			
'
'

 @@ script infofile_;_ZIP::ssf11.xml_;_
'
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").SAPTable("micclass:=SAPTable",index:=2).Highlight
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").SAPTable("logical name:=All Items").SetCellData 2,"Material","102"


 @@ script infofile_;_ZIP::ssf14.xml_;_

'













''Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").SAPButton("logical name:=Save").Highlight
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").SAPButton("logical name:=Save").Click
'
'abc =Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").WebElement("outertext:=Standard Order \d.*").GetROProperty("outertext") 
'
'
'value = split (abc," ")
'MsgBox value (0)
'MsgBox value (2)

'Dim oDesc,iCounter
' 
'Set oDesc = Description.Create
' 
'oDesc("micclass").value = "SAPTable"
' 
''Find all the checkboxes
' 
'Set objChkBox = Browser("name:=Create Standard Order.*").Page("title:=Create Standard Order.*").ChildObjects(oDesc)
''objChkBox(2).highlight 
'objChkBox(2).SetCellData 2,3,"102"
'objChkBox(3).SetCellData 2,5,"1"
'objChkBox(3).SetCellData 2,14,"1710"
''iCounter value has the number of all checkboxes in the web page
' 
'iCounter=objChkBox.Count
'' MsgBox iCounter
'For i = 0 to iCounter-1
'                
'   'Click each checkbox one by one
' 
'   objChkBox(i).highlight
'   
'Next
'
'Browser("Create Standard Order:_2").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items").SelectCell 2,6 @@ script infofile_;_ZIP::ssf15.xml_;_
'Browser("Create Standard Order:_2").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items").SetCellData 2,6,"1" @@ script infofile_;_ZIP::ssf16.xml_;_
'Browser("Create Standard Order:_2").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items").SelectCell 2,14 @@ script infofile_;_ZIP::ssf17.xml_;_
'Browser("Create Standard Order:_2").Page("Create Standard Order:").SAPFrame("Create Standard Order:").SAPTable("All Items").SetCellData 2,14,"1710" @@ script infofile_;_ZIP::ssf18.xml_;_
'Browser("name:=Create Standard Order: Overview").Page("title:=Create Standard Order: Overview").SAPTable("micclass:=SAPTable",index=2).Highlight
