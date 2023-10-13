'*********************************************************************************************************************
'**  Project            : 
'**  File Name          : WebControls.vbs
'**  Version            : 2.0
'**  Created on         : 
'**  Updated on         : 
'**  Description        : Common Methods for executing the functionality in the scripts across Business Units
'**  Copyright          : Arsin Corporation.
'**  Author             :       
'*********************************************************************************************************************

'-----------Note------------------------------------------------------------------------------------------------------
'ExecuteFile "D:\ExportResults.Vbs"
'dim objExport
'Set objExport=new clsExportResults
'objExport.ExportResults "C:\Results\results.xml", "C:\Results\Results.xls"
'---------------------------------------------------------------------------------------------------------------------

Class clsExportResults

Public Sub CaptureResultPath 
	
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oResPath = oFSO.OpenTextFile( RESULTS_DIR & "ResultPaths.rst",ForReading,True)
		
		If Not oFSO.FileExists(RESULTS_DIR & "ResultPaths.rst") Then
			oFSO.CreateTextFile RESULTS_DIR & "ResultPaths.rst",True
		End If
		sTestName	 = Environment.Value("TestName")
		
		Do While Not oResPath.AtEndOfStream
			sExtTestName = oResPath.ReadLine
			If InStr(Trim(sExtTestName),"Test Case Name") = 1 Then
				If sTestName=Trim(Split(sExtTestName,":")(1)) Then
					Exit Sub
				End If
			End If
		Loop
		
		Set oResPath = Nothing
		Set oResPath = oFSO.OpenTextFile (RESULTS_DIR & "ResultPaths.rst",ForAppending,True)
		
		oResPath.WriteBlankLines 1
		oResPath.WriteLine "Test Case Name: " & sTestName
		oResPath.WriteLine Environment.Value("ResultDir")
		
		Set oFSO = Nothing
		Set oResPath = Nothing
		
End Sub
'--------------------------------------------------------------------------------------------------------------------
Public Sub GenrateResults
	
		Dim iSNo 
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oResPath = oFSO.OpenTextFile( RESULTS_DIR & "ResultPaths.rst",ForReading,True)
		WriteTitleToXL EXCELRESULTS
		iSNo = 1
		
		While Not oResPath.AtEndOfStream
			sTestName = oResPath.ReadLine
			If InStr(Trim(sTestName),"Test Case Name") = 1 Then
				sTestName=Trim(Split(sTestName,":")(1))
				sResultPath = oResPath.ReadLine & "\Report\Results.xml"
				ExportResults sResultPath,EXCELRESULTS, iSNo
				iSNo = iSNo + 1
				'ExportResults "C:\VMWare_AutoSuite\VMWare\Scripts\StartResults\Res11\Report\Results.xml",EXCELRESULTS
			End If
		Wend
		
		oResPath.Close
		
		If  oFSO.FileExists(RESULTS_DIR & "ResultPaths.rst") Then
			'MsgBox "Delete"
			oFSO.DeleteFile RESULTS_DIR & "ResultPaths.rst",True
		End If
		
		Set oFSO = Nothing
		Set oResPath = Nothing
End Sub
'--------------------------------------------------------------------------------------------------------------------
Public Sub WriteToXL(ExcelFilePath,TestCaseResult, FailedSteps, iSNo)

	    Dim xl
	    Set xl = CreateObject("Excel.Application")
	    Dim xlsheet
	    Dim xlwbook
	    Dim startRow
	    Dim IsEmpty

	    Set xlwbook = xl.Workbooks.open(ExcelFilePath)
	    Set xlsheet = xlwbook.Sheets.Item(1)
	    
	    'Find for starting row
	    For i = 1 To 32687
			IsEmpty = True
			For j = 1 To 14
		    	If xlsheet.Cells(i, j) <> "" Then
				IsEmpty = False
				Exit For
		    	End If
			Next
			If IsEmpty Then
			    startRow = i
			    Exit For
			End If
	    Next	    
		
	    If Not IsEmpty Then
			Reporter.ReportEvent micFail,"Result Reporting", "No of rows exceeded in result excel file."
			ExitTest
	    End If
	    
	    
	    'TestCaseResult.open
	    'Writing into excelfile
	    'Test Case Level
	    TestCaseResult.movefirst
	    xlsheet.Cells(startRow, 1) = iSNo 'TestCaseResult("TestCaseName")
	    xlsheet.Cells(startRow, 2) = TestCaseResult("TestCaseName")
	    xlsheet.Cells(startRow, 3) = TestCaseResult("TestResult")
	    xlsheet.Cells(startRow, 4) = TestCaseResult("sTime")
	    xlsheet.Cells(startRow, 5) = TestCaseResult("eTime")
	    xlsheet.Cells(startRow, 6) = CInt(TestCaseResult("noPassed"))
	    xlsheet.Cells(startRow, 7) = CInt(TestCaseResult("noFailed"))
	    xlsheet.Cells(startRow, 8) = CInt(TestCaseResult("noWarnings"))
	    xlsheet.Cells(startRow, 9) = TestCaseResult("Stopped")

	    'Step Level result
	    i = 0
	    
	    If FailedSteps.state = 1 Then
	       FailedSteps.MoveFirst
		   While Not FailedSteps.EOF 
		   		i = i + 1
			    xlsheet.Cells(startRow + i, 10) = FailedSteps("FailedAction")
			    xlsheet.Cells(startRow + i, 11) = FailedSteps("Obj")
			    xlsheet.Cells(startRow + i, 12) = FailedSteps("Details")
			    xlsheet.Cells(startRow + i, 13) = FailedSteps("Time")
			    xlsheet.Cells(startRow + i, 14) = FailedSteps("Disp")
			    FailedSteps.movenext
		   Wend
		End If
	    
	    xlwbook.save

	    'Don't forget to do this or you'll not be able to open'book1.xls again, untill you restart you pc.
	    xl.ActiveWorkbook.Close False, ExcelFilePath
	    xl.Quit
	    
	    Set xlwbook = Nothing
	    Set xl = Nothing
	    
End Sub
'--------------------------------------------------------------------------------------------------------------------
Public Sub ExportResults(ResultFilePath, ExcelFilePath , iSNo)
	
	' @HELP
	' @class	: clsExportResults
	' @method	: ExportResults(ResultFilePath, ExcelFilePath)
	' @returns	: None 
	' @parameter: ResultFilePath     			: Result xml file path.
	' @parameter: ExcelFilePath				: Destination Excel file path.  
	' @notes	: This method exports the results from result xml file to Excel file.
	' @END
	
	  Dim FailedSteps
	  Dim TestCaseResult
	  Dim objDoc
	  Set objDoc = CreateObject("Microsoft.XMLDOM.1.0")
	  Dim objNodeList
	  Dim objNode
	  Dim objStepNodeList
	  Dim objStepNode
	  Dim objEndNodeList
	  Dim objEndNode
	  
	  Dim tFailedAction
	  Dim tObj
	  Dim tDetails
	  Dim tTime
	  Dim tDisp

	  Set TestCaseResult = CreateObject("ADODB.RecordSet")
	    'Specify client-side cursors
	    TestCaseResult.CURSORLOCATION = 3
	    'Add  fields
	    TestCaseResult.fields.APPEND "TestCaseName", 8  ', 40, ADFLDISNULLABLE)
	    TestCaseResult.fields.APPEND "TestResult", 8
	    TestCaseResult.fields.APPEND "sTime", 8
	    TestCaseResult.fields.APPEND "eTime", 8
	    TestCaseResult.fields.APPEND "noPassed", 8
	    TestCaseResult.fields.APPEND "noFailed", 8
	    TestCaseResult.fields.APPEND "noWarnings", 8
	    TestCaseResult.fields.APPEND "Stopped", 8
	  
	  Set FailedSteps = CreateObject("ADODB.RecordSet")
	    'Specify client-side cursors
	    TestCaseResult.CURSORLOCATION = 3
	    'Add  fields
	    FailedSteps.fields.APPEND "FailedAction", 8
	    FailedSteps.fields.APPEND "Obj", 8
	    FailedSteps.fields.APPEND "Details", 8
	    FailedSteps.fields.APPEND "Time", 8
	    FailedSteps.fields.APPEND "Disp", 8
	    
        If objDoc.Load(ResultFilePath) Then
	  		If Not objDoc Is Nothing Then
	    		    
			Set objNodeList = objDoc.selectNodes("Report/Doc") 'Test Case Level
			Set objNode = objNodeList(0)
			TestCaseResult.open
			TestCaseResult.AddNew
			TestCaseResult("TestCaseName") = objNode.childNodes(0).text
			Set objNode = objDoc.selectSingleNode("Report/Doc/NodeArgs")
			TestCaseResult("TestResult") = objNode.Attributes(3).text
			Set objNode = objDoc.selectSingleNode("Report/Doc/Summary")
			TestCaseResult("sTime") = objNode.Attributes(0).Text
			TestCaseResult("eTime") = objNode.Attributes(1).Text
			TestCaseResult("noPassed") = objNode.Attributes(2).Text
			TestCaseResult("noFailed") = objNode.Attributes(3).Text
			TestCaseResult("noWarnings") = objNode.Attributes(4).Text
			TestCaseResult("Stopped") = objNode.Attributes(5).text
			Set objNodeList = objDoc.selectNodes("Report/Doc/Action")
			'ReDim Preserve FailedSteps(0)
			
			If Not objNodeList Is Nothing Then
		
			'Loop though each node in the node list
		    For Each objNode In objNodeList              'Action Level
			tFailedAction = objNode.childNodes(0).Text
			Set objStepNodeList = objNode.childNodes
			For Each objStepNode In objStepNodeList     'Step Level
			    Set objEndNodeList = objStepNode.childNodes
			    For Each objEndNode In objEndNodeList     'Step Detail Level
			    If objEndNode.nodeName = "Obj" Then
				    tObj = objEndNode.Text
				ElseIf objEndNode.nodeName = "Details" Then
				    tDetails = objEndNode.Text
				ElseIf objEndNode.nodeName = "Time" Then
				    tTime = objEndNode.Text
				ElseIf objEndNode.nodeName = "NodeArgs" Then
				    tDisp = objEndNode.childNodes(0).text
				    
				    If  objEndNode.Attributes(3).Text = "Failed" Then
				     
				    	'(objEndNode.Attributes(0).Text = "User" Or objEndNode.Attributes(0).Text = "Replay") And
				     If FailedSteps.state = 0 Then
				     	
				     	FailedSteps.Open
				     End If
				    
				     FailedSteps.AddNew
				     FailedSteps("FailedAction") = tFailedAction
				     FailedSteps("Obj") = tObj
				     FailedSteps("Details") = tDetails
				     FailedSteps("Time") = tTime
				     FailedSteps("Disp") = tDisp
				    End If
				End If
			    Next
			Next
		    Next
		End If
		
		WriteToXL ExcelFilePath, TestCaseResult, FailedSteps , iSNo
		
	  End If
	  End If
	  
	  Set FailedSteps	= Nothing
	  Set TestCaseResult = Nothing
	  
End Sub
'--------------------------------------------------------------------------------------------------------------------	
Public Sub WriteTitleToXL(ExcelFilePath)

	    Dim xl
	    Set xl = CreateObject("Excel.Application")
	    Dim xlsheet
	    Dim xlwbook


	    Dim startRow
	    Dim IsEmpty

	    Set xlwbook = xl.Workbooks.open(ExcelFilePath)
	    Set xlsheet = xlwbook.Sheets.Item(1)
	    'Find for starting row
	    For i = 1 To 32687
		IsEmpty = True
		For j = 1 To 14
		    If xlsheet.Cells(i, j) <> "" Then
			IsEmpty = False
			Exit For
		    End If
		Next
		If IsEmpty Then
		    startRow = i
		    Exit For
		End If
	    Next
	    
		
	    If Not IsEmpty Then
			Reporter.ReportEvent micFail,"Result Reporting", "No of rows exceeded in result excel file."
			ExitTest
	    End If
	    
	    xlsheet.Cells(startRow, 1) = "VMWare Store Application Regression Results Executed On  " & Now
	      
	    xlwbook.save

	    'don't forget to do this or you'll not be able to open'book1.xls again, untill you restart you pc.
	    xl.ActiveWorkbook.Close False, ExcelFilePath
	    xl.Quit
	    
	    Set xlwbook = Nothing
	    Set xl = Nothing
	    
End Sub
'--------------------------------------------------------------------------------------------------------------------
End Class
