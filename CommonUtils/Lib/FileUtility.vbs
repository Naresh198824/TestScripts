'*********************************************************************************************************************
'**  Project            : 
'**  File Name          : FileUtility.vbs
'**  Version            : 
'**  Created on         : 
'**  Updated on         : 
'**  Description        : Common Methods for executing the functionality in the scripts across Business Units
'**  Copyright          : Arsin Corporation.
'**  Author             :       
'*********************************************************************************************************************
	
Function ReadFileData (sFileName)

		' @HELP
		' @group	: FileUtility	
		' @method	: ReadFileData (sFileName)
		' @returns	: None
		' @parameter: sFileName	: Name of the File to read data
		' @notes	: To Read data from the Specified File
		' @END
		
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim fso, f
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(sFileName, ForReading)
	ReadFileData = f.ReadLine    

	Set fso = Nothing
	Set f = Nothing
	
End Function
'---------------------------------------------------------------------------------------------------------
Function JoinStrings(aFileData)
	
	' @HELP
	' @group	: FileUtility
	' @funtion	: JoinStrings(aFileData)
	' @returns	: Nothing
	' @parameter: aFileData	: Array of data from file
	' @notes	: This function is used to join two strings from a text file.
	' @END
	
	Dim aFileLine(0)
	aFileLine(0) = Join(aFileData,"")
	JoinStrings = aFileLine (0)
	
End Function
'---------------------------------------------------------------------------------------------------------
Function MoveFile (sFileName)

		' @HELP
		' @group	: Files
		' @funtion	: MoveFile (sFileName)
		' @returns	: Nothing
		' @parameter: sFileName: Path of the file to be moved.
		' @notes	: This function moves the file to desired location.
		' @END
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(sFileName, ForAppending, True)
		f.WriteLine ("")
		f.Close
		
		fso.MoveFile sFileName , EXP_DATA_MOVE & Environment.Value("TestPlan") & "_" & Environment.Value("IterationNo") - 1 &".txt"
		
		Set fso = Nothing
		Set f = Nothing
   
End Function
'---------------------------------------------------------------------------------------------------------
Function MoveFileToDestination(sFileName,sPath)

		' @HELP
		' @group	: Files
		' @funtion	: MoveFile (sFileName)
		' @returns	: Nothing
		' @parameter: sFileName	: Filename to be moved.
		' @parameter: sPath		: Destination Location to move the file
		' @notes	: This function moves the file to a specified Location.
		' @END
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(sFileName, ForAppending, True)
		f.WriteLine ("")
		f.Close
		
		fso.MoveFile sFileName, sPath
		
		Set fso = Nothing
		Set f = Nothing

End Function
'---------------------------------------------------------------------------------------------------------
'Writes to any file, file name and values
Function WriteLineToFile (sFileName, sLine)

		' @HELP
		' @group	: FileUtility	
		' @method	: WriteLineToFile (sFileName, sLine)
		' @returns	: None
		' @parameter: sFileName: Name of the File to write data 
		' @parameter: sLine: Text to write into file
		' @notes	:	 	
		' @END
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		Dim fso, f
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(sFileName, ForAppending, True)
		f.WriteLine (sLine)
		f.Close
		
		Set fso = Nothing
		Set f = Nothing
   
End Function
'------------------------------------------------------------------------------------------------------------------------
'Read the text files from different Locations,Compares files  and stores the results in another folder with specified file name, need to pass text file names and result file name
Function ReadAndCompareFilesData (sFile1,sFile2,sResultFile)

		' @HELP
		' @group	: FileUtility	
		' @method	: ReadAndCompareFilesData
		' @returns	: None
		' @parameter: sFile1: First  File Name
		' @parameter: sFile2: Second File Name
		' @parameter: sResultFile: Result File Name to store Result
		' @notes	: To read the Files and compare them 
		' @END
		
		dim errorNb
		errorNb = 0
		
		Const ForReading = 1, ForWriting = 2, BinaryCompare = 0
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set sFile1 = fso.OpenTextFile(sFile1, ForReading)
		Set sFile2 = fso.OpenTextFile(sFile2, ForReading)
		
		Do While ((sFile1.AtEndOfStream <> True) OR (sFile2.AtEndOfStream <> True))
			sData1 = sFile1.ReadLine
			sData2 = sFile2.ReadLine
			comp = strcomp( sData1, sData2, BinaryCompare)
		
			If (comp <> 0) Then
				errorNb = errorNb + 1
				WriteLineToFile sResultFile,sData1
				WriteLineToFile sResultFile,sData2
				WriteLineToFile sResultFile," "
				Reporter.ReportEvent micFail,"Text Comparision Results", "Expected Value :"& "  "& sData1 &" ; "&"Actual Values :"&sData2
			End If
		
		Loop
		
		If errorNb <> 0 Then
			Reporter.ReportEvent micFail, "Text file comparision", errorNb & " Errors found"
		Else 
			Reporter.ReportEvent micPass, "Text file comparision", "No errors found"
		End If
		
		Set fso = Nothing
		Set f = Nothing

End Function
'------------------------------------------------------------------------------------------------------------------------
Sub DeleteAFile(sFilePath)

		' @HELP
		' @group	: FileUtility	
		' @method	: DeleteAFile(filespec)
		' @returns	: None
		' @parameter: sFilePath	: Location of File Name
		' @notes	: To delete a file in the specified location
		' @END
		
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		If Lcase(fso.FileExists(filespec))="true" Then
		   fso.DeleteFile(filespec)
		End If	
		
End Sub
'------------------------------------------------------------------------------------------------------------------------
