Function UpdateValuesIntoAccessDBB(sSQL)
	Dim Retlng
	cnnAccess.ExecuteSQL sSQL
	'MsgBox "hai"
	If Retlng >= 1 Then
		MsgBox "Executed Query : "&sSQL
	Else
		MsgBox "Query Failed : "&sSQL
	End If
End Function
'****************************************************************************************************************************************************************************************
Dim oConn
Public Function DBConnect(sDB_UID,sDB_PWD,sDB_Host,sDatabase)
	Set oConn = CreateObject("ADODB.Connection")
	Set RecSet = CreateObject("ADODB.RecordSet")
	Srvname="DRIVER={Oracle in instantclient10_2};DBQ="&sDB_Host&":1521/"&sDatabase&";uid="&sDB_UID&";pwd="&sDB_PWD&";"
	oConn.open Srvname
	If oConn.State = 1 Then
	  	ReportWriter "Pass","Data Connection is Passed","DB Connected Successfully",0
	Else
		ReportWriter "Fail","Data Connection is Failed","DB Connected Not Successfully",0
	End If
End Function
'****************************************************************************************************************************************************************************************
Function DBDisconnect()
	oConn.Close
   	If oConn.State = 0 Then
  		ReportWriter "Pass","Data DisConnected Closed is Passed","DB DisConnected Successfully",0
 	Else
  		ReportWriter "Fail","Data DisConnected Closed is Failed","DB DisConnected Successfully",0
    End If
End Function
'****************************************************************************************************************************************************************************************
Function ExecuteSQL(sSQL)
	Dim aDBData(25)
	'On Error Resume Next
	Set rs = CreateObject("ADODB.recordset")
	rs.Open sSQL, oConn
	If Not rs.EOF Then
	 	k = 0
		rs.MoveFirst
		Do
			If k > 0 Then
				rs.MoveNext
			End If
			aDBData(k) = rs.fields(k)	
		    k = k + 1
		Loop Until k >= rs.fields.Count	
		ExecuteSQL = aDBData
		Else
		ExecuteSQL=""
	End If
End Function
'*********