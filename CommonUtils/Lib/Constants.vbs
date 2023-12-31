' $Filename:		Constants.vbs
' $Description: 	Constants defined for project
' $Copyright: 		Mahesh
'-------------------------------------------------------------------------------------------------------------------------
'Common framework folders
COMMON_REPO_DIR	= sAutomationPath	& "\CommonUtils\Repo\"
'COMMON_LIB_DIR										= sAutomationPath	& "\Common\Lib\"
REM MsgBox "Contsts in"
COMMON_RECOVERY_DIR									= COMMON_LIB_DIR 	& "Recovery\"
COMMON_UTILITIES_DIR								= COMMON_LIB_DIR 	& "Common\Utilities\"

'Project related folders
PROJECT_DIR											= sProjectPath		& "\"
CLASS_DIR 											= sAutomationPath 		& "\BusinessClass\Classes\"
DATA_DIR 											= PROJECT_DIR 		& "Data\"
ERRORS_DIR											= PROJECT_DIR 		& "Errors\"
LIB_DIR												= PROJECT_DIR 		& "Lib\"
LOGS_DIR											= PROJECT_DIR 		& "Logs\"
OR_DIR												= PROJECT_DIR 		& "OR\"
RESULTS_DIR											= PROJECT_DIR 		& "Results\"
TEMP_DIR											= PROJECT_DIR 		& "Temp\"
EXCEL_RESULT_FILE									= RESULTS_DIR 		& "ExecutionStatus.xlsx"
GENERATE_HTML_REPORT								= "TRUE"
MODULE_NAME											= GetModuleName
SR_ATTACHEDFILE 									= DATA_DIR			& "TestAttached.txt"
'PROJECT_MASTER										= PROJECT_DIR 		& "Lib\ProjectMaster.vbs"
'Public Const VERY_SHORT_TIME_OUT					= 3		'Seconds
'Public Const SHORT_TIME_OUT							= 5		'Seconds
'Public Const MEDIUM_TIME_OUT						= 10	'Seconds
'Public Const LONG_TIME_OUT							= 30	'Seconds

'Browser Type
'Public Const SAP_Version							= "SAP"
'Public Const BROWSER_NAME							= "FF"

'Flag for Execution Mode
'Public Const EXEC_MODE								= "QTP"

'To Debug the QTP Scripts 
'BVALUE_DEBUG_MODE                                   = True
'BVALUE_DEBUG_MODE                                   = False

'EMAIL_REPORT_TAG                                   = "F"

'DB_PATH = DATA_DIR & "TestData.accdb"
'DB_PATH = DATA_DIR & "TestData.xlsx"
DB_PATH = DATA_DIR & "TestData.xlsx"
'--------------------------------------------------------------------------------------------------------------------------
'************************************************ Constants to be supplied ********************************************************	 

Set oDataAccess		= new clsDataAccess

'sSQL = "SELECT * FROM Execution_Environment WHERE ExecutionTag='T' "
'sSQL="Select * From [Execution_Environment$] where ExecutionTag='T' "
'aLoginData = oDataAccess.GetSingleRowValuefromAccessDB(sSQL)

'sSQL = "SELECT  * from [Scenarios_CreditControl$] where ExecutionTag='T' order by SNo Asc"
'oData  = oDataAccess.GetSingleRowValuefromAccessDB (sSQL)
sSQL="Select * From [Execution_Environment$] where ExecutionTag='T' "
aLoginData = oDataAccess.GetSingleRowValuefromAccessDB(sSQL)

'sSQL = "SELECT  * from [Execution_Environment$] where ExecutionTag='T' order by SNO Asc"
'oData  = oDataAccess.GetSingleRowValuefromAccessDB (sSQL)
'
'sSQL_GETRECORD = "SELECT  * from [IOData_CustomerData$] where TestDataTableLinkID='Aftercare_Home_Demo'"
'	Set InputData = oDataAccess.ExecSQLStatementWithWhereClass (sSQL_GETRECORD)

If UBound(aLoginData)>=0 Then
	EXECUTION_ENVIRONMENT_0 = aLoginData(0,0)
	EXECUTION_ENVIRONMENT = aLoginData(0,1)
	EXECUTION_ENVIRONMENT_2 = aLoginData(0,2)
	EXECUTION_ENVIRONMENT_3 = aLoginData(0,3)
	EXECUTION_ENVIRONMENT_4 = aLoginData(0,4)

	REM MsgBox EXECUTION_ENVIRONMENT
	REM sVSR_URL = aLoginData(0,2)
	REM MsgBox sVSR_URL
	REM sVSR_UserName = aLoginData(0,3)
	REM MsgBox sVSR_UserName
	REM sVSR_Password = aLoginData(0,4)
	REM MsgBox sVSR_Password
Else
	Reporter.ReportEvent micFail,"Set Execution Environment","Does not found Execution Environment in DB"
	ExitTest
End If

'If UBound(oData)>=0 Then
'	EXECUTION_ENVIRONMENT_0 = oData(0,0)
'	EXECUTION_ENVIRONMENT = oData(0,1)
'	EXECUTION_ENVIRONMENT_2 = oData(0,2)
'	EXECUTION_ENVIRONMENT_3 = oData(0,3)
'
'	REM MsgBox EXECUTION_ENVIRONMENT
'	REM sVSR_URL = aLoginData(0,2)
'	REM MsgBox sVSR_URL
'	REM sVSR_UserName = aLoginData(0,3)
'	REM MsgBox sVSR_UserName
'	REM sVSR_Password = aLoginData(0,4)
'	REM MsgBox sVSR_Password
'	Else
'	Reporter.ReportEvent micFail,"Set Execution Environment","Does not found Execution Environment in DB"
'	ExitTest
'End If

'*****************************************************************************************************************	  
REM 'Tag Name Constants
REM Public Const LINK_TAG								= "Link"
REM Public Const NAME_TAG								= "name"
REM Public Const IMAGE_TAG								= "Image"
REM Public Const ALT_TAG								= "alt"
REM Public Const INNER_TEXT_TAG							= "innertext"
REM Public Const HTML_ID_TAG							= "html id"
REM Public Const HTML_TAG								= "html tag"
REM Public Const INNERTEXT_TAG							= "innertext"
REM Public Const TEXT_TAG								= "text"
REM Public Const URL_TAG								= "url"
REM Public Const HREF_TAG								= "href"
REM Public Const INDEX_TAG								= "index"
REM Public Const WEB_BUTTON_TAG							= "WebButton"
REM Public Const WEBEDIT_NAME							= "WebEdit"
REM Public Const WEBLIST_TAG      						= "WebList"
REM Public Const FILE_NAME_TAG							= "file name"
REM Public Const WEBELEMENT_TAG							= "WebElement"
REM Public Const WEBFILE_TAG							= "WebFile"
REM Public Const WEB_CHECK_BOX_TAG						= "WebCheckBox"
REM Public Const WEB_RADIO_GROUP_TAG    				= "WebRadioGroup"
REM Public Const OK_BUTTON								= "OK"
REM Public Const YES_BUTTON								= "YES"
REM Public Const NO_BUTTON								= "NO"
REM Public Const CANCEL_BUTTON							= "Cancel"
REM Public Const SAY_YES      							= "Y"
REM Public Const SAY_NO       							= "N"
REM Public Const INDEX       							= "index"
REM Public Const PARAGRAPH_TAG							= "P"
REM Public Const VALUE_TAG								= "value"
REM Public Const HEIGHT_TAG								= "height"
REM Public Const SPAN_TAG								= "SPAN"
REM Public Const SELECT_ONE_TAG							= "Select One"
REM Public Const NONE_TAG								= "None"
'Public Const INDEX_VALUE_ZERO    					=  0
'Public Const INDEX_VALUE_ONE    					=  1
'Public Const INDEX_VALUE_TWO    					=  2
'Public Const INDEX_VALUE_THREE	    				=  3			

'-------------------------------------------------------------------------------------------

