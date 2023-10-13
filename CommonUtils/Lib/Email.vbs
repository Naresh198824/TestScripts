'*********************************************************************************************************************
'**  Project            : 
'**  File Name          : Email.vbs
'**  Version            : 
'**  Created on         : 
'**  Updated on         : 
'**  Description        : Common Methods to interact with emails(verify/send/receive) emails
'**  Copyright          : Arsin Corporation.
'**  Author             :       
'*********************************************************************************************************************
'Dim Email
'Set Email = New clsEmail
'Email.VerifyEmailReceivedSubject "customer_service@vmware.com","VMware Order Confirmation - 1240674" 

Public Const olFolderCalendar  	=	9 
Public Const olFolderContacts  =	10 
Public Const olFolderDeletedItems = 3 
Public Const olFolderDrafts = 16 
Public Const olFolderInbox = 6 
Public Const olFolderJournal = 11 
Public Const olFolderJunk = 23 
Public Const olFolderNotes = 12 
Public Const olFolderOutbox =  4 
Public Const olFolderSentMail =  5 
Public Const olFolderTasks = 13 
Public Const olPublicFoldersAllPublicFolders = 18 
Public Const olFolderConflicts = 19 
Public Const olFolderLocalFailures = 21 
Public Const olFolderServerFailures = 22 
Public Const olFolderSyncIssues = 20
	
Class clsEmail

Function  SendMail (sEmailID, sFileName)
	
		' @HELP
		' @class	: clsEmail
		' @method	: SendMail (sEmailID, sFileName)
		' @returns	: None
		' @parameter: sEmailID: Email-id Of Receiver
		' @parameter: sFileName: File path of attachment
		' @notes	: This function sends mail with attachment.
		' @END
		
		Set out = CreateObject("Outlook.Application") 
		Set mapi = out.GetNameSpace("MAPI") 
		Set email = out.CreateItem(0) 
		email.Recipients.Add(sEmailID) 
		email.Subject = "Test Results" 
		email.Body = "Test results from today's run" 
		Set oAttachment = email.Attachments.Add(sFileName) 
		
		email.Send 

		Set oAttachment = Nothing
		Set outlook = Nothing 
		Set mapi = Nothing 
		Set out = Nothing
End Function
'--------------------------------------------------------------------------------------------------------------------	
Function VerifyEmailReceivedBody(sSenderEmail,sSearchString)
		
		' @HELP
		' @class	: clsEmail
		' @method	: VerifyEmailReceivedBody(sSenderEmail,sSearchString)
		' @returns	: None
		' @parameter: sSenderEmail: Email-id Of Sender
		' @parameter: sSearchString: String value to verify
		' @notes	: This function verifies whether email has been received by the sender based on the Body of the Email 
		' @END
	
		Dim OlApp 'As Outlook.Application
		Dim Inbox 'As Outlook.MAPIFolder
		Dim InboxItems 'As Outlook.Items
		Dim Mailobject 'As Object
	
		Set OlApp = CreateObject("Outlook.Application")
		Set Inbox = OlApp.GetNamespace("Mapi").GetDefaultFolder(olFolderInbox) 
		Set InboxItems = Inbox.Items
		
		For Each Mailobject In InboxItems
			If Mailobject.SenderName = sSenderEmail And InStr(Trim(Mailobject.Body),sSearchString) > 0 And FormatDateTime(Mailobject.ReceivedTime,2) = FormatDateTime(Date,2) Then
				'MsgBox Mailobject.SenderName
				'MsgBox Mailobject.Body
				VerifyEmailReceivedBody = True
				Reporter.ReportEvent micPass, "VerifyEmailReceivedBody", "'"&sSearchString&"' was matched"
				Exit For
			End If
		
		    '!Subject = Mailobject.Subject
		    '!from = Mailobject.SenderName
		    '!To = Mailobject.To
		    '!Body = Mailobject.Body
		    '!DateSent = Mailobject.SentOn
		    '.Update		
		Next
		
		Set OlApp = Nothing
		Set Inbox = Nothing
		Set InboxItems = Nothing
		Set Mailobject = Nothing
	
End Function
'--------------------------------------------------------------------------------------------------------------------   
Function VerifyEmailReceivedSubject(sSenderEmail,sSearchString)

		' @HELP
		' @class	: clsEmail
		' @method	: VerifyEmailReceivedSubject(sSenderEmail,sSearchString)
		' @returns	: None
		' @parameter: sSenderEmail: Email-id Of Sender
		' @parameter: sSearchString: String value to verify
		' @notes	: This function verifies whether email has been received by the sender based on the Subject of the Email 
		' @END
	
		Dim OlApp 'As Outlook.Application
		Dim Inbox 'As Outlook.MAPIFolder
		Dim InboxItems 'As Outlook.Items
		Dim Mailobject 'As Object
	
		Set OlApp = CreateObject("Outlook.Application")
		Set Inbox = OlApp.GetNamespace("Mapi").GetDefaultFolder(olFolderInbox) 
		Set InboxItems = Inbox.Items
		
		For Each Mailobject In InboxItems			
			If Mailobject.SenderName = sSenderEmail And InStr(Trim(Mailobject.Subject),sSearchString) > 0 And FormatDateTime(Mailobject.ReceivedTime,2) = FormatDateTime(Date,2) Then
				'MsgBox Mailobject.SenderName
				'MsgBox Mailobject.Subject
				VerifyEmailReceivedSubject = True
				Reporter.ReportEvent micPass, "VerifyEmailReceivedSubject", "'"&sSearchString&"' was matched"
				Exit For
			End If
		
		    '!Subject = Mailobject.Subject
		    '!from = Mailobject.SenderName
		    '!To = Mailobject.To
		    '!Body = Mailobject.Body
		    '!DateSent = Mailobject.SentOn
		    '.Update		
		Next
		
		Set OlApp = Nothing
		Set Inbox = Nothing
		Set InboxItems = Nothing
		Set Mailobject = Nothing
	
End Function
'--------------------------------------------------------------------------------------------------------------------
Function VerifyEmailReceivedSubjectAndBody(sSenderEmail,sSearchSubjectString,sSearchBodyString)
		
		' @HELP
		' @class	: clsEmail
		' @method	: VerifyEmailReceivedSubjectAndBody(sSenderEmail,sSearchSubjectString,sSearchBodyString)
		' @returns	: None
		' @parameter: sSenderEmail: Email-id Of Sender
		' @parameter: sSearchSubjectString: Subject String value to verify
		' @parameter: sSearchBodyString: Body String value to verify
		' @notes	: This function verifies whether email has been received by the sender based on the Subject & Body of the Email 
		' @END
	
		Dim OlApp 'As Outlook.Application
		Dim Inbox 'As Outlook.MAPIFolder
		Dim InboxItems 'As Outlook.Items
		Dim Mailobject 'As Object
				
		Set OlApp = CreateObject("Outlook.Application")
		Set Inbox = OlApp.GetNamespace("Mapi").GetDefaultFolder(olFolderInbox) 
		Set InboxItems = Inbox.Items
		'
		For Each Mailobject In InboxItems			
			If Mailobject.SenderName = sSenderEmail And InStr(Trim(Mailobject.Subject),sSearchSubjectString) > 0 And InStr(Trim(Mailobject.Body),sSearchBodyString) > 0 And FormatDateTime(Mailobject.ReceivedTime,2) = FormatDateTime(Date,2) Then
				VerifyEmailReceivedSubjectAndBody = True
				Reporter.ReportEvent micPass, "VerifyEmailReceivedSubjectAndBody", "Subject:='"&sSearchSubjectString&"' And Body:='"&sSearchBodyString&"' was matched"
				Exit For
			End If		
		Next
		
		Set OlApp = Nothing
		Set Inbox = Nothing
		Set InboxItems = Nothing
		Set Mailobject = Nothing
	
End Function
'--------------------------------------------------------------------------------------------------------------------
End Class

Function MailSmokeReport(Attachment,TeamMerEmailid)
If EMAIL_REPORT_TAG<> "T" Then
	Exit Function
End If
Const EmailFrom = "automationqauser@vmware.com"
Const EmailFromName = "SDP_Automation"
 EmailTo = TeamMerEmailid
Const SMTPServer = "email.vmware.com"
Const SMTPLogon = "automationqauser"
Const SMTPPassword ="@ut0M@ti0nP0rtal123"
Const SMTPSSL = False
Const SMTPPort = 25

Const cdoSendUsingPickup = 1 'Send message using local SMTP service pickup directory.
Const cdoSendUsingPort = 2       'Send the message using SMTP over TCP/IP networking.
Const cdoAnonymous = 0            ' No authentication
Const cdoBasic = 1           ' BASIC clear text authentication
Const cdoNTLM = 2         ' NTLM, Microsoft proprietary authentication

   
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "AutomatedEmailAlert-SDP QA Smoke Dry run report"
objMessage.From = """" & EmailFromName & """ <" & EmailFrom & ">"
objMessage.To = EmailTo
objMessage.AddAttachment Attachment
objMessage.TextBody = "AutomatedEmailAlert-attached QA Smoke Dry run report."
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTPLogon
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTPPassword
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPPort
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SMTPSSL
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
objMessage.Configuration.Fields.Update

' Now send the message!
objMessage.Send
Set objMessage = Nothing

End Function