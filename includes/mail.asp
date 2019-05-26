<!--
'*******************************************************************************************
' Transworld Interactive Projects - Version 2.7.2
' Written by Professor L. T. Kirby - Copyright 1998-2011 (c) Professor L. T. Kirby, Transworld Interactive. All Rights Reserved.
' This work is licensed under the Creative Commons Attribution-NonCommercial-ShareAlike 2.5 License. 
' To view a copy of this license, visit http://creativecommons.org/licenses/by-nc-sa/2.5/ 
' or send a letter to Creative Commons, 543 Howard Street, 5th Floor, San Francisco, California, 94105, USA.
' All Copyright notices MUST remain in place at ALL times.
'******************************************************************************************** 
-->

<%Sub SendEmail(eToName, eToEmail, eFromName, eFromEmail, eSubject, eBody, eCCName, eCCEmail, _
		eBCCName, eBCCEmail, eAttachedFile, eBoolHTML)

	if gsMailComponent = "CDONTS" then
		
		Set Mailer = Server.CreateObject("CDONTS.NewMail")
		Mailer.To = eToEmail
		Mailer.From = eFromEmail
		Mailer.Subject = eSubject
		Mailer.Body = eBody
		if eAttachedFile <> "" then
			Mailer.AttachFile = eAttachedFile
		end if
		if eCCEmail <> "" then
			Mailer.Cc = eCCEmail
		end if
		if eBCCEmail <> "" then
			Mailer.Bcc = eBCCEmail
		end if
		if eBoolHTML then
			Mailer.BodyFormat = 0
			Mailer.MailFormat = 0
		end if
		Mailer.Send
		Set Mailer = Nothing

	elseif gsMailComponent = "CDO" then
		const cdoOutlookExvbsss = 2
		const cdoIIS = 1
 
		'comment out the following section if you do not need to set a specific smtp 
		'server to use for sending email 
		If IsEmpty(Conf) Then
			Const cdoSendUsingPort = 2
			Set Conf = CreateObject("CDO.Configuration")
			With Conf.Fields
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = gsMailHost
				.Update
			End With
			Conf.Load cdoIIS
		End If
  
		Set Mailer = CreateObject("CDO.Message")
		With Mailer
			Set .Configuration = Conf   
			'if you are not going to use CDO configuration section above to set smtp 
			'server then comment out the above line (set .congfiguration = conf)
			'and remove the comment from the below line
    		'.Configuration.Load cdoIIS
			.To = """" & eToName & """ <" & eToEmail & ">"
			.Subject = eSubject
			if eBoolHTML then
				.HTMLBody = eBody
			else
				.TextBody = eBody
			end if
			if len(eCCEmail) > 0 then
				.CC = """" & eCCName & """ <" & eCCEmail & ">"
			end if
			if len(eBCCEmail) > 0 then
				.BCC = """" & eBCCName & """ <" & eBCCEmail & ">"
			end if
			If len(eFromEmail) > 0 Then
				.From = """" & eFromName & """ <" & eFromEmail & ">"
			end if
			if len(eAttachedFile) > 0 then
				.AddAttachment eAttachedFile			
			end if
			.Send
		End With
		set Mailer = nothing

	elseif gsMailComponent = "ASPXP" then
	
		Set mail = Server.CreateObject("ASPXP.Mail") 
		mail.Host = gsMailHost
		mail.Port = 25 
		mail.Subject = eSubject
		mail.FromName = eFromName
		mail.FromAddress = eFromEmail
		'mail.Importance = 2 
		mail.ToList.Add eToEmail, EToName		
		if eBoolHTML then
			mail.HTMLBody = eBody 
		else
			mail.PlainBody = eBody
		end if
		if eCCEmail <> "" then
			mail.CCList.Add eCCEmail, eCCName
		end if
		if eBCCEmail <> "" then
			mail.BCCList.Add eBCCEmail, eBCCName
		end if
		if eAttachedFile <> "" then
			mail.Attachments.Add eAttachedFile
		end if
		
		on error resume next	
		mail.Send

		set mail = nothing

	elseif gsMailComponent = "PersitsASPMail" then
		'this is just the framework, it needs to be setup for attachment, cc, bcc, etc
		
		set mail 	= server.CreateObject("Persits.MailSender")
		mail.Host 	= gsMailHost
		mail.From 	= eFromEmail
		mail.FromName 	= eFromName
		mail.AddAddress  eToEmail
		mail.AddReplyTo eFromEmail
		mail.Subject 	= eSubject
		mail.Body 	= eBody

		on error resume next
		mail.Send
		
	elseif gsMailComponent = "Jmail" then
		'this is just the framework, it needs to be setup for attachment, cc, bcc, etc
		
		set mail 	= server.CreateOBject( "JMail.Message" ) 
		mail.Logging 	= true
		mail.silent 	= true
		mail.From 	= eFromEmail
		mail.FromName 	= eFromName
		mail.Subject	= eSubject
		mail.AddRecipient eToEmail
		mail.Body 	= eBody
 
		on error resume next
 		mail.send(gsMailHost) 	
		
	elseif gsMailComponent = "ServerObjectsASPMail" then
		'this is just the framework, it needs to be setup for attachment, cc, bcc, etc
		
		set mail = Server.CreateObject("SMTPsvg.mail")
		mail.RemoteHost  = gsMailHost
		mail.FromName    = eFromName
		mail.FromAddress = eFromEmail
		mail.AddRecipient eToName, eToEmail
		mail.Subject     = eSubject
		mail.BodyText    = eBody
		if eBoolHTML then
			mail.ContentType = "text/html"
		end if

		on error resume next
		mail.SendMail 

	elseif gsMailComponent = "MYASPMAILComponent" then ' some other component	
		'then do something else

	end if

End Sub%>
