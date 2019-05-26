<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->
<!--#include file="../../includes/forum_common.asp"-->
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
<table border="0" cellpadding="3" cellspacing="3" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Admin")%> - <%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Post_Message")%></td></tr>
	<tr>
		<td colspan="2" > 							
			<%	' Message parameters
				Dim iForumId, iThreadId, iThreadParent, iThreadLevel
				Dim sSubject, sMessage, bNotify
				
				Dim sName, sEmail ' User Info from Cookies

				Dim iNewMessageId ' Id of the message we're adding

				Select Case Request.QueryString("action")
					Case "save"
						' Retrieve parameters
						iForumId = Request.Form("forum_id")
						iThreadId = Request.Form("thread_id")
						iThreadParent = Request.Form("thread_parent")
						iThreadLevel = Request.Form("thread_level")
						sName = Request.Form("name")
						sEmail = Request.Form("email")
						sSubject = Request.Form("subject")
						sMessage = Request.Form("message")
						bNotify = Request.Form("notify")

						If bNotify = "yes" Then
							bNotify = True
						Else
							bNotify = False
						End If

						 ' Validate Input
						If InputIsValid("save", iForumId, iThreadId, iThreadParent, iThreadLevel, sName, sSubject, sMessage) Then
							' Insert the New Message
							iNewMessageId = InsertRecord(iForumId, iThreadId, iThreadParent, iThreadLevel, sName, sEmail, bNotify, sSubject, sMessage)
							
							' Show The Thanks Page
							ShowThanks iNewMessageId, iThreadParent, iForumId, sName, sEmail, true
							
							' Send Email Notification
							SendEmailNotification iNewMessageId, iThreadId, sEmail
						Else
							ShowForm iForumId, iThreadId, iThreadParent, iThreadLevel, sName, sEmail, sSubject, sMessage, true
						End If
					Case Else
						' Retrieve Parameters
						iForumId = Request.QueryString("fid")
						iThreadId = Request.QueryString("tid")
						iThreadParent = Request.QueryString("pid")
						iThreadLevel = Request.QueryString("level")
						'sName = Request.Cookies("name")
						sName = Session("userid")
						sEmail = Request.Cookies("email")
						sSubject = Request.QueryString("subject")
						'sMessage = Request.Form("message")
						If Len(sSubject) <> 0 And Left(sSubject, 3) <> "Re:" Then
							If Len(sSubject) > 46 Then ' If Re: won't fit!
								sSubject = "Re: " & Left(sSubject, 43) & "..."
							Else
								sSubject = "Re: " & sSubject
							End If			
						End If

						If InputIsValid("post", iForumId, iThreadId, iThreadParent, iThreadLevel, sName, sSubject, sMessage) Then
							ShowForm iForumId, iThreadId, iThreadParent, iThreadLevel, sName, sEmail, sSubject, sMessage, true
						Else
							' A message should have been displayed by the validation routine so we do nothing!
						End If
				End Select
			%>

	</TD>		
	<TR>		
</table>
<!--#include file="../../includes/main_page_close.asp"-->