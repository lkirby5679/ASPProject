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
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Admin")%> - <%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Edit_Message")%></td></tr>
	<tr>
		<td colspan="2" >  								
							
			<%
			Dim iMessage, strSQL, strAction, iActiveMessageId
			iActiveMessageId = Request.Querystring("mid")

			strName = SQLEncode(Request.Form("Name"))
			strEmail = SQLEncode(Request.Form("Email"))
			strSubject = SQLEncode(Request.Form("Subject"))
			optnotify = Request.Form("Notify")
			iMessage = SQLEncode(Request.Form("message"))
			strAction = Request.Querystring("action")
				
			if strAction = "" then
				strAction = "display"
			end if



			Select Case strAction
				case "display"
					rsAdmin = GetMessages(iActiveMessageId)
					'Display message in form 
					%>
					<form action="admin_message_update.asp?action=done&mid=<%=iActiveMessageId%>" method=post id=form1 name=form1>
					<table BORDER="0" CELLSPACING="0" CELLPADDING="0">
						<tr>
							<td VALIGN="top" class="bolddark"><%=dictLanguage("Name")%>:&nbsp;</td>
							<td><input TYPE="text" NAME="name" MAXLENGTH="50" VALUE="<%=rsAdmin(0)(messages_message_author)%>" class="formstyleLong"></td>
						</tr>
						<tr>
							<td VALIGN="top" class="bolddark"><%=dictLanguage("Email")%>:&nbsp;</td>
							<td><input TYPE="text" NAME="email" MAXLENGTH="50" VALUE="<%=rsAdmin(0)(messages_message_author_email)%>" class="formstyleLong">&nbsp;(<%=dictLanguage("Optional")%>)</td>
						</tr>
						<tr>
							<td VALIGN="top" class="bolddark"><%=dictLanguage("Subject")%>:&nbsp;</td>
							<td><input TYPE="text" NAME="subject" MAXLENGTH="50" VALUE="<%=rsAdmin(0)(messages_message_subject)%>" class="formstyleLong"></td>
						</tr>
						<tr>
							<td VALIGN="top" class="bolddark"><%=dictLanguage("Message")%>:&nbsp;</td>
							<td><textarea ROWS="10" NAME="message" WRAP="virtual" class="formstyleLong"><%=rsAdmin(0)(messages_message_body)%></textarea></td>
						</tr>
						<% If SEND_EMAIL Then %>
						<tr>
							<td VALIGN="top" ALIGN="right">
								<input TYPE="checkbox" NAME="notify" VALUE="<%=rsAdmin(0)(messages_message_author_notify)%>">
							</td>
							<td COLSPAN="2">E-mail me when someone posts a new message in this thread.</td>
						</tr>
						<% End If %>
						<tr>
							<td COLSPAN="2" ALIGN="center">
							<input TYPE="reset" VALUE="Reset Form" id="reset1" name="reset1" class="formButton">&nbsp;&nbsp;
							<input TYPE="submit" VALUE="Update Message" id="submit1" name="submit1" class="formButton"></td>
						</tr>
					</table>
					</form>		
				<%
				Case "done"
					UpdateMessage iActiveMessageId, strName, strEmail, strSubject, optnotify, iMessage 
					%>
					<%=dictLanguage("Message_Updated")%><br><br>
					<a href='admin_message_display.asp?mid=<%=iActiveMessageId%>'><%=dictLanguage("Back_To_Message")%></a> 
			<%
			End Select

			%>


	</TD>		
	<TR>		
</table>

<!--#include file="../../includes/main_page_close.asp"-->