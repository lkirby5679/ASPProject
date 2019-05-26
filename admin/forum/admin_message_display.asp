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
<table border="0" width="500" cellpadding="3" cellspacing="3" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader">Discussion Forum Admin - Display Message</td></tr>
	<tr>
		<td colspan="2" >  								
										

			<script LANGUAGE="JavaScript">
			<!--
			  function chkDelete(what)
			  {
			     return confirm('<%=dictLanguage("Confirm_Delete_Forum")%>')
			  }
			//-->
			</script>

			<%
				Dim iActiveMessageId
				Dim iActiveForumId
				Dim objMessageRS
				'Dim objMiscRS

				iActiveMessageId = Request.QueryString("mid")
				If IsNumeric(iActiveMessageId) Then
					iActiveMessageId = CInt(iActiveMessageId)
				Else
					iActiveMessageId = 0
				End If

				objMessageRS = GetMessages(iActiveMessageId)

				If isarray(objMessageRS)  Then
					iActiveForumId = objMessageRS(0)(Messages_forum_id)
					%>
					<BR>
					<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>
						<TR>
							<TD VALIGN="top" class="bolddark">Author:&nbsp;</TD>
							<TD><%= objMessageRS(0)(Messages_message_author) %></TD>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark">E-mail:&nbsp;</TD>
							<% If IsNull(objMessageRS(0)(Messages_message_author_email)) Then %>
							<TD><FONT SIZE="-1"><I>not available</I></FONT></TD>
							<% Else %>
							<TD><A HREF="mailto:<%=objMessageRS(0)(Messages_message_author_email)%>"><%= objMessageRS(0)(Messages_message_author_email) %></A></TD>
							<% End If %>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark">Date:&nbsp;</TD>
							<TD><%=objMessageRS(0)(Messages_message_timestamp) %></TD>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark">Subject:&nbsp;</TD>
							<TD><%= Lineify(objMessageRS(0)(Messages_message_subject)) %></TD>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark">Message:&nbsp;</TD>
							<TD><%= Lineify(objMessageRS(0)(Messages_message_body)) %></TD>
						</TR>
					</TABLE>
					<BR>
					<A HREF="admin_message_post.asp?fid=<%= iActiveForumId %>&tid=<%=objMessageRS(0)(Messages_thread_id) %>&pid=<%=objMessageRS(0)(Messages_message_id) %>&level=<%=objMessageRS(intCOunter)(Messages_thread_level) + 1 %>&subject=<%= Server.URLEncode(objMessageRS(intCOunter)(Messages_message_subject)) %>"><%=dictLanguage("Post_A_Reply")%></A><BR>
					<A HREF="admin_forum_display.asp?fid=<%= iActiveForumId %>"><%=dictLanguage("Back_To_The_Folder")%></A><BR>
					<A onClick='return chkDelete()' HREF="admin_message_delete.asp?mid=<%= iActiveMessageId %>"><%=dictLanguage("Delete_Message")%></A><BR>
					<A HREF="admin_message_update.asp?mid=<%= iActiveMessageId %>"><%=dictLanguage("Modify_Message")%></A><BR>
					<BR>
					<HR>
					<BR>
					<%'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
					objMiscRS = GetForums(iActiveForumId)
					If isarray(objMiscRS) Then
						'Sub ShowForumLine(iId, sFolderStatus, sName, sDescription, iMessageCount, iPendingCount, blnAdmin)
						ShowForumLine objMiscRS(0)(forums_forum_id), "open", _
						  objMiscRS(0)(forums_forum_name), _
						  objMiscRS(0)(forums_forum_description), 0 , 0, TRUE			  
					End If
					'Response.write objMessageRS(0)(message_thread_id)
					objMiscRS = GetThreads(objMessageRS(0)(message_thread_id))
					ShowChildren 0, 0, 0, iActiveMessageId, objMiscRS, TRUE
				Else
					WriteLine dictLanguage("Unable_To_Locate_Msg") & "<BR>"
				End If
			%>

	</TD>		
	<TR>		
</table>

<!--#include file="../../includes/main_page_close.asp"-->