<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
<!--#include file="../includes/forum_common.asp"-->

<!-- forum code based on Ti's Portal forum sample.  http://www.transworldportal.com -->

<table width="600" border="0" cellpadding="3" cellspacing="3" align="center">
	<tr>
		<td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Display_Message")%></td>
    </tr>
	<tr>
		<td colspan="2" >  								

								
			<%
			'Response.Write "<BR>su_id -"&session("su_id")
			'Response.Write "<BR>su_section-"&session("su_section")&"<BR><BR>"


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

				objMessageRS = GetMessage(iActiveMessageId)

				If isarray(objMessageRS) Then
					iActiveForumId = objMessageRS(0)(Messages_forum_id)
					%>
					<BR>
					<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0>
						<TR>
							<TD VALIGN="top" class="bolddark"><%=dictLanguage("Author")%>:&nbsp;</TD>
							<TD><%=objMessageRS(0)(Messages_message_author)%></TD>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark"><%=dictLanguage("Email")%>:&nbsp;</TD>
							<% If IsNull(objMessageRS(0)(Messages_message_author_email)) Then %>
							<TD><I>not available</I></TD>
							<% Else %>
							<TD><A HREF="mailto:<%=objMessageRS(0)(Messages_message_author_email) %>"><%=objMessageRS(0)(Messages_message_author_email)%></A></TD>
							<% End If %>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark"><%=dictLanguage("Date")%>:&nbsp;</TD>
							<TD><%=objMessageRS(0)(Messages_message_timestamp)%></TD>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark"><%=dictLanguage("Subject")%>:&nbsp;</TD>
							<TD><%=Lineify(objMessageRS(0)(Messages_message_subject)) %></TD>
						</TR>
						<TR>
							<TD VALIGN="top" class="bolddark"><%=dictLanguage("Message")%>:&nbsp;</TD>
							<TD><%=Lineify(objMessageRS(0)(Messages_message_body)) %></TD>
						</TR>
					</TABLE>
					<BR>
					<A HREF="post_message.asp?fid=<%=iActiveForumId %>&tid=<%=objMessageRS(0)(Messages_thread_id) %>&pid=<%= objMessageRS(0)(Messages_message_id) %>&level=<%=objMessageRS(0)(Messages_thread_level) + 1 %>&subject=<%= Server.URLEncode(objMessageRS(0)(Messages_message_subject)) %>"><%=dictLanguage("Post_A_Reply")%></A><BR>
					<A HREF="default.asp?fid=<%= iActiveForumId %>"><%=dictLanguage("Back_To_The_Folder")%></A><BR>
					<BR>
					<HR>
					
					<BR>
					<%'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
					objMiscRS = GetForums(iActiveForumId)
					If isarray(objMiscRS) Then
						'Sub ShowForumLine(iId, sFolderStatus, sName, sDescription, iMessageCount, iPendingCount, blnAdmin)
						ShowForumLine objMiscRS(0)(forums_forum_id), "open", _
						  objMiscRS(0)(forums_forum_name), _
						  objMiscRS(0)(forums_forum_description), 0 , 0, false			  
					End If
					'Response.write objMessageRS(0)(message_thread_id)
					objMiscRS = GetThreads(objMessageRS(0)(message_thread_id))
					ShowChildren 0, 0, 0, iActiveMessageId, objMiscRS, false
				Else
					WriteLine dictLanguage("Unable_To_Locate_Msg") & "<BR>"
				End If
				'objMessageRS.Close
				'Set objMessageRS = Nothing
			%>						
	</TD>		
	</TR>	
	<tr><td>&nbsp;</td></tr>
	<tr><td align="right"><small><%=dictLanguage("Forum_Code_Based")%> <a href="http://www.transworldportal.com/" target="_blank"><small>Ti Portal</small></a>'s forum.</small></td></tr>		
</table>

<!--#include file="../includes/main_page_close.asp"-->