<%
if not session("permAdminForum") then
	Response.Redirect("default.asp")
end if
%>
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
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Admin")%> - <%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Add_Forum")%></td></tr>
	<tr>
		<td colspan="2" >  								
								
			<%

			'Find out what to do
			Dim strAction 
			strAction = Request.Querystring("action")
			if strAction = "" then
				strAction = "start"
			end if

			'Variables for new forum record
			Dim iForumName, iForumDescription, iForumStart, iForumGroup, strSQL

			iForumName = SQLEncode(Trim(cStr(Request.Form("forum_name"))))
			iForumDescription = SQLEncode(Trim(cStr(Request.Form("forum_description"))))
			iForumStart = Date()
			iForumGroup = Trim(cStr(Request.Form("forum_grouping")))
			su_section = request("su_section")

			if iForumGroup = "" then
				iForumGroup = "none"
			end if
			select case strAction
			 
				case "start"
					'display form for adding a forum
					%>
					<table>
						<tr><td><%=dictLanguage("forumadmininst_1")%></td></tr>
						<tr>
							<td>
							<form action="admin_forum_add.asp?action=done" method="post" id=form1 name=form1>
							<table>
								<tr><td><%=dictLanguage("Forum_Name")%>:</td><td><input type="text" name="forum_name" size="20" maxlength="50" class="formstyleLong"></input></td></tr>
								<tr><td><%=dictLanguage("Description")%>:</td><td><input type="text" name="forum_description" size="50" maxlength="50" class="formstyleLong"></input></td></tr>
		
								
								<tr>
									<td>
									</td>
									<td>
										<input type="submit" value="Add Forum" id=submit1 name=submit1 class="formButton">
										</form>
									</td>
								</tr>
							</table>
							</td>
						</tr>
					</table>

				<%
				case "done"
					InsertForum iForumName, iForumDescription, iForumStart, 1, iForumGroup
					%>
					<br><br><br>
					<%=dictLanguage("The_Forum")%>&nbsp;<%= iForumName %>&nbsp;<%=dictLanguage("Has_Been_Created")%><br>
					<br><a href='admin.asp'><%=dictLanguage("Back_To_Forums")%></a>
					<%
			End Select

			%>

	</TD>		
	<TR>		
</table>

<!--#include file="../../includes/main_page_close.asp"-->