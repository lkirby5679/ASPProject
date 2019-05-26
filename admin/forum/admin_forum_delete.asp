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
	<tr>
		<td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Admin")%> - <%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Delete_Forum")%></td>
    </tr>
	<tr>
		<td colspan="2" >  								
							

			<%
			Dim strSQL, strAction, iForumId, Item
			strAction = Request.Querystring("action")
			iForumId = Request.Querystring("fid")

			'Depending on what is selected we display info
			select case strAction
				'Ask if they are sure
				case "verify"
					rsForumAdmin = GetForums(iForumId)
					%>
					<br>
					<table align="center">
						<tr>
							<td colspan=2><%=dictLanguage("deleteforum_1")%></td>
						</tr>
						<tr>
							<td align="center" colspan=2>
								<b><%=rsForumAdmin(0)(Forums_Forum_name)%></b>
							</td>
						</tr>
						<tr>
							<td colspan=2><%=dictLanguage("deleteforum_2")%><BR><BR>
							</td>
						</tr>
						<tr>
							<td align="center" width="50%">
								<form action="admin_forum_delete.asp?action=delete&fid=<%=rsForumAdmin(0)(Forums_Forum_ID)%>" method="post" id=form1 name=form1><input type="submit" value="Delete" id=submit1 name=submit1 class="formButton"></input>
								</form>
							</td>
							<td align="center" width="50%">
								<form action="admin_forum_delete.asp?action=cancel&fid=<%=rsForumAdmin(0)(Forums_Forum_ID)%>" method="post" id=form2 name=form2><input type="submit" value="Cancel" id=submit1 name=submit1 class="formButton"></input>
								</form>
							</td>
						</tr>
					</table>
					<%
				case "cancel"
					'If they want to cancel, close everything and send them back to forums
					Response.Redirect("admin.asp")
				case "delete"
				
					DeleteForum iForumId
					%>
					<%=dictLanguage("Forum_Deleted")%><br>
					<br><a href='admin.asp'><%=dictLanguage("Back_To_Forums")%></a>
					<%
			End Select
			%>

							
	</TD>		
	<TR>		
</table>

<!--#include file="../../includes/main_page_close.asp"-->