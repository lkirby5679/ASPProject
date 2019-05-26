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
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Admin")%> - <%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Update_Forum")%></td></tr>
	<tr>
		<td colspan="2">  								
							

			<%
			Dim iForumName, iForumDescription, iForumID
			dim iForumStart, iForumGroup, strAction, strSQL, Item

			iForumID = trim(cInt(Request.QueryString("fid")))
			if iForumID = 0 then
				iForumID = trim(cInt(Request.Form("forum_id")))	
			end if

			iForumName = SQLEncode(Request.Form("forum_name"))
			iForumDescription = SQLEncode(Request.Form("forum_description"))
			iForumStart = Request.Form("forum_start_date")
			iForumGroup = Request.Form("forum_grouping")
			su_section =  Request.Form("su_section")

			if iForumGroup = "" then
				iForumGroup = "none"
			end if

			strAction = Request.Querystring("action")
			if strAction = "" then
				strAction = "search"
			end if

			Select case strAction
				case "search"
					%>
					<b><%=dictLanguage("updateforum_1")%>:</b>
					<form action="admin_forum_update.asp?action=display" method="post" id=form1 name=form1>
					<input type="text" size="2" name="forum_id" class="formstyleTiny">
					<input type="submit" value="Submit" id=submit1 name=submit1 class="formButton"></input>
					</form>
					<%
				case "display"
					rsAdmin = GetForums(iForumID)
					if isarray(rsadmin) then 
						%>
						
						<form Action='admin_forum_update.asp?action=update' method='post' id=form2 name=form2>
						<input type=hidden name=forum_id value='<%=rsadmin(0)(forums_Forum_id)%>'>
							<table align="center" border="0">
							<tr><td><%=dictLanguage("Forum_Name")%>:  </td><td><input type=text name='forum_name' value='<%=rsadmin(0)(forums_forum_name)%>' size=20 maxlength="50" class="formstyleLong"></td></tr>
							<tr><td><%=dictLanguage("Description")%>: </td><td><input type=text name='forum_description' value='<%=rsadmin(0)(forums_forum_description)%>' maxlength="50" size="50" class="formstyleLong"></td></tr>
							<!--<tr><td>Forum Start Date:	</td><td><input type=text name='forum_start_date' value='<%=rsadmin(0)(forums_forum_start_date)%>'><BR></td></tr>-->
							<tr><td>&nbsp;				              </td><td><input type=hidden name='forum_grouping' value='none'></td></tr>
							<tr>
								<td align="right">
									<input type='submit' value='Update' id='submit' name='submit' class="formButton">
								</td>
								</form>
								<form action='admin_forum_update.asp?action=cancel' method=post id=form3 name=form3>
								<td>												
									<input type='submit' value='Cancel' id='submit' name='submit' class="formButton"></input>
								</td>
								</form>
							</tr>
						</table>
					<%end if
				case "cancel"
						'If they want to cancel, close everything and send them back to forums
						'cnnAdmin.Close
						Response.Redirect("admin.asp")
				case "update"
					UpdateForum iForumID, iForumName, iForumDescription, 1, iForumGroup
					'Let them know that it has been updated and give them option to start over.
					%>
					<%=dictLanguage("Forum_Updated")%><br><br>
					<a href='admin.asp'><%=dictLanguage("Back_To_Forums")%></a>
					<%
			End Select
			%>

	</TD>		
	<TR>		
</table>

<!--#include file="../../includes/main_page_close.asp"-->