<%@ LANGUAGE="VBSCRIPT" %>
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
<!--#include file="../includes/main_page_header.asp"-->


<%
active = request.querystring("active")
if active = "" or active="Active" then 
	active="Active" 
	activeView = dictLanguage("Active")
elseif active="All" then
	active = "All"
	activeView = dictLanguage("All")
end if

if active = "Active" then
	sql = sql_GetActiveProjectsWithClients()
else
	sql = sql_GetAllProjectsWithClients()
end if
'response.write sql
Call RunSQL(sql, rsViewProjects)
%>

<!--#include file="../includes/main_page_open.asp"-->

<table width="100%" border=0 cellpadding=2 cellspacing=2 style="border-style: solid; border-width: 2; border-color: <%=gsColorBlack%>">
	<tr>
		<td valign=top colspan=6>
			<b class="header"><%=dictLanguage("View")%>&nbsp;<%=activeView%>&nbsp;<%=dictLanguage("Projects")%><br>
			<%if active="Active" then%>
				<a href="project-view.asp?active=All" class="small"><%=dictLanguage("View_All_Projects")%></a>
			<%else%>
				<a href="project-view.asp?active=Active" class="small"><%=dictLanguage("View_Active_Projects")%></a>
			<%end if%>
		</td>
	</tr>
	<tr>
		<td valign="top" width="70%" class="columnheader" bgcolor="<%=gsColorHighlight%>"><%=dictLanguage("Project")%></td>
		<td valign="top" width="30%" class="columnheader" bgcolor="<%=gsColorHighlight%>"><%=dictLanguage("Client")%></td>
	</tr>
 
<%
intRowCounter = 0
Do While Not rsViewProjects.EOF
	intRowCounter = intRowCounter + 1
%>

	<tr <%If intRowcounter MOD 2 = 1 then %>bgcolor="<%=gsColorWhite%>"<%Else%>bgcolor="#FFFFFF"<%End If%>>
		<td valign=top>
			<a href="project-edit.asp?project_id=<%=rsViewProjects("Project_ID")%>"><%=rsViewProjects("Description")%></a>
		</td>
		<td valign=top>
			<a href="<%=gsSiteRoot%>clients/client-edit.asp?client_id=<%=rsViewProjects("Client_ID")%>"><%=rsViewProjects("Client_Name")%></a>
		</td>
	</tr>
<%
	rsViewProjects.MoveNext
Loop
rsViewProjects.close
set rsViewProjects = nothing %>
</table>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<!--#include file="../includes/main_page_close.asp"-->
