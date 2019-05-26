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
active = Request.QueryString("Active")
if active = "" or active = "Active" then
	active = "Active"
	activeView = dictLanguage("Active")
elseif active = "All" then
	active = "All"
	activeView = dictLanguage("All")
end if
if active = "Active" then
	sql = sql_GetActiveClientsByRepID(session("userid"))
	active = "My Active"
else
	sql = sql_GetAllClientsByRepID(session("userid"))
	active = "All My"
end if

Call RunSQL(sql, rsViewClients)
%>

<!--#include file="../includes/main_page_open.asp"-->

<table width="100%" border=0 cellpadding=2 cellspacing=2 style="border-style: solid; border-width: 2; border-color: <%=gsColorBlack%>">
	<tr>
		<td valign=top colspan=6>
			<b class="header"><%=dictLanguage("View")%>&nbsp;<%=activeview%>&nbsp;<%=dictLanguage("Clients")%><br>
			<%if active="My Active" then%>
				<a href="view_mine.asp?active=All" class="small"><%=dictLanguage("View_All_My_Clients")%></a>
			<%else%>
				<a href="view_mine.asp?active=Active" class="small"><%=dictLanguage("View_My_Active_Clients")%></a>
			<%end if%>
		</td>
	</tr>
	<tr>
		<td valign="top" width="70%" class="columnheader" bgcolor="<%=gsColorHighlight%>"><%=dictLanguage("Client")%></td>
		<td valign="top" width="30%" class="columnheader" bgcolor="<%=gsColorHighlight%>"><%=dictLanguage("Account_Rep")%></td>
	</tr>
 
<%
intRowCounter = 0
Do While Not rsViewClients.EOF
	intRowCounter = intRowCounter + 1
%>

<tr <%If intRowcounter MOD 2 = 1 then %>bgcolor="<%=gsColorWhite%>"<%Else%>bgcolor="#FFFFFF"<%End If%>>


	<td valign=top>
		<a href="client-edit.asp?client_id=<%=rsViewClients("Client_ID")%>"><%=rsViewClients("Client_Name")%></a>
	</td>
	<td valign=top><%=rsViewClients("EmployeeName")%></td>
</tr>
<%
	rsViewClients.MoveNext
Loop
rsViewClients.close
set rsViewClients = nothing%>
</table>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<!-- #include file="../includes/main_page_close.asp"-->