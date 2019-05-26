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
sql = sql_DeleteTasks(Request.QueryString("id"))
Call DoSQL(sql)
%>

<!--#include file="../includes/main_page_open.asp"-->

<table cellpadding="3" cellspacing="3" border="0" align="center">
	<tr><td align="center" bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("Task_Details")%></b></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td><%=dictLanguage("Deleted_Task")%></td></tr>
</table>
<p align="center">
<a href="view.asp"><%=dictLanguage("View_Tasks")%></a><br>
<a href="view_my_assigned.asp"><%=dictLanguage("View_Tasks_I_Assigned")%></a><br>
<a href="<%=gsSiteRoot%>projects/project-status.asp"><%=dictLanguage("Project_Status")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->