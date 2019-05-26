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


<!--#include file="../includes/main_page_open.asp"-->


<table cellpadding="2" cellspacing="2" border=0 align="center">
	<tr><td colspan="3" align="center" bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("View_Tasks")%>: <%=dictLanguage("Options")%></b></td></tr>
	<tr>
		<td colspan="3"><b class="alert">
<%'display any Message and clear it
if session("Msg") <> "" then
	response.write session("Msg")
	session("Msg") = ""
end if %></b>
		</td>
	<tr>
		<td><%=dictLanguage("View_Tasks_Client")%></td>
		<td align="center"><a href="view1a.asp?active=1"><%=dictLanguage("Active")%></a></td>
		<td align="center"><a href="view1a.asp?active=0"><%=dictLanguage("Archived")%></a></td>
	</tr>
	<tr>
		<td><%=dictLanguage("View_All_Tasks")%> -- <%=dictLanguage("By_Assignor")%></td>
		<td align="center"><a href="view_all_by_assigner.asp"><%=dictLanguage("Active")%></a></td>
		<td align="center"><a href="view_all_by_assigner.asp?active=0"><%=dictLanguage("Archived")%></a></td>
	</tr>
	<tr>
		<td><%=dictLanguage("View_All_Tasks")%> -- <%=dictLanguage("By_Client")%></td>
		<td align="center"><a href="view_all.asp"><%=dictLanguage("Active")%></a></td>
		<td align="center"><a href="view_all.asp?active=0"><%=dictLanguage("Archived")%></a></td>
	</tr>
	<tr>
		<td><%=dictLanguage("View_All_Tasks")%> -- <%=dictLanguage("By_Employee")%></td>
		<td align="center"><a href="view_all_by_employee.asp"><%=dictLanguage("Active")%></a></td>
		<td align="center"><a href="view_all_by_employee.asp?active=0"><%=dictLanguage("Archived")%></a></td>
	</tr>
	<tr>
		<td><%=dictLanguage("View_My_Tasks")%></td>
		<td align="center"><a href="view_mine.asp"><%=dictLanguage("Active")%></a></td>
		<td align="center"><a href="view_mine.asp?active=0"><%=dictLanguage("Archived")%></a></td>
	</tr>
	<tr>
		<td><%=dictLanguage("View_Tasks_I_Assigned")%></td>
		<td align="center"><a href="view_my_assigned.asp"><%=dictLanguage("Active")%></a></td>
		<td align="center"><a href="view_my_assigned.asp?active=0"><%=dictLanguage("Archived")%></a></td>
	</tr>
	<tr>
		<form action="view_employee.asp" method="POST" id=form2 name=form2>
		<td>
			<Select name="employee" class="formstyleLong">
				<option value="">--<%=dictLanguage("View_Tasks_Employee")%>--</option>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'query the database for all employees and build a select menu from the recordset
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetActiveEmployees()
Call RunSQL(sql, rsEmployees)
while not rsEmployees.EOF %>
				<option value="<%=rsEmployees("employee_id")%>"><%=rsEmployees("employeeName")%></option>
<%	rsEmployees.MoveNext
wend 
rsEmployees.close
set rsEmployees = nothing %>
			</select>
		</td>
		<td align="center"><input type=Submit value="Active" class="formButton" id=Submit2 name=Submit></td>
		<td align="center"><input type=Submit value="Archived" class="formButton" id=Submit2 name=Submit></td>
		</form>
	</tr>
	<tr>
		<td><a href="<%=gsSiteRoot%>reports/task_graph.asp"><%=dictLanguage("Task_Graph_View")%></a></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr><td colspan="3">&nbsp;</td></tr>
	<td colspan="3" align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></td></tr>
</table>

<!--#include file="../includes/main_page_close.asp"-->