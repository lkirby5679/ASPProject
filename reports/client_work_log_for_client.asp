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
sqlRS = sql_GetActiveProjectsWithClients()
sqlType = sql_GetProjectTypes()
sqlWork = sql_GetActiveTimecardTypes()
sqlDept = sql_GetActiveDepartments()
sqlEmp = sql_GetActiveEmployees()
%>

<!--#include file="../includes/main_page_open.asp"-->


<%	if session("strErrorMessage")<>"" then 
		Response.Write "<P class='alert'>" & session("strErrorMessage") & "<P>"
		Session("strErrorMessage") = ""
	end if %>
		
<form method="POST" action="client_work_log_results_for_client.asp" id="strForm" name="strForm">
<table align="center" cellpadding="2" cellspacing="2" border="0">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Work_Log_For_Client")%></td></tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Client")%>?</td>
		<td>			
			<select name="client_id" class="formstyleXL">
				<option value=""><%=dictLanguage("All_Sites")%></option>		
<%	Call RunSQL(sqlRS, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("cl_id")%>P<%=rs("Project_ID")%>"><%=rs("Client_Name")%>{<%=rs("Description")%>}</option>
<%		rs.MoveNext
	Loop
	rs.Close
	set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Employee")%>?</td>
		<td>
			<select name="emp_id" class="formstyleLong">
				<option value=""><%=dictLanguage("All_Employees")%></option>
<%	Call RunSQL(sqlEmp, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("employee_id")%>"><%=rs("EmployeeName")%>
<%		rs.MoveNext
	Loop
	rs.Close
	set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Timecard_Type")%>?</td>
		<td>
			<select name="typeofwork_id" class="formstyleLong">
				<option value=""><%=dictLanguage("All_Timecard_Types")%></option>
<%	Call RunSQL(sqlWork, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("TimeCardType_ID")%>"><%=rs("TimeCardTypeDescription")%>
<%		rs.MoveNext
	Loop
	rs.Close
	set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Reconciled")%>:</td>
		<td>
			<input type="radio" name="reconciled" value="All" Checked>All
			<input type="radio" name="reconciled" value="Yes">Yes
			<input type="radio" name="reconciled" value="No">No
		</td>
	</tr>	
	<tr>
		<td class="bolddark"><%=dictLanguage("Start_Date")%>:</td>
		<td>
			<input type="text" name="start_date" size=10 class="formstyleShort" onKeyPress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','start_date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("End_Date")%>:</td>
		<td>
			<input type="text" name="end_date" size=10 class="formstyleShort" onKeyPress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','end_date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr><td colspan="2" align="center"><input type="submit" value="GO!" class="formButton"></td></tr>
</Table>
</form>
<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<!--#include file="../includes/main_page_close.asp"-->
