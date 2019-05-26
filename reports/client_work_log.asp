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
		
<form method="POST" action="client_work_log_results.asp" name="strForm" id="strForm">
<table align="center" cellpadding="2" cellspacing="2" border="0">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Work_Log_Internal")%></td></tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Client")%>?</td>
		<td>			
			<select name="client_id" class="formstyleXL">
				<option value=""><%=dictLanguage("All_Sites")%>
<%	Call RunSQL(sqlRS, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("cl_id")%>P<%=rs("Project_ID")%>"><%=rs("Client_Name")%>{<%=rs("Description")%>}</option>
<%		rs.MoveNext
	Loop
	rs.close
	Set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Project_Type")%>?</td>
		<td>
			<select name="proj_type_id" class="formstyleLong">
				<option value=""><%=dictLanguage("All_Project_Types")%>
<%	Call RunSQL(sqlType, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("ProjectType_ID")%>"><%=rs("ProjectTypeDescription")%>
<%		rs.MoveNext
	Loop
	rs.close
	set rs = Nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Timecard_Type")%>?</td>
		<td>
			<select name="work_type_id" class="formstyleLong">
				<option value=""><%=dictLanguage("All_Timecard_Types")%>
<%	Call RunSQL(sqlWork, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("TimeCardType_ID")%>"><%=rs("TimeCardTypeDescription")%>
<%		rs.MoveNext
	Loop
	rs.close
	Set rs = Nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Department")%>?</td>
		<td>
			<select name="department" class="formstyleLong">
				<option value=""><%=dictLanguage("All_Departments")%></option>
<%	Call RunSQL(sqlDept, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("department")%>"><%=rs("department")%></option>
<%		rs.MoveNext
	Loop
	rs.close
	set rst = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Employee")%>?</td>	
		<td>
			<select name="emp_id" class="formstyleLong">
				<option value=""><%=dictLanguage("All_Employees")%>
<%	Call RunSQL(sqlEmp, rs)
	Do While Not rs.EOF %>
				<option value="<%=rs("employee_id")%>"><%=rs("EmployeeName")%>
<%		rs.MoveNext
	Loop %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Which_Account_Rep")%>?</td>
		<td>
			<select name="rep_id" class="formstyleLong">
				<option value=""><%=dictLanguage("All_Account_Reps")%>
<%	if not (rs.eof and rs.bof) then 
		rs.MoveFirst
	end if
	Do While Not rs.eof %>
				<option value="<%=rs("employee_id")%>"><%=rs("EmployeeName")%>
<%		rs.MoveNext
	Loop %>
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


