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

<%Call ConfirmPermission("permPTOAdmin", "")%>

<%
'display any Message and clear it
if session("Msg") <> "" then
	response.write "<br><p align=""center"" class=""alert"">" & session("Msg") & "</p>"
	session("Msg") = ""
end if
%>

<table border="0" cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("View_PTO_Requests")%></b></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<td colspan="2"><a href="view_all_by_employee.asp" class="bolddark"><%=dictLanguage("View_PTO_By_Employee")%></a></td>
	</tr>
	<tr>
		<td colspan="2"><a href="view_all_not_complete.asp" class="bolddark"><%=dictLanguage("View_PTO_Not_Completed_By_Employee")%></a></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<td colspan="2">
			<form action="view_employee.asp" method="POST" name="strForm2" id="strForm2">
				<Select name="employee" class="formstylelong">
<%
strEmployeesSQL = sql_GetActiveEmployees()
Call RunSQL(strEmployeesSQL, rsEmployees)
while not rsEmployees.EOF %>
					<option value="<%=rsEmployees("employee_id")%>"><%=rsEmployees("employeeName")%></option>
<%	rsEmployees.MoveNext
wend
rsEmployees.close
set rsEmployees = nothing
%>
				</select>
				<input type=Submit value="<%=dictLanguage("View_PTO_For_Employee_Button")%>" class="formbutton">
			</form>
		</td>
	</tr>
	<tr>
		<td colspan="2">
			<form method = "POST" action = "view_all_by_dates.asp" name="strForm" id="strForm">
				<table cellpadding="2" cellspacing="2" border="0">
					<tr>
						<td><b class="bolddark"><%=dictLanguage("Start_Date")%>:</b></td>
						<td>
							<input type="text" name="Start_Date" value="<%=date()%>" class="formstylemed" onKeyPress="txtDate_onKeypress();" maxlength="10">
							<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Start_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
						</td>
						<td rowspan="2"><input type="submit" value="GO!" class="formbutton" id=submit1 name=submit1></td>
					</tr>
					<tr>
						<td><b class="bolddark"><%=dictLanguage("End_Date")%>:</b></td>
						<td>
							<input type="text" name="End_Date" value="<%=date()%>" class="formstylemed" onKeyPress="txtDate_onKeypress();" maxlength="10">
							<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','End_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<br>

<!--#include file="../includes/main_page_close.asp"-->