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

<%if session("strErrorMessage")<>"" then
	Response.Write "<font class=""alert"">" & Session("strErrorMessage") & "</font>"
	Session("strErrorMessage") = ""
  end if %>

<form method="post" action="pto-prelim-processor.asp" name="strForm" id="strForm">
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("Preliminary_PTO_Form")%></b></td></tr>
		<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Employee")%>:</b><font class="alert">*</font></td>
			<td><select name="employee" class="formstylelong">
	<%
	'''''''''''''''''''''''''''''''''''''
	'Query the database and bring back a recordset
	'containing all of the employees' names
	'and their IDs
	'''''''''''''''''''''''''''''''''''''
	strEmployeesSQL = sql_GetActiveEmployees()
	Call RunSQL(strEmployeesSQL, rsEmployees)
	while not rsEmployees.EOF %>
					<option value="<%=rsEmployees("Employee_ID")%>"><%=rsEmployees("EmployeeName")%></option>
	<%
		rsEmployees.MoveNext
	wend
	rsEmployees.close
	%>			</select>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Date")%>:</b></td>
			<td>
				<input type="Text" name="Start_Date" value="<%if session("Start_Date")="" then%><%=Date()%><%else%><%=session("Start_Date")%><%end if%>" class="formstylemed" onKeyPress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Start_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Time")%>:</b></td>
			<td><input type="Text" name="Start_Time" class="formstylemed" value="<%if session("Start_Time")="" then%><%=Time()%><%else%><%=session("Start_Time")%><%end if%>" onKeyPress="txtTime_onKeypress();" maxlength="11"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Date")%>:</b></td>
			<td>
				<input type="Text" name="End_Date" class="formstylemed" value="<%if session("End_Date")="" then%><%=Date()%><%else%><%=session("End_Date")%><%end if%>" onKeyPress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','End_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Time")%>:</b></td>
			<td><input type="Text" name="End_Time" class="formstylemed" value="<%if session("End_Time")="" then%><%=Time()%><%else%><%=session("End_Time")%><%end if%>" onKeyPress="txtTime_onKeypress();" maxlength="11"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Total_Hours")%>:</b></td>
			<td><input type="Text" name="Total_Hours" class="formstyleshort" value="<%if session("Total_Hours")<>"" then Response.Write session("Total_Hours") else Response.Write "0"%>"  onkeypress="txtDate_onKeypress();" maxlength="3"></td>
		</tr>
		<tr>
			<td colspan="2" align="center">
				<input type="Submit" name="Submit" value="Submit" class="formbutton">
			</td>
		</tr>
</table>
</form>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<br>

<!--#include file="../includes/main_page_close.asp"-->

