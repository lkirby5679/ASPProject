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
''''''''''''''''''''''''''''''''''''''''
'This page allows Hourly Employees to Edit their Time Sheet Data
''''''''''''''''''''''''''''''''''''''''
id = Request.QueryString("id")
empID = Request.QueryString("empID")
%>

<!--#include file="../includes/main_page_open.asp"-->

<%
if session("strErrorMessage")<>"" then 
	Response.Write "<P class='alert'>" & session("strErrorMessage") & "<P>"
	Session("strErrorMessage") = ""
end if %>

<form action="edit2.asp" method="POST" id="strForm" name="strForm">
<table cellpadding="2" cellspacing="2" border="0" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td align="center" colspan="2" class="homeheader"><%=dictLanguage("Timesheet_Details")%></td></tr>
	<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
<%
sql = sql_GetTimesheetsByID(id)
Call runSQL(sql, adoRS)
if not adoRS.eof then 
	day1 = datevalue(adoRS("DateHoursLogged"))
	if adoRS("in1")<>"" then
		in1 = timevalue(adoRS("In1"))
		in1Hour   = hour(in1)
		in1Minute = minute(in1)
		if in1Hour > 11 then
			in1AMPM = "PM"
			in1Hour = in1Hour - 12
		else
			in1AMPM = "AM"
		end if	
	else
		in1 = ""
	end if
	if adoRS("in2")<>"" then
		in2 = timevalue(adoRS("in2"))
		in2Hour   = hour(in2)
		in2Minute = minute(in2)	
		if in2Hour > 11 then
			in2AMPM = "PM"
			in2Hour = in2Hour - 12
		else
			in2AMPM = "AM"
		end if			
	else
		in2 = ""
	end if
	if adoRS("out1")<>"" then
		out1 = timevalue(adoRS("out1"))
		out1Hour   = hour(out1)
		out1Minute = minute(out1)	
		if out1Hour > 11 then
			out1AMPM = "PM"
			out1Hour = out1Hour - 12
		else
			out1AMPM = "AM"
		end if				
	else
		out1 = ""
	end if
	if adoRS("out2")<>"" then
		out2 = timevalue(adoRS("out2"))
		out2Hour   = hour(out2)
		out2Minute = minute(out2)		
		if out2Hour > 11 then
			out2AMPM = "PM"
			out2Hour = out2Hour - 12
		else
			out2AMPM = "AM"
		end if							
	else
		out2 = ""
	end if
%>
	<TR>
		<td class="bolddark"><%=dictLanguage("Date")%>:<font class="alert">*</font></td>
		<td>
			<input type=text name="datehourslogged" value="<%=day1%>" class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','datehourslogged',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<TR>
		<td class="bolddark"><%=dictLanguage("In_1")%>:</td>
		<td>
			<select name="in1Hour" class="formstyleTiny">
<%	for i = 1 to 12
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = in1Hour then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="in1Minute" class="formstyleTiny">
<%	for i = 0 to 59
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = in1Minute then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="in1AMPM" class="formstyleTiny">
				<option value="AM"<%if in1AMPM = "AM" then Response.Write " Selected"%>>AM</option>
				<option value="PM"<%if in1AMPM = "PM" then Response.Write " Selected"%>>PM</option>
			</select>								
		</td>
	</tr>
	<TR>
		<td class="bolddark"><%=dictLanguage("Out_1")%>:</td>
		<td>
			<select name="out1Hour" class="formstyleTiny">
<%	for i = 1 to 12
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = out1Hour then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="out1Minute" class="formstyleTiny">
<%	for i = 0 to 59
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = out1Minute then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="out1AMPM" class="formstyleTiny">
				<option value="AM"<%if out1AMPM = "AM" then Response.Write " Selected"%>>AM</option>
				<option value="PM"<%if out1AMPM = "PM" then Response.Write " Selected"%>>PM</option>
			</select>					
		</td>
	</tr>
	<TR>
		<td class="bolddark"><%=dictLanguage("In_2")%>:</td>
		<td>
			<select name="in2Hour" class="formstyleTiny">
<%	for i = 1 to 12
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = in2Hour then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="in2Minute" class="formstyleTiny">
<%	for i = 0 to 59
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = in2Minute then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="in2AMPM" class="formstyleTiny">
				<option value="AM"<%if in2AMPM = "AM" then Response.Write " Selected"%>>AM</option>
				<option value="PM"<%if in2AMPM = "PM" then Response.Write " Selected"%>>PM</option>
			</select>			
		</td>
	</tr>
	<TR>
		<td class="bolddark"><%=dictLanguage("Out_2")%></td>
		<td>
			<select name="out2Hour" class="formstyleTiny">
<%	for i = 1 to 12
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = out2Hour then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="out2Minute" class="formstyleTiny">
<%	for i = 0 to 59
		if len(i) < 2 then
			iShow = "0" & i
		else
			iShow = i
		end if 
		Response.Write "<option value='" & iShow & "'"
		if i = out2Minute then 
			Response.Write " Selected"
		end if
		Response.Write ">" & iShow & "</option>"
	next %>
			</select>
			<select name="out2AMPM" class="formstyleTiny">
				<option value="AM"<%if out2AMPM = "AM" then Response.Write " Selected"%>>AM</option>
				<option value="PM"<%if out2AMPM = "PM" then Response.Write " Selected"%>>PM</option>
			</select>	
		</td>
	</tr>								
	<%if session("permTimesheetsEdit") or trim(session("employee_id"))=trim(empID) then%>
	<tr>
		<td colspan="2" align="center">
			<input type=hidden name="id" value="<%=id%>">
			<input type="hidden" name="empID" value="<%=empID%>">
			<input type="Submit" value="Submit" class="formButton">		
		</td>
	</tr>
	<%end if%>
<%
end if
adoRS.close
set adoRS = nothing%>
</table>
</form>
<p align="center">
<a href="view.asp"><%=dictLanguage("Hourly_Timesheets")%></a><br>
<a href="main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->
