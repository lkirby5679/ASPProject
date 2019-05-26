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

<%
pto_ID = request.querystring("pto_ID")

if pto_ID = "" then
	response.redirect("ptoform.asp")
end if

sql = sql_GetPTOByID(pto_ID)
Call RunSQL(sql, rs)
%>

<font face="Arial" color="Red"><%=Session("strErrorMessage")%></font>
<% Session("strErrorMessage") = ""%>

<form method="post" action="ptoapproval-processor.asp" name="strForm" id="strForm">
<table border="0" cellpadding="1" cellspacing="1">
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("PTO_Form")%></b></td></tr>
		<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Employee")%>:</b></td>
			<td><select name="employee" class="formstylelong">
<%	strEmployeesSQL = sql_GetActiveEmployees()
	Call RunSQL(strEmployeesSQL, rsEmployees)
	while not rsEmployees.EOF %>
					<option value="<%=rsEmployees("Employee_ID")%>"<%if pto_ID <>"" then if rsEmployees("Employee_ID")=rs("Employee_ID") then response.write(" Selected ") end if end if%>><%=rsEmployees("EmployeeName")%></option>
	<%	rsEmployees.MoveNext
	wend
	rsEmployees.close
	set rsEmployees = nothing %>
				</select>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Date")%>:</b></td>
			<td>
				<input type="Text" name="Start_Date" class="formstylemed" value="<%if pto_ID<>"" then%><%=rs("Start_Date")%><%else if session("Start_Date")="" then%><%=Date()%><%else%><%=session("Start_Date")%><%end if%><%end if%>" onKeyPress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Start_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Time")%>:</b></td>
			<td><input type="Text" name="Start_Time" class="formstylemed" value="<%if pto_ID<>"" then%><%=timevalue(rs("Start_Time"))%><%else if session("Start_Time")="" then%><%=Time()%><%else%><%=session("Start_Time")%><%end if%><%end if%>" onKeyPress="txtTime_onKeypress();" maxlength="11"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Date")%>:</b></td>
			<td>
				<input type="Text" name="End_Date" class="formstylemed" value="<%if pto_ID<>"" then%><%=rs("End_Date")%><%else if session("End_Date")="" then%><%=Date()%><%else%><%=session("End_Date")%><%end if%><%end if%>" onKeyPress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','End_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Time")%>:</b></td>
			<td><input type="Text" name="End_Time" class="formstylemed" value="<%if pto_ID<>"" then%><%=timevalue(rs("End_Time"))%><%else if session("End_Time")="" then%><%=Time()%><%else%><%=session("End_Time")%><%end if%><%end if%>" onKeyPress="txtTime_onKeypress();" maxlength="11"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Total_Hours")%>:</b></td>
			<td><input type="Text" name="Total_Hours" class="formstylemed" value="<%if pto_ID<>"" then%><%=rs("Total_Hours")%><%else%><%=session("Total_Hours")%><%end if%>"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Paid")%>:</b></td>
			<td><input type="checkbox" name="Paid" value="1" <%if pto_ID<>"" then%><%if rs("Paid")=1 or rs("Paid")=-1 or rs("Paid") then Response.Write " Checked"%><%end if%>> <%=dictLanguage("Yes")%></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Current_PTO_Balance")%>:</b></td>
			<td><input type="Text" name="Balance" class="formstylemed" value="<%if pto_ID<>"" then%><%=rs("Balance")%><%else%><%=Session("Balance")%><%end if%>"></td>
		</tr>
		<tr>
			<td valign="top"><b class="bolddark"><%=dictLanguage("Reason")%>:</b></td>
			<td><TEXTAREA NAME="Reason" class="formstylelong" rows="5"><%if pto_ID<>"" then%><%=rs("Reason")%><%else%><%=Session("Reason")%><%end if%></TEXTAREA></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Excused")%>:</b></td>
			<td><input type="checkbox" name="Excused" value="1" <%if pto_ID<>"" then%><%if rs("Excused")=1 or rs("Excused")=-1 or rs("Excused") then Response.Write "Checked"%><%end if%>> <%=dictLanguage("Yes")%></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Approved")%>:</b></td>
			<td><input type="checkbox" name="Approved" value="1" <%if pto_ID<>"" then%><%if rs("Approved")=1 or rs("Approved")=-1 or rs("Approved") then Response.Write "Checked"%><%end if%>> <%=dictLanguage("Yes")%></td>
		</tr>
		<tr>
			<td colspan="2" align="center">
				<input type="Hidden" name="pto_ID" value="<%=pto_ID%>">
				<input type="Submit" name="Submit" value="Submit" class="formbutton">			
			</td>
		</tr>
</table>
</form>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<br>

<!--#include file="../includes/main_page_close.asp"-->