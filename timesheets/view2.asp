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
Function GetEmpName(ID)
	if ID <> "" then
		sql = sql_GetEmployeesByID(ID)
		'response.write sql
		Call RunSQL(sql, rs)
		if not rs.eof then
			empName = rs("EmployeeName")
		else
			empName = ""
		end if
		rs.close
		set rs = nothing
	else
		empName = ""
	end if
	GetEmpName = empName
End Function

''''''''''''''''''''''''''''
'This page outputs current week timeSheet data
'for hourly employees
''''''''''''''''''''''''''''
empID = Request("employee")
empName = GetEmpName(empID)


if empID = "" or empName = "" then
	session("errMsg") = dictLanguage("Error_TimesheetNoEmp")
	Response.Redirect "view.asp"
end if

''''''''''''''''''''''''''''
'What is the current pay period?
''''''''''''''''''''''''''''
sql = sql_GetCurrentPayPeriod()
Call RunSQL(sql, rsCurrentPayPeriod)
if not rsCurrentPayPeriod.eof then
	datFirstDayOfCurrentPeriod = CDate(rsCurrentPayPeriod("StartDate"))
	datLastDayOfCurrentPeriod = CDate(rsCurrentPayPeriod("EndDate"))
	intRecordNumber = rsCurrentPayPeriod("ID")
	datLastDayOfPreviousPeriod = CDate(DateAdd("d",datFirstDayOfCurrentPeriod,-1))
end if
rsCurrentPayPeriod.close
set rsCurrentPayPeriod = nothing

if datFirstDayOfCurrentPeriod = "" or datLastDayOfCurrentPeriod = "" or datLastDayOfPreviousPeriod = "" then
	session("errMsg") = dictLanguage("Error_TimesheetNoPayPeriod")
	Response.Redirect "view.asp"
end if

sql = sql_GetPayPeriodsByEndDate(datLastDayOfPreviousPeriod)
Call RunSQL(sql, rsPreviousPayPeriod)
if not rsPreviousPayPeriod.eof then
	datFirstDayOfPreviousPeriod = CDate(rsPreviousPayPeriod("StartDate"))
	'datLastDayOfPreviousPeriod = rsPreviousPayPeriod("EndDate")
	session("intRecordNumber") = rsPreviousPayPeriod("ID")
end if
rsPreviousPayPeriod.close
set rsPreviousPayPeriod = nothing

'if datFirstDayOfPreviousPeriod = "" or datLastDayOfPreviousPeriod = "" then
'	Response.Redirect "view.asp"
'end if


%>

<!--#include file="../includes/main_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border="0" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td align="center" class="homeheader"><%=dictLanguage("Timesheet_Details")%></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td class="homeheader"><%=dictLanguage("Timesheet_For")%>&nbsp;<%=empName%>&nbsp;<%=dictLanguage("For_PayPeriod_Beginning")%>&nbsp;<%=datFirstDayOfCurrentPeriod%>:</td></tr>
	<tr>
		<td>
			<table cellpadding="2" cellspacing="2" border="0" width="600">
				<tr bgcolor="<%=gsColorHighlight%>">
					<td class="columnheader"><%=dictLanguage("Date")%></td>
					<td class="columnheader"><%=dictLanguage("In_1")%></td>
					<td class="columnheader"><%=dictLanguage("Out_1")%></td>
					<td class="columnheader"><%=dictLanguage("In_2")%></td>
					<td class="columnheader"><%=dictLanguage("Out_2")%></td>
					<td class="columnheader"><%=dictLanguage("Total_Hours")%></td>
				</tr>
<%	
sql = sql_GetTimesheetsByEmployeeIDStartDate(empID, datFirstDayOfCurrentPeriod)
Call RunSQL(sql, adoRS)
do while not adoRS.eof
	day1 = datevalue(adoRS("DateHoursLogged"))
	if adoRS("in1")<>"" then
		in1 = timevalue(adoRS("In1"))
	else
		in1 = ""
	end if
	if adoRS("in2")<>"" then
		in2 = timevalue(adoRS("in2"))
	else
		in2 = ""
	end if
	if adoRS("out1")<>"" then
		out1 = timevalue(adoRS("out1"))
	else
		out1 = ""
	end if
	if adoRS("out2")<>"" then
		out2 = timevalue(adoRS("out2"))
	else
		out2 = ""
	end if
	varDisplayTimeWorked = tallyHours(In1,Out1,In2,Out2,Day1) 
	if strBGColor = "#FFFFFF" then
		strBGColor = "#EEEEEE"
	else
		strBGColor = "#FFFFFF"
	end if 	%>	
				<TR bgcolor="<%=strBGColor%>">
					<td class="small"><%=day1%></td>
					<td class="small"><%=In1%></td>
					<td class="small"><%=Out1%></td>
					<td class="small"><%=In2%></td>
					<td class="small"><%=Out2%></td>
					<td class="small"><a href="<%=gsSiteRoot%>timesheets/edit.asp?id=<%=adoRS("id")%>&empID=<%=empID%>"><%=varDisplayTimeWorked%></a></td>
				</tr>
<%	if not IsNumeric(varDisplayTimeWorked) and CDate(adoRS("DateHoursLogged")) <> Date() then
		boolDisplayExplanation = True
	elseif IsNumeric(varDisplayTimeWorked) then
		intTotalHoursWorked = CDbl(intTotalHoursWorked) + CDbl(varDisplayTimeWorked)
	end if
		adoRS.movenext
loop 
adoRS.close
set adoRS = nothing %>
				<TR>
					<td class="bolddark"><%=dictLanguage("Total")%></td>
					<td><b><%=intTotalHoursWorked%></b></td>
				</tr>
			</table>
<%	
intTotalHoursWorked = 0
if boolDisplayExplanation = True then %>
			<br>
			<p class="alert"><%=dictLanguage("Timesheet_Inst_1")%></p>
<%end if%>
		</td>
	</tr>
<% if datFirstDayOfPreviousPeriod <> "" and datLastDayOfPreviousPeriod <> "" then %> 	
	<tr><td>&nbsp;</td></tr>
	<tr><tr><td class="homeheader"><%=dictLanguage("Timesheet_For")%>&nbsp;<%=empName%>&nbsp;<%=dictLanguage("For_PayPeriod_Beginning")%>&nbsp;<%=datFirstDayOfPreviousPeriod%>:</td></tr>
	<tr>
		<td>
			<table cellspacing="2" cellspacing="2" border="0" width="600">
				<tr bgcolor="<%=gsColorHighlight%>">
					<td class="columnheader"><%=dictLanguage("Date")%></td>
					<td class="columnheader"><%=dictLanguage("In_1")%></td>
					<td class="columnheader"><%=dictLanguage("Out_1")%></td>
					<td class="columnheader"><%=dictLanguage("In_2")%></td>
					<td class="columnheader"><%=dictLanguage("Out_2")%></td>
					<td class="columnheader"><%=dictLanguage("Total_Hours")%></td>
				</tr>
<%
sql = sql_GetTimesheetsByEmployeeIDStartDateEndDate(empID, datFirstDayOFPreviousPeriod, datLastDayOfPreviousPeriod)
Call RunSQL(sql, adoRS)
do while not adoRS.eof
	day1 = datevalue(CDATE(adoRS("DateHoursLogged")))
	if adoRS("in1")<>"" then
		in1 = timevalue(adoRS("In1"))
	else
		in1 = ""
	end if
	if adoRS("in2")<>"" then
		in2 = timevalue(adoRS("in2"))
	else
		in2 = ""
	end if
	if adoRS("out1")<>"" then
		out1 = timevalue(adoRS("out1"))
	else
		out1 = ""
	end if
	if adoRS("out2")<>"" then
		out2 = timevalue(adoRS("out2"))
	else
		out2 = ""
	end if
	varDisplayTimeWorked = tallyHours(In1,Out1,In2,Out2,DateHoursLogged) 
	if strBGColor = "#FFFFFF" then
		strBGColor = "#EEEEEE"
	else
		strBGColor = "#FFFFFF"
	end if %>
				<TR bgcolor="<%=strBGColor%>">
					<td class="small"><%=Day1%></td>
					<td class="small"><%=In1%></td>
					<td class="small"><%=Out1%></td>
					<td class="small"><%=In2%></td>
					<td class="small"><%=Out2%></td>
					<td class="small"><a href="edit.asp?id=<%=adoRS("id")%>&empID=<%=empID%>"><%=varDisplayTimeWorked%></a></td>
				</tr>
<%
	if not IsNumeric(varDisplayTimeWorked) and CDate(adoRS("DateHoursLogged")) <> Date() then
		boolDisplayExplanation = True
	elseif IsNumeric(varDisplayTimeWorked) then
		intTotalHoursWorked = CDbl(intTotalHoursWorked) + CDbl(varDisplayTimeWorked)
	end if
	adoRS.movenext
loop 
adoRs.close
set adoRS = nothing %>
				<TR>
					<td class="bolddark"><%=dictLanguage("Total")%></td>
					<td><b><%=intTotalHoursWorked%></b></td>
				</tr>
			</table>
<%if boolDisplayExplanation = True then%>
			<p class="alert"><%=dictLanguage("Timesheet_Inst_1")%></p>
<%end if%>
		</td>
	</tr>
<%end if%>	
</table>

<p align="center">
<%if datFirstDayOfPreviousPeriod <> "" and datLastDayOfPreviousPeriod <> "" then%>
	<a href="view3.asp?id=<%=empID%>&name=<%=empName%>&PreviousFirstOfWeek=<%=datFirstDayOfPreviousPeriod%>"><%=dictLanguage("View_Timesheet_History")%></a><br>
<%end if%>
<a href="view.asp"><%=dictLanguage("Hourly_Timesheets")%></a><br>
<a href="main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<%
'''''''''''''''''''''''''''''''''''''
'This function tallies a daily hours-worked count
'
'takes:
'datIn1 - a time value representing employee's first clock in
'datOut1 - a time value representing employee's first clock out
'datIn2 - a time value representing employee's last clock in
'datOut2 - a time value representing employee's last clock out
'datDate - a date value representing the day in question
'''''''''''''''''''''''''''''''''''''
Function tallyHours(datIn1,datOut1,datIn2,datOut2,datDate)
	'''''''''''''''''''''''''''''''''''''
	'How many times did the user clock in, out for this day,
	'and format in and out times as Date/Time data
	'''''''''''''''''''''''''''''''''''''
	datIn1 = CDate(datIn1)
	if len(datOut1 & "***") > 3 then
		datOut1 = CDate(datOut1)
		boolDatOut1Exists = True
	end if
	if len(datIn2 & "***") > 3 then
		datIn2 = CDate(datIn2)
		boolDatIn2Exists = True
	end if
	if len(datOut2 & "***") > 3 then
		datOut2 = CDate(datOut2)
		boolDatOut2Exists = True
	end if
	'''''''''''''''''''''''''''''''''''''
	'The following represent special cases where the
	'employee forgot to clock in or out the appropriate
	'number of times
	'''''''''''''''''''''''''''''''''''''
	if (boolDatOut1Exists = True and boolDatIn2Exists = True and boolDatOut2Exists <> True) or (boolDatOut1Exists <> True and boolDatIn2Exists = True) or (boolDatOut1Exists = True and boolDatIn2Exists <> True and boolDatOut2Exists = True) and (boolDatOut1Exists <> True and boolDatOut2Exists <> True) or (boolDatOut1Exists <> True) then
		if CDate(datDate) <> Date() Then
			tallyHours = "*"
		else
			tallyHours = dictLanguage("Workday_In_Progress")
		end if
	else
		'''''''''''''''''''''''''''''''''''''
		'If no lunch break
		'''''''''''''''''''''''''''''''''''''
		If boolDatOut2Exists <> True then
			On Error Resume Next
			if datOut1 < datIn1 then
				intMinutesWorked = CInt(CInt(DateDiff("n",datIn1,datOut1)))
			else
				intMinutesWorked = CInt(DateDiff("n",datIn1,datOut1))
			end if
		'''''''''''''''''''''''''''''''''''''
		'If they did take a lunch break
		'''''''''''''''''''''''''''''''''''''
		Else
			On Error Resume Next
			intMinutesWorked = CInt(DateDiff("n",datIn1,datOut1)) + CInt(DateDiff("n",datIn2,datOut2))			
		End if
		intMinutesWorked = CInt(intMinutesWorked)
		if intMinutesWorked < 0 then
			intMinutesWorked = CInt(1440 + intMinutesWorked)
		end if
		'''''''''''''''''''''''''''''''''''''
		'get hours from minutes
		'''''''''''''''''''''''''''''''''''''
		intHoursWorked = 0
		do while intMinutesWorked > 60
			intHoursWorked = intHoursWorked + 1
			intMinutesWorked = intMinutesWorked - 60
		loop
		intMinutesWorked = CDbl(intMinutesWorked/60)		
		'''''''''''''''''''''''''''''''''''''
		'prepare intTimeWorked for return value
		'''''''''''''''''''''''''''''''''''''
		intTimeWorked = CDbl(intHoursWorked + intMinutesWorked)
		intTimeWorked = formatNumber(intTimeWorked, 2)
		if CDate(datDate) <> Date() Then
			tallyHours = intTimeWorked
		else
			tallyHours = "<font color=red>" & dictLanguage("Workday_In_Progress") & "</font>"
		end if
	end if
End Function
%>

<!--#include file="../includes/main_page_close.asp"-->
