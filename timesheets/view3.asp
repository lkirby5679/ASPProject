<%@ LANGUAGE="VBSCRIPT" %>
<%Server.ScriptTimeout = 1200%>
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
empID = request("id")
empName = request("name")
prevFirstWeek = Request("PreviousFirstOfWeek")
%>

<table border="0" cellpadding="2" cellspacing="2" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td align="center" class="homeheader"><%=dictLanguage("Timesheet_History")%>--<%=empName%></td></tr>
	<tr>
		<td>
			
<%
''''''''''''''''''''''''''''
'This Page Displays Hourly Employeee Time Sheet History Info
''''''''''''''''''''''''''''
'Initialize datPreviousFirstOfPeriod, passed to GetDateRange()
'and boolRecordToDisplay
''''''''''''''''''''''''''''
datPreviousFirstOfPeriod = CDate(prevFirstWeek)
boolRecordToDisplay = True
sql = sql_GetFirstTimesheetForEmployeeID(empID)
'Response.Write sql & "<BR>"
Call RunSQL(sql, rs)
if not rs.eof then
	datFirstRecord = CDate(rs("DateHoursLogged"))
	'Response.Write datFirstRecord & "<BR>"
else
	boolRecordToDisplay = FALSE
end if
rs.close
set rs = nothing
if boolRecordToDisplay then
	sql = sql_GetPayPeriodsByDate(datFirstRecord)
	Call RunSQL(sql, rs)
	if not rs.eof then
		datFirstRecord = CDate(rs("startdate"))
		'Response.Write datFirstRecord & "<BR>"
	else 
		sql = sql_GetFirstPayPeriod()
		Call RunSQL(sql, rs2)
		if not rs2.eof then
			datFirstRecord = CDate(rs2("startdate"))
			'Response.Write datFirstRecord & "<BR>"
		else
			boolRecordToDisplay = FALSE
		end if
		rs2.close
		set rs2 = nothing
	end if
	rs.close
	set rs = nothing
end if	

if boolRecordToDisplay then
	sql = sql_GetPayPeriodsBetweenStartDateEndDate(datFirstRecord, datPreviousFirstOfPeriod)
	'Response.Write sql & "<BR>"
	Call RunSQL(sql, rs)
	if not rs.eof then
		while not rs.eof 
			'Response.Write dictLanguage("Timesheet_For_Period") & "&nbsp;" & rs("startdate") & " - " & rs("enddate") & "<BR>"
			boolRecordFound = GetBoolRecordToDisplay(CDate(rs("startdate")), CDate(rs("enddate")), empID)
			intTotalHoursWorkedForPeriod = 0
			if boolRecordFound then %>
				<table border="0" cellpadding="2" cellspacing="2" width="600" align="center">
					<tr><td colspan=6 class="homeheader"><%=dictLanguage("Timesheet_For_Period")%>&nbsp;<%=CDate(rs("startdate"))%>-<%=CDate(rs("enddate"))%></td></tr>
					<tr bgcolor="<%=gsColorHighlight%>">
						<td class="columnheader"><%=dictLanguage("Date")%></td>
						<td class="columnheader"><%=dictLanguage("In_1")%></td>
						<td class="columnheader"><%=dictLanguage("Out_1")%></td>
						<td class="columnheader"><%=dictLanguage("In_2")%></td>
						<td class="columnheader"><%=dictLanguage("Out_2")%></td>
						<td class="columnheader"><%=dictLanguage("Total_Hours")%></td>
					</tr>
<%
			'''''''''''''''''''''''''
			'Display Daily Hours Worked Record for Given Date Range
			'''''''''''''''''''''''''
				sql = sql_GetTimesheetsByEmployeeIDStartDateEndDate( _
					empID, _
					CDate(rs("startdate")), _
					CDate(rs("enddate")))
				set adoRS = Conn.execute(sql)
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
					end if  %>
						<TR>
							<td class="small"><%=day1%></td>
							<td class="small"><%=in1%></td>
							<td class="small"><%=out1%></td>
							<td class="small"><%=in2%></td>
							<td class="small"><%=out2%></td>
							<td class="small"><a href="edit.asp?id=<%=adoRS("id")%>&empID=<%=empID%>"><%=varDisplayTimeWorked%></a></td>
						</tr>
<%					if IsNumeric(varDisplayTimeWorked) Then
						intTotalHoursWorkedForPeriod = CDbl(intTotalHoursWorkedForPeriod) + CDbl(varDisplayTimeWorked)
					else
						boolDisplayExplanation = True
					end if
					adoRS.movenext
				loop
				adoRS.close 
				set adoRS = nothing %>
				<tr>
					<td colspan="4">&nbsp;</td>
					<td><b class="bolddark"><%=dictLanguage("Total")%></b></td>
					<td><b><%=intTotalHoursWorkedForPeriod%></b></td>
				</tr>
			</table>
<%			else %>
			<br><%=dictLanguage("No_Timesheets_Archived")%>&nbsp;<%=empName%>.<%
			end if
			rs.movenext
		wend
	else
		Response.Write dictLanguage("No_Timesheets_Archived") & "&nbsp;" & empName & "<BR>" 
	end if
	rs.close 
	set rs = nothing
else %>
	<br><%=dictLanguage("No_Timesheets_Archived")%>&nbsp;<%=empName%>.<%
end if%>
		</td>
	</tr>
</table>

<%''''''''''''''''''''''''''''''''''
'If there was a bad record message,give full explanation here:
''''''''''''''''''''''''''''''''''
if boolDisplayExplanation = True then %>
<font class="alert"><%=dictLanguage("Timesheet_Inst_1")%></font>
<%end if%>
<br>
<p align="center">
<a href="view.asp"><%=dictLanguage("Hourly_Timesheets")%></a><br>
<a href="main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>
<%
''''''''''''''''''''''''''''
'This function returns True if there is another week's worth 
'of time sheet history to display.  
'It is passed the date range created by GetDateRange (ArrayDateRange)
''''''''''''''''''
'This function depends on the Conn set earlier in the page
'and on session("userid")
''''''''''''''''''''''''''''
Function GetBoolRecordToDisplay(startdate, enddate, empID)
	datLastOfWeek = CDate(enddate)
	sql = sql_GetTimesheetsByEmployeeIDEndDate(empID,datLastOfWeek)	
	Call RunSQL(sql, adoRS)
	if adoRS.eof then
		BoolRecordToDisplay = False
	else
		BoolRecordToDisplay = True
	end if
	adoRS.close
	GetBoolRecordToDisplay = BoolRecordToDisplay
End Function

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
			intMinutesWorked = CInt(DateDiff("n",datIn1,datOut1))
			On Error Goto 0
		'''''''''''''''''''''''''''''''''''''
		'If they did take a lunch break
		'''''''''''''''''''''''''''''''''''''
		Else
			On Error Resume Next
			intMinutesWorked = CInt(DateDiff("n",datIn1,datOut1)) + CInt(DateDiff("n",datIn2,datOut2))
			On Error Goto 0
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
