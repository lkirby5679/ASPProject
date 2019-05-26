<% Server.ScriptTimeout = 600 
Response.Contenttype="application/vnd.ms-excel"
Response.AddHeader "content-disposition","attachment;filename=test.xls"
%>	

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
	
<!--#include file="../includes/SiteConfig.asp"-->
<!--#include file="../includes/connection_open.asp" -->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<%
monthlyTotalProductionGoal=160

datToday = date()
disyear = Year(datToday)
strWeekDayName = WeekDayName(WeekDay(datToday), False, vbSunday)
ThisMon = GetCurrentMonday(datToday)
ThisSun = DateAdd("d",6,ThisMon)
PrevMon = DateAdd("d",-7,ThisMon)
PrevMonth = DateAdd("m",-1,datToday)
PrevMonthFirst = CDate(MonthName(Month(PrevMonth)) & " 1, " & Year(PrevMonth))
ThisMonthFirst = CDate(MonthName(Month(datToday)) & " 1, " & Year(datToday))
%>

<HTML>
<HEAD>
	<TITLE><%=dictLanguage("Production_Report_Grid")%></TITLE>
</HEAD>

<BODY BGCOLOR="#FFFFFF" LINK="#FFFFFF" TEXT="#FFFFFF" VLINK="#FFFFFF" ALINK="#FFFFFF">
	<table border=0 cellpadding=2 cellspacing=2 width="100%" bgcolor="<%=gsColorBackground%>">
		<tr>
			<td><font face="Verdana,Arial,Helvetica" size=3 color="<%=gsColorHighlight%>"><b><%=dictLanguage("Client_Hours")%></b></font></td>			
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("Today")%></font></td>
<%
'--------------------------------------------------------
' if today is Monday then show Fri/Sat/Sun else show Yday
'--------------------------------------------------------
if strWeekDayName = "Monday"  then%>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("Fri/Sat/Sun")%></font></td>
<%else%>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("Y'day")%></font></td>
<%end if%>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("This_Week")%></font></td>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("This_Month")%></font></td>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("Goal")%></font></td>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("diff")%> %</font></td>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("Last_Week")%></font></td>
			<td align=right bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1><%=dictLanguage("Last_Month")%></font></td>
		</tr>
		<tr>
			<td>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%=dictLanguage("Total")%></b>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get today's total client hours
''''''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetTeamDailyBillableHours(datToday)
'Response.Write sql & "<BR>"
Call RunSQL(sql,rs)
tot = 0
if not rs.eof then
	if rs("WorkToday")<>"" then
		tot = rs("WorkToday")
	end if
end if
rs.close
set rs = nothing
response.write FormatNumber(tot,2) %></b>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
'''''''''''''''''''''''''''''''''''''''''''''''''''
' Get yesterday's total client hours (unless it is monday
' then get the weekend's client hours)
'''''''''''''''''''''''''''''''''''''''''''''''''''
if strWeekDayName = "Monday"  then
	sql = sql_GetTeamWeekendBillableHours(datToday)
else
	sql = sql_GetTeamDailyBillableHours(dateAdd("d",-1,datToday))
end if
'Response.Write sql & "<BR>"
Call RunSQL(sql,rs)
tot=0
if not rs.eof then
	if rs("WorkToday")<>"" then
		tot = rs("WorkToday")
	end if
end if
rs.close
set rs = nothing
response.write FormatNumber(tot,2) %></b>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
'''''''''''''''''''''''''''''''''''''''''''''''''''
' Get this weeks total client hours
'''''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetTeamWeeklyBillableHours(ThisMon, ThisSun)
Call RunSQL(sql,rs)
tot=0
if not rs.eof then
	if rs("WorkThisWeek")<>"" then
		tot = rs("WorkThisWeek")
	end if
end if
rs.close
set rs = nothing 
response.write FormatNumber(tot,2) %></b>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
''''''''''''''''''''''''''''''''''''''''''''''''
' Get this month's total client hours
''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetTeamMonthlyBillableHours(datToday)
Call RunSQL(sql, rs)
tot=0
if not rs.eof then
	if rs("WorkThisMonth")<>"" then
		tot = rs("WorkThisMonth")
	end if
end if
rs.close
set rs = nothing
WorkThisMonth = FormatNumber(tot,2)
response.write WorkThisMonth %></b>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
'''''''''''''''''''''''''''''''''''''''''''''
' Get the total production goal for the month
'''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetTeamMonthlyProductionGoal()
Call RunSQL(sql,rs)
if rs.eof then
	response.write "0"
	ProductionGoal = 0
else
	ProductionGoal = rs("SumOfproductionGoal")
	response.write FormatNumber(ProductionGoal,2)
end if
rs.close
set rs = nothing %></b>
				</font>
			</td>
			<td align=right>
<%
''''''''''''''''''''''''''''''''''''''''''''''''''
' show the percentage of the total production goal that has
' been met
''''''''''''''''''''''''''''''''''''''''''''''''''
%>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
if ProductionGoal > 0 then
	response.write round((WorkThisMonth/ProductionGoal)*100,0)
else
	response.write "0"
end if
					%></b>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Show last weeks total client hours
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetTeamWeeklyBillableHours(prevMon, dateAdd("d",-1,ThisMon))
Call runSQL(sql,rs)
tot=0
if not rs.eof then
	if rs("WorkThisWeek")<>"" then
		tot = rs("WorkThisWeek")
	end if
end if
rs.close
set rs = nothing 
WorkLastWeek = FormatNumber(tot,2)
response.write WorkLastWeek %></b>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=2>
					<b><%
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' show last months total client hours
''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetTeamMonthlyBillableHours(PrevMonthFirst)
Call runSQL(sql,rs)
tot=0
if not rs.eof then
	if rs("WorkThisMonth")<>"" then
		tot = rs("WorkThisMonth")
	end if
end if
rs.close
set rs = nothing
WorkLastMonth = FormatNumber(tot,2)
response.write WorkLastMonth %></b>
				</font>
			</td>
		</tr>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''
' get list of all employees that have a production goal
'''''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetActiveEmployeesPRGView()
Call RunSQL(sql,rs)
if not rs.eof then
	dept = rs("department")
%>
	<tr><td colspan="9" bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><%=ucase(dept)%></b></font></td></tr>
<%
end if
do while not rs.eof
	employeeID = rs("employee_id")
	if dept <> rs("department") then
		dept = rs("department")
%>
	<tr>
		<td><font face="Verdana,Arial,Helvetica" size=1>&nbsp;</font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(todayTotal,2)%></b></font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(yesterdayTotal,2)%></b></font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(thisweekTotal,2)%></b></font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(thismonthTotal,2)%></b></font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(goalTotal,2)%></b></font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=Int((thismonthTotal/goalTotal*100))%></b></font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(lastweekTotal,2)%></b></font></td>
		<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(lastmonthTotal,2)%></b></font></td>
	</tr>
	<tr><td colspan="9" bgcolor="<%=gsColorHighlight%>"><font face="Verdana,Arial,Helvetica" size=1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><%=ucase(dept)%></b></font></td></tr> 
<%
		todayTotal = 0
		yesterdayTotal = 0
		thisweekTotal = 0
		thismonthTotal = 0
		goalTotal = 0
		lastweekTotal = 0
		lastmonthTotal = 0
	end if  %>
	<tr>
		<td><font face="Verdana,Arial,Helvetica" size=1><%=rs("employeeName")%></font></td>
		<td align=right>
				<font face="Verdana,Arial,Helvetica" size=1> <%
''''''''''''''''''''''''''''''''''''''''''''''
' Show client hours for today
''''''''''''''''''''''''''''''''''''''''''''''
	sql = sql_GetEmployeeDailyBillableHours(employeeID, datToday)
	Call RunSQL(sql,rs2)
	tot=0
	if not rs2.eof then
		if rs2("WorkToday")<>"" then
			tot = rs2("WorkToday")
		end if
	end if
	rs2.close	
	set rs2 = nothing	
	todayTotal = todayTotal + tot
	WorkThisWeek=FormatNumber(tot,2)
	response.write WorkThisWeek %>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=1>
<%
'--------------------------------------------------------
' if today is Monday then show Fri/Sat/Sun else show Yday
'--------------------------------------------------------
	if strWeekDayName = "Monday" then
		sql = sql_GetEmployeeWeekendBillableHours(employeeID, datToday)
	else
		sql = sql_GetEmployeeDailyBillableHours(employeeID, dateAdd("d",-1,datToday))
	end if

	Call RunSQL(sql,rs2)
	tot=0
	if not rs2.eof then
		if rs2("WorkToday")<>"" then
			tot = rs2("WorkToday")
		end if
	end if
	rs2.close
	set rs2 = nothing 
	yesterdayTotal = yesterdayTotal + tot
	WorkThisWeek = FormatNumber(tot,2)
	response.write WorkThisWeek %>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=1>
					<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' show this weeks client hours
''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql = sql_GetEmployeeWeeklyBillableHours(employeeID, ThisMon, ThisSun)
	Call RunSQL(sql, rs2)
	tot=0
	if not rs2.eof then
		if rs2("WorkThisWeek")<>"" then
			tot= rs2("WorkThisWeek")
		end if
	end if
	rs2.close
	set rs2 = nothing 	
	thisweekTotal = thisweekTotal + tot
	WorkThisWeek = FormatNumber(tot,2)
	response.write WorkThisWeek %>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=1>
					<%
''''''''''''''''''''''''''''''''''''''''''''''''''''
' show this months client hours
''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql = sql_GetEmployeeMonthlyBillableHours(employeeID, datToday)
	Call RunSQL(sql,rs2)
	tot=0
	if not rs2.eof then
		if rs2("WorkThisMonth")<>"" then
			tot = rs2("WorkThisMonth")
		end if
	end if
	rs2.close						
	set rs2 = nothing	
	thismonthTotal = thismonthTotal + tot
	WorkThisMonth = FormatNumber(tot,2)
	response.write WorkThisMonth %>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=1>
					<%=rs("productionGoal")%>
					<% goalTotal = goalTotal + rs("productionGoal") %>
				</font>
			</td>
			<td align=right>
<%'
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' show percent of production goal reached
''''''''''''''''''''''''''''''''''''''''''''''''''''''
%>
				<font face="Verdana,Arial,Helvetica" size=1>
					<%=round(WorkThisMonth/rs("productionGoal")*100,0)%>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=1>
					<%
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' show last weeks client hours
''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql = sql_GetEmployeeWeeklyBillableHours(employeeID, PrevMon, dateAdd("d",-1,ThisMon))
	Call RunSQL(sql,rs2)
	tot=0
	if not rs2.eof then
		if rs2("WorkThisWeek")<>"" then
			tot = rs2("WorkThisWeek")
		end if
	end if
	rs2.close
	set rs2 = nothing 	
	lastweekTotal = lastWeekTotal + tot
	WorkLastWeek = FormatNumber(tot,2)
	response.write WorkLastWeek %>
				</font>
			</td>
			<td align=right>
				<font face="Verdana,Arial,Helvetica" size=1>
					<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''
' show last months client hours
'''''''''''''''''''''''''''''''''''''''''''''''''''''
	PrevMonth = DateAdd("m",-1,date())
	PrevMonthFirst = CDate(MonthName(Month(PrevMonth)) & " 1, " & Year(PrevMonth))
	ThisMonthFirst = CDate(MonthName(Month(Date())) & " 1, " & Year(Date()))
	
	sql = sql_GetEmployeeMonthlyBillableHours(employeeID, PrevMonthFirst)
	Call RunSQL(sql,rs2)
	tot=0
	if not rs2.eof then
		if rs2("WorkThisMonth")<>"" then
			tot = rs2("WorkThisMonth")
		end if
	end if
	rs2.close
	set rs2 = nothing		
	lastmonthTotal = lastmonthTotal + tot
	WorkLastMonth=FormatNumber(tot,2)
	response.write WorkLastMonth %>
				</font>
			</td>
		</tr>
<%
	rs.movenext
loop
rs.close
set rs = nothing %>

<tr>
<td><font face="Verdana,Arial,Helvetica" size=1>&nbsp;</font></td>
<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(todayTotal,2)%></b></font></td>
<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(yesterdayTotal,2)%></b></font></td>
<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(thisweekTotal,2)%></b></font></td>
<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(thismonthTotal,2)%></b></font></td>
	<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(goalTotal,2)%></b></font></td>
	<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=round((thismonthTotal/goalTotal*100),0)%></b></font></td>
	<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(lastweekTotal,2)%></b></font></td>
	<td bgcolor="<%=gsColorHighlight%>" align="right"><font face="Verdana,Arial,Helvetica" size=1><b><%=formatnumber(lastmonthTotal,2)%></b></font></td>
</tr>

</table>

</BODY>
</HTML>

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''
'' This function returns the date of the monday of the
'' current week
'''''''''''''''''''''''''''''''''''''''''''''''''''
function GetCurrentMonday(datToday)
		daDate = datToday
		ctr = 1
		while ctr < 8
			if Weekdayname(Weekday(daDate)) = "Monday" then
				datMonday = daDate
			end if
			daDate = DateAdd("d",-1,daDate)
			ctr = ctr+1
		wend
		GetCurrentMonday = datMonday
end function
%>

<!--#include file="../includes/connection_close.asp" -->