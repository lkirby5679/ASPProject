<%@ LANGUAGE="VBSCRIPT" %>
<%Server.ScriptTimeout = "600"%>
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

<%	sql = sql_GetActiveEmployees()
	Call RunsQL(sql, rs) %>

<!--#include file="../includes/main_page_open.asp"-->

<table border="0" cellspacing="2" cellpadding="2" width="100%">
	<tr><td colspan="4" bgcolor="<%=gsColorHighlight%>" class="homeheader" align="center"><%=dictLanguage("Task_Hours_Assigned")%></td></tr>
	<tr>
		<td colspan="4"><%=dictLanguage("Task_Inst_1")%></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td bgcolor="<%=gsColorHighlight%>" class="columnHeader">&nbsp;</td>
		<td bgcolor="<%=gsColorHighlight%>" class="columnHeader"><%=WeekdayName(weekday(date()))%> - <%=WeekdayName(6)%></td>
		<td bgcolor="<%=gsColorHighlight%>" class="columnHeader"><%=dictLanguage("Next_Week")%></td>
		<td bgcolor="<%=gsColorHighlight%>" class="columnHeader"><%=dictLanguage("Total")%></td>
<%	i=1	
	bgcolor = gsColorWhite
	while not rs.eof %>
	<tr bgcolor="<%=bgcolor%>">
		<td class="bolddark"><%=rs("EmployeeName")%></b></td>
		<td>
<%		sql = sql_GetTaskHoursByEmployee(rs("Employee_ID"))
		'response.write(sql)
		Call RunSQL(sql, rsHours)
		if not rsHours.eof then
			if rsHours("curHours")<>"" then
				estHours = rsHours("curHours")
			else
				estHours = 0
			end if
		else
			estHours = 0
		end if
		rsHours.close
		set rsHours = nothing
		sql = sql_GetTaskHoursCompleteByEmployee(rs("Employee_ID"))
		Call RunSQL(sql, rsDoneHours)
		if not rsDoneHours.eof then
			if rsDoneHours("TimeCharged")<>"" then
				timeCharged = rsDoneHours("TimeCharged")
			else
				timeCharged = 0
			end if
		else
			timeCharged = 0
		end if
		rsDoneHours.close
		set rsDoneHours = nothing		
			
		estHoursLeft = estHours - timeCharged
		weekHours = ((6 - Weekday(date) + 1) * 8)
		if estHoursLeft>=0 then
			while (i <= estHoursLeft) and (i <= weekHours)
				Response.Write "<img src='" & gsSiteRoot & "images/chartblock.gif' width='10' height='10' border='0'>"
				i = i + 1
			wend
			if estHoursLeft <>"" then
				if weekHours < estHoursLeft then
					Response.Write weekHours
				else
					Response.write estHoursLeft
				end if
			else
				Response.Write dictLanguage("No_Tasks_Assigned")
			end if
			i = 1
		else
			if estHoursLeft < 0 then
				Response.Write "More time used on tasks than allocated"
			else
				Response.Write dictLanguage("No_Tasks_Assigned")
			end if
		end if %>
		</td>
		<td>
<%		i=weekHours
		if estHoursLeft-weekHours >0 then
			while (i <= estHoursLeft)
				Response.Write "<img src='" & gsSiteRoot & "images/chartblock.gif' width='10' height='10' border='0'>"
				i = i + 1
			wend
			if estHoursLeft <>"" then
				if estHoursLeft<0 then
					Response.Write dictLanguage("More_Time_Used")
				end if
				Response.Write estHoursLeft-weekHours
			else
				Response.Write "0"
			end if
			i = 1
		else	
			i = 1
			Response.Write "&nbsp;"
		end if%>
		</td>
		<td><%=estHoursLeft%>&nbsp;</td>
	</tr>
<%		rs.Movenext
		if bgcolor="#FFFFFF" then 
			bgcolor = gsColorWhite 
		else 
			bgcolor = "#FFFFFF" 
		end if
	wend %>
</table>

<p align="center">
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->

