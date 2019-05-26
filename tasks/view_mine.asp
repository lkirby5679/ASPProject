<%@ LANGUAGE="VBSCRIPT"%>
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

<%
active = request("active")
if active = "0" then
	active = 0
	strShowActive = dictLanguage("Archived")
else
	active = 1
	strShowActive = dictLanguage("Active")
end if
for each i in request.form
if InStr(request.form(i),"~") then
	temp=split(request.form(i),"~")
	session(i)=temp(0)
	session(i & "Desc")=temp(1)
else
	session(i)=request.form(i)
	session(i & "Desc")=request.form(i)
end if
next
%>

<!--#include file="../includes/main_page_open.asp"-->

<table border="0" cellpadding="2" cellspacing="2" width="100%">
	<tr>
		<td valign=top colspan=11>
			<b class="homeheader"><%=strShowActive%>&nbsp;<%=dictLanguage("Tasks_List_Report")%>--<%=dictLanguage("My_Tasks")%></b><br>
			<%=now()%><br>
		</td>
	</tr>
</table>

<%

strEmpID = session("employee_id")
if strEmpID = "" then
	Response.Redirect gsSiteRoot & "default.asp"
end if

sql = sql_GetTaskViewByEmployeeID(active, strEmpID)
'response.write sql
Call RunSQL(sql, rs)%>

<table border=0 width="100%" cellpadding="2" cellspacing="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td valign=top nowrap class="columnheader">&nbsp;</td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Client/Project")%></td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Done")%>?</td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Date_Due")%></td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Priority")%></td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Created_By")%></td>
		<td valign=top width="100%" class="columnheader"><%=dictLanguage("Description")%></td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Est_Total_Hours")%></td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Est_Hours_Left")%></td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Date_Created")%></td>
		<td valign=top nowrap class="columnheader"><%=dictLanguage("Create_Timecard")%></td>
	</tr>

<%do while not rs.eof
	boolDone = rs("done")
	New_Client_Name = rs("Client_Name") %>

	<tr <%if boolDone then%>bgcolor="#EEEEEE"<%else%>bgcolor="#FFFF77"<%end if%>>
		<td valign="top" class="small"><a href="edit.asp?id=<%=rs("Task_ID")%>"><%=dictLanguage("Edit")%></a></td>
		<td valign="top" class="small"><%=rs("Client_Name")%> - <%=rs("ProjectName")%></td>
		<td valign="top" class="small" align="center"><a href="toggle_done.asp?id=<%=rs("Task_ID")%>&value=
<%	if rs("Done") = 0 then
		response.write "no"
	else
		response.write "yes"
	end if%>&referring_page=view_mine.asp&active=<%=active%>">
<%	if boolDone = True then
		response.write "<img src='" & gsSiteRoot & "images/done.gif' border=0>"
	else
		response.write "<img src='" & gsSiteRoot & "images/notDone.gif' border=0>"
	end if%></a></td>
		<td valign=top class="small"><%=rs("DateDue")%></td>
		<td valign=top class="small"> 
<%	select case rs("Priority")
		case 1
			response.write dictLanguage("Low")
		case 2
			response.write "<font color=Blue><b>" & dictLanguage("Medium") & "</b></font>"
		case 3
			response.write "<font color=red><blink><strong>" & dictLanguage("High") & "</strong></blink></font>"
		case 0 
			response.write "<font color=green><b>" & dictLanguage("On_Hold") & "</b></font>"
		case else
			response.write "<font color=blue><b>" & dictLanguage("Medium") & "</b></font>"
	end select%></td>
		<td valign=top class="small"><%=rs("EmployeeNameOrderedBy")%></td>
		<td valign=top class="small"><%=rs("Description")%></td>
		<td valign=top class="small"><%=rs("EstimatedHours")%></td>
<%	sqlHours =	sql_GetTimecardHoursByTaskID(rs("Task_ID"))
	Call RunSQL(sqlHours, rsHours)
	if rsHours("TimeCharged")<>"" then
		timeCharged = rsHours("TimeCharged")
	else 
		timeCharged = 0
	end if
	estHoursLeft = rs("EstimatedHours") - timeCharged
	rsHours.close
	set rsHours = nothing %>
		<td valign=top class="small"><font class="small"<%if estHoursLeft<0 then%> Color=red <%end if%>><%=estHoursLeft%></td>
		<td valign=top class="small"><%=rs("DateCreated")%></td>
<%	urlDesc = Server.URLEncode(rs("Description")) %>
		<td valign=top class="small" align="center">
			<a href="<%=gsSiteRoot%>timecards/timecard.asp?desc=<%=urlDesc%>&from=tasks&project=<%=rs("project")%>&client=<%=rs("client")%>&task=<%=rs("Task_ID")%>">
<%	if boolDone = True then
		response.write dictLanguage("Do_It")
	else
		response.write "<img src='" & gsSiteRoot & "images/clock.GIF' border=0>"
	end if%></a></td>
	</tr>
<%
	Old_Client_Name = New_Client_Name
	rs.movenext
loop
rs.close
set rs = nothing %>
</table>

<p align="center">
<a href="view.asp"><%=dictLanguage("View_Tasks")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->
