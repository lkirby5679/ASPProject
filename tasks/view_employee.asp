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

<%
active = Request.form("Submit")
if active = "Archived" then
	active = 0
	strShowActive = dictLanguage("Archived")
else
	active = 1
	strShowActive = dictLanguage("Active")
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'validate form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if request.form("employee") = "" and request.querystring("sql") = "" then
	session("Msg") = dictLanguage("Error_TaskNoEmployee")
	response.redirect "view.asp"
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'bring back all of this employee's tasks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if request.querystring("sql") = "" then
	sql = sql_GetTaskViewByEmployeeID(active, request.form("employee"))
else
	sql = request.querystring("sql")
end if
Call RunSQL(sql, rsTasks)
if rsTasks.EOF then
	session("Msg") = "<b>" & dictLanguage("Error_TaskNoTasks") & "</b><BR><br>"
	response.redirect "view.asp"
else
	strName = rsTasks("EmployeeNameAssignedTo")
%>

<!--#include file="../includes/main_page_open.asp"-->

<table border="0" cellpadding="2" cellspacing="2" width="100%">
	<tr>
		<td valign=top colspan=11>
			<b class="homeheader"><%=strShowActive%>&nbsp;<%=dictLanguage("Tasks_List_Report")%>--<%=strName%></b><br>
			<%=now()%><br>
		</td>
	</tr>
</table>

<table width="100%" border=0 cellpadding="2" cellspacing="2">	
	<tr bgcolor="<%=gsColorHighlight%>">
		<td valign=top class="columnheader" nowrap>&nbsp;</td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Client/Project")%></td>		
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Done")%>?</td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Date_Due")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Priority")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Created_By")%></td>
		<td valign=top width="100%" class="columnheader"><%=dictLanguage("Description")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Est_Total_Hours")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Est_Hours_Left")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Date_Created")%></td>
		<td valign=top class="columnheader" nowrap>&nbsp;</td>
	</tr>
<%	do while not rsTasks.EOF
		boolDone = rsTasks("Done") %>	
	<tr <%if boolDone then%>bgcolor="#EEEEEE"<%else%>bgcolor="#FFFF77"<%end if%>>
		<td valign="top" class="small"><a href="edit.asp?id=<%=rsTasks("Task_ID")%>"><%=dictLanguage("Edit")%></a></td>
		<td valign="top" class="small"><%=rsTasks("Client_Name")%> - <%=rsTasks("ProjectName")%></td>
		<td valign="top" class="small" align="center"><a href="toggle_done.asp?id=<%=rsTasks("Task_ID")%>&value=
<%			if rsTasks("Done") = 0 then
				response.write "no"
			else
				response.write "yes"
			end if%>&referring_page=view_employee.asp&active=<%=active%>&sql=<%=Server.URLEncode(strTasksSQL)%>">
<%			if boolDone = True then
				response.write "<img src='" & gsSiteRoot & "images/done.gif' border=0>"
			else
				response.write "<img src='" & gsSiteRoot & "images/notDone.gif' border=0>"
			end if%></a></td>
		<td valign=top class="small"><%=rsTasks("DateDue")%></td>
		<td valign=top class="small">
<%			select case rsTasks("Priority")
				case 1
					response.write "<font color=Black><b>" & dictLanguage("Low") & "</b></font>"
				case 2
					response.write "<font color=Blue><b>" & dictLanguage("Medium") & "</b></font>"
				case 3
					response.write "<font color=red><b>" & dictLanguage("High") & "</b></font>"
				case 0 
					response.write "<font color=green><b>" & dictLanguage("On_Hold") & "</b></font>"
				case else
					response.write "<font color=Blue></b>" & dictLanguage("Medium") & "</b></font>"
			end select%></td>
		<td valign=top class="small"><%=rsTasks("EmployeeNameOrderedBy")%></td>
		<td valign=top class="small"><%=rsTasks("Description")%></td>
		<td valign=top class="small"><%=rsTasks("EstimatedHours")%></td>
<%			sqlHours =	sql_GetTimecardHoursByTaskID(rsTasks("Task_ID"))
			Call RunSQL(sqlHours, rsHours)
			if rsHours("TimeCharged")<>"" then
				timeCharged = rsHours("TimeCharged")
			else 
				timeCharged = 0
			end if
			estHoursLeft = rsTasks("EstimatedHours") - timeCharged
			rsHours.close
			set rsHours = nothing %>
		<td valign=top class="small"><font class="small"<%if estHoursLeft<0 then%> Color=red <%end if%>><%=estHoursLeft%></td>
		<td valign=top class="small"><%=rsTasks("DateDue")%></td>
		<td valign=top class="small">
<%			if (active = "1" or active="-1" or active) and (session("permTasksDelete") or session("employee_id")=rsTasks("EmployeeNameOrderedBy")) then %>			
		<a href="delete.asp?id=<%=rsTasks("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Delete")%>');"><%=dictLanguage("Delete")%></a>
<%			elseif (session("permTasksDelete") or session("employee_id")=rsTasks("EmployeeNameOrderedBy")) then %>
		<a href="undelete.asp?id=<%=rsTasks("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Restore")%>');"><%=dictLanguage("Undelete")%></a>
<%			else %>
		&nbsp;
<%			end if %>
		</td>
	</tr>
<%		rsTasks.MoveNext		
	loop
	rsTasks.close
	set rsTasks = nothing %>
</table>	

<p align="center">
<a href="view.asp"><%=dictLanguage("View_Tasks")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<%
end if%>

<!--#include file="../includes/main_page_close.asp"-->