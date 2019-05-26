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
nDate = request("date")
if nDate = "" or not isDate(nDate) then
	nDate = date()
end if 

nDate = datevalue(nDate)
nDateNext = dateadd("d",1,nDate)
nDatePrev = dateadd("d",-1,nDate)
nDateNextShow = CDate(MonthName(month(nDateNext)) & " " & day(nDateNext) & ", " & year(nDateNext))
nDatePrevShow = CDate(MonthName(month(nDatePrev)) & " " & day(nDatePrev) & ", " & year(nDatePrev)) %>

<!--#include file="../includes/popup_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border="0" width="100%" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td class="homeheader" colspan="2" align="center"><%=dictLanguage("Calendar")%>: <%=nDate%></td></tr> 
	<tr>
		<td nowrap><a class="small" href="default.asp?date=<%=nDatePrev%>"><< <%=dictLanguage("Previous_Day")%> (<%=nDatePrevShow%>)</a></td>
		<td width="100%" align="right" nowrap><a class="small" href="default.asp?date=<%=nDateNext%>"><%=dictLanguage("Next_Day")%> (<%=nDateNextShow%>) >></a></td>	</tr>
</table>	
<br>
<table cellpadding="2" cellspacing="2" border="0" width="100%" align="center">
<%
sql = sql_GetEventsByDate(nDate)
Call RunSQL(sql, rs)
if not rs.eof then 
	while not rs.eof %>

	<tr>
		<td nowrap valign="top"><a href="javascript: popup2('<%=gsSiteRoot%>calendar/event.asp?id=<%=rs("id")%>');" class="bolddark"><%=timevalue(rs("calendar_startTime"))%></a></td>
		<td nowrap><a href="javascript: popup2('<%=gsSiteRoot%>calendar/event.asp?id=<%=rs("id")%>');" class="bolddark"><%=rs("calendar_heading")%></a></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td><%=rs("calendar_abstract")%>&nbsp;<a href="javascript: popup2('<%=gsSiteRoot%>calendar/event.asp?id=<%=rs("id")%>');" class="small"><%=dictLanguage("more")%> >></a></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
<%			rs.movenext 
		wend
else %>
	<tr><td colspan="2"><%=dictLanguage("No_Events_Scheduled")%><%=nDate%>.</td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
<%
end if
rs.close
set rs = nothing 

if gsPTO then
	sql = sql_GetPTOByDate(nDate)
	Call RunSQL(sql, rsPTO)
	while not rsPTO.eof %>
	<tr>
		<td nowrap valign="top"><%=dictLanguage("PTO")%>:</td>
		<td width="100%"><%=rsPTO("employeename")%><br><%=rsPTO("start_date")%>&nbsp;<%=timevalue(rsPTO("start_time"))%> to <%=rsPTO("end_date")%>&nbsp;<%=timevalue(rsPTO("end_time"))%></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%		rsPTO.movenext
	wend
	rsPTO.close
	set rsPTO = nothing
end if	

if gsCalendarClients then
	sql = sql_GetClientEventsByDate(nDate)
	Call RunSQL(sql, rs)
	while not rs.eof %>
	<tr>
		<td nowrap valign="top"><a href="javascript: opener.location='<%=gsSiteRoot%>clients/client-edit.asp?client_id=<%=rs("client_id")%>'; window.location.reload();" class="bolddark"><%=rs("client_since")%></a>&nbsp;&nbsp;</td>
		<td width="100%"><a href="javascript: opener.location='<%=gsSiteRoot%>clients/client-edit.asp?client_id=<%=rs("client_id")%>'; window.location.reload();" class="bolddark">New Client: <%=rs("client_name")%></a></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
<%		rs.movenext
	wend
	rs.close
	set rs = nothing
end if 

if gsCalendarProjects then
	sql = sql_GetProjectEventsByDate(nDate)
	Call RunSQL(sql, rs)
	while not rs.eof 
		if datevalue(rs("start_date")) = datevalue(nDate) then %>
	<tr>
		<td nowrap valign="top"><a href="javascript: opener.location='<%=gsSiteRoot%>projects/project-edit.asp?project_id=<%=rs("project_id")%>'; window.location.reload();" class="bolddark"><%=rs("start_date")%></a>&nbsp;&nbsp;</td>
		<td width="100%"><a href="javascript: opener.location='<%=gsSiteRoot%>projects/project-edit.asp?project_id=<%=rs("project_id")%>'; window.location.reload();" class="bolddark">New Project Started for <%=rs("client_name")%>: <%=rs("description")%></a></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
<%		end if
		if datevalue(rs("launch_date")) = datevalue(nDate) then %>
	<tr>
		<td nowrap valign="top"><a href="javascript: opener.location='<%=gsSiteRoot%>projects/project-edit.asp?project_id=<%=rs("project_id")%>'; window.location.reload();" class="bolddark"><%=rs("launch_date")%></a>&nbsp;&nbsp;</td>
		<td width="100%"><a href="javascript: opener.location='<%=gsSiteRoot%>projects/project-edit.asp?project_id=<%=rs("project_id")%>'; window.location.reload();" class="bolddark">Project Scheduled for Completion: <%=rs("client_name")%>, <%=rs("description")%></a></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>		
<%		end if
		rs.movenext
	wend
	rs.close
	set rs = nothing
end if 

if gsCalendarTasks then
	sql = sql_GetTaskEventsByDate(nDate, session("employee_id"))
	Call RunSQL(sql, rs)
	while not rs.eof 
		if rs("assignedto") = session("employee_id") then %>
	<tr>
		<td nowrap valign="top"><a href="javascript: opener.location='<%=gsSiteRoot%>tasks/view_mine.asp'; window.location.reload();" class="bolddark"><%=rs("datedue")%></a>&nbsp;&nbsp;</td>
		<td width="100%"><a href="javascript: opener.location='<%=gsSiteRoot%>tasks/view_mine.asp'; window.location.reload();" class="bolddark">Incomplete Task due for you: <%=rs("client_name")%>, <%=rs("project_name")%></a><br><%=rs("task_desc")%></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
<%		elseif rs("orderedby") = session("employee_id") then %>
	<tr>
		<td nowrap valign="top"><a href="javascript: opener.location='<%=gsSiteRoot%>tasks/view_my_assigned.asp'; window.location.reload();" class="bolddark"><%=rs("datedue")%></a>&nbsp;&nbsp;</td>
		<td width="100%"><a href="javascript: opener.location='<%=gsSiteRoot%>tasks/view_my_assigned.asp'; window.location.reload();" class="bolddark">Incomplete Task due for <%=rs("empName")%>: <%=rs("client_name")%>, <%=rs("project_name")%></a><br><%=rs("task_desc")%></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>		
<%		end if
		rs.movenext
	wend
	rs.close
	set rs = nothing
end if %>
	
</table>
<p>&nbsp;</p>

<!--#include file="../includes/popup_page_close.asp"-->
