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
for each i in request.form
if InStr(request.form(i),"~") then
	temp=split(request.form(i),"~")
	session(i)=temp(0)
	session(i & "Desc")=temp(1)
	strCompanyName = temp(1)
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
			<b class="homeheader"><%=dictLanguage("Tasks_List_Report")%>--<%=strCompanyName%></b><br>
			<%=now()%><br>
		</td>
	</tr>
	<tr bgcolor="<%=gsColorHighlight%>">
		<td valign=top class="columnheader">&nbsp;</td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Done")%>?</td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Date_Due")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Priority")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Created_By")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Assigned_To")%></td>
		<td valign=top class="columnheader" width="100%"><%=dictLanguage("Description")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Est_Total_Hours")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Est_Hours_Left")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Date_Created")%></td>
		<td valign=top class="columnheader">&nbsp;</td>
	</tr>
<%
	if request.querystring("sql") = "" then
		sql = sql_GetTaskViewByClientID(session("active"), session("Client_ID"))
	else
		sql = request.querystring("sql")
	end if
	Call RunSQL(sql, rs)
	while not rs.eof
		boolDone = rs("Done") %>
	<tr<%if boolDone = 0 or not boolDone then%> bgcolor="#FFFF77" <%else%> bgcolor="#EEEEEE" <%end if%>>
		<td valign="top" class="small"><a href="edit.asp?id=<%=rs("Task_ID")%>"><%=dictLanguage("Edit")%></a></td>
		<td valign="top" align="center"><a href="toggle_done.asp?id=<%=rs("Task_ID")%>&value=
<%		if rs("Done") = 0 then
			response.write "no"
		else
			response.write "yes"
		end if %>&referring_page=view2.asp&sql=<%=Server.URLEncode(sql)%>">
<%		if boolDone = True then
			response.write "<img src='" & gsSiteRoot & "images/done.gif' border=0>"
		else
			response.write "<img src='" & gsSiteRoot & "images/notDone.gif' border=0>"
		end if %>
		</td>
		<td valign=top class="small"><%=rs("DateDue")%></td>
		<td valign="top" align=center class="small"> 
<%		select case rs("Priority")
			case 1
				response.write dictLanguage("Low")
			case 2
				response.write "<font color=Blue><b>" & dictLanguage("Medium") & "</b></font>"
			case 3
				response.write "<font color=red><b>" & dictLanguage("High") & "</b></font>"
			case 0 
				response.write "<font color=green><b>" & dictLanguage("On_Hold") & "</b></font>"
			case else
				response.write "<font color=blue><b>" & dictLanguage("Medium") & "</b></font>"
		end select%>
		</td>
		<td valign=top class="small"><%=rs("EmployeeName")%></td>
		<td valign=top class="small"><%=rs("EmployeeName_1")%></td>
		<td valign=top class="small"><%=rs("Description")%></td>
		<td valign=top class="small"><%=rs("EstimatedHours")%></td>
<%
		sqlHours = sql_GetTimecardHoursByTaskID(rs("Task_ID"))
		Call RunSQL(sqlHours, rsHours)
		if rsHours("TimeCharged")<>"" then
			timeCharged = rsHours("TimeCharged")
		else 
			timeCharged = 0
		end if
		estHoursLeft = rs("EstimatedHours") - timeCharged
		rsHours.close
		set rsHours = nothing %>
		<td valign=top class="small"><font <%if estHoursLeft<0 then%>color="#FF0033" <%end if%>class="small"><%=estHoursLeft%></font></td>
		<td valign=top class="small"><%=rs("DateCreated")%></td>
		<td valign="top" class="small">
<%	if rs("show") = "1" or rs("show") = "-1" or rs("show") then %>			
			<a href="delete.asp?id=<%=rs("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Delete")%>');"><%=dictLanguage("Delete")%></a>
<%	else %>
			<a href="undelete.asp?id=<%=rs("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Restore")%>');"><%=dictLanguage("Undelete")%></a>
<%	end if %>
		</td>
	</tr>
<%		rs.movenext
	wend
	rs.close
	set rs = nothing
	session("Client_ID")="" %>
</table>
<p align="center">
<a href="view.asp"><%=dictLanguage("View_Tasks")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->
