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
project_id = request.querystring("PROJECT_ID")

if session("errorMsg") <> "" then
	response.write "<p class='alert'>" & session("errorMsg") & "</p>"
	session("errorMsg") = ""
end if
%>

<form method="POST" action="new2.asp" id=strForm name=strForm>
<table cellpadding="2" cellspacing="2" border=0 align="center">
	<tr><td colspan="2" bgcolor="<%=gsColorHighlight%>" class="homeheader" align="center"><%=dictLanguage("Task_Details")%></td></tr>
	<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Client")%>:<font class="alert">*</font></td>
		<td>
			<select name="client" class="formstyleXL">
<%	if project_id <> "" then
		sql = sql_GetProjectViewByID(project_id)
		Call RunSQL(sql, rsProjectSelect)
		if not rsProjectSelect.eof then %>
				<option value="<%=rsProjectSelect("cl_id")%>P<%=rsProjectSelect("Project_ID")%>"><%=rsProjectSelect("Client_Name")%> {<%=rsProjectSelect("Description")%>}</option>
<%		end if
		rsProjectSelect.close
		Set rsProjectSelect = nothing
	end if%>
				<option value=""> - <%=dictLanguage("Select_One_Clients_Projects")%> - </option>
<%  sql = sql_GetActiveProjectsWithClients()
	Call RunSQL(sql, rsClientSelect)
	While Not rsClientSelect.EOF
		intClientId = rsClientSelect("cl_id")
		intProjectId = rsClientSelect("Project_ID")
		strClientProject = intClientId & "P" & intProjectId %>
				<option value="<%=intClientId%>P<%=intProjectId%>" <%if session("client") = strClientProject then%> selected <%end if%>><%=rsClientSelect("Client_Name")%> {<%=rsClientSelect("Description")%>}</option>
<%		rsClientSelect.MoveNext
	Wend
	rsClientSelect.close
	Set rsClientSelect = Nothing%>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark" valign="top"><%=dictLanguage("Description")%>:<font class="alert">*</font></td>
		<td><textarea name="description" rows=4 class="formstyleLong"><%=session("description")%></textarea></td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Priority")%>:<font class="alert">*</font></td>
		<td>
			<select name="priority" class="formstyleMed">
				<option value="1" <%if session("priority")="1" then%> selected <%end if%>>Low</option>
				<option value="2" <%if session("priority")="2" then%> selected <%end if%>>Medium</option>
				<option value="3" <%if session("priority")="3" then%> selected <%end if%>>High</option>
				<option value="0" <%if session("priority")="4" then%> selected <%end if%>>On Hold</option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Created_By")%>:<font class="alert">*</font></td>
		<td>
			<select name="orderedBy" class="formstyleLong">
				<option value="">-none-</option>
<%	sql= sql_GetActiveEmployees()
	Call RunSQL(sql, rs)
	do while not rs.eof
		intEmployeeId = rs("Employee_ID")
		response.write "<option value='" & intEmployeeId & "'"
		if session("orderedBy")<>"" then
			if cint(intEmployeeID) = cint(session("orderedBy")) then
				response.write " selected "
			end if
		end if
		response.write ">" & rs("EmployeeName") & "</option>"
		rs.movenext
	loop %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Assigned_To")%>:<font class="alert">*</font></td>
		<td>
			<select name="assignedTo" class="formstyleLong">
				<option value="">-none-</option>
<%	if session("assignedTo") <> "" then%>
				<option value="<%=session("assignedTo")%>"><%=session("assignedTo")%></option>
<%	end if%>
<%	if not (rs.eof and rs.bof) then
		rs.movefirst
		do while not rs.eof
			intEmployeeId = rs("Employee_ID")
			response.write "<option value='" & intEmployeeId & "'"
			if session("assignedTo")<>"" then
				if CINT(intEmployeeId) = INT(session("assignedTo")) then
					response.write " selected "
				end if
			end if
			response.write ">" & rs("employeeName") & "</option>"
		rs.movenext
		loop
	end if
	rs.close
	set rs = nothing %>
			</select>
		</td>
	</tr>
<%	if session("estimatedHours") <> "" then
		if not isNumeric(session("estimatedHours")) then
			session("estimatedHours") = ""
		end if
	end if %>
	<tr>
		<td class="bolddark"><%=dictLanguage("Estimated_Hours")%>:<font class="alert">*</font></td>
		<td><input type="text" value="<%= session("estimatedHours") %>" class="formstyleTiny" name="estimatedHours" maxlength="8"></td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Date_Created")%>:<font class="alert">*</font></td>
<%	if session("dateCreated") <> "" then
		if not isDate(Session("dateCreated")) then
			d_created = date()	
		else
			d_created = session("dateCreated")		
		end if		
	else
		d_created = date()
	end if %>
		<td>
			<input type="Text" name="dateCreated" class="formstyleShort" value="<%=d_created%>" onkeypress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','dateCreated',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Date_Due")%>:<font class="alert">*</font></td>
<%	if session("dateDue") <> "" then
		if not isDate(Session("dateDue")) then
			d_due = date()	
		else
			d_due = session("dateDue")		
		end if	
	else
		d_due = date()
	end if %>
		<td>
			<input type="text" name="dateDue" value="<%=d_due%>" class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','dateDue',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<%if session("permTasksAdd") then%>
	<tr>
		<td colspan="2" align="center"><input type="submit" value="Submit" class="formButton"></td>
	</tr>
	<%end if%>
</table>
</form>

<!--#include file="../includes/main_page_close.asp"-->