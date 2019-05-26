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

<%'load in record
task_id = request("id")
sql = sql_GetTasksByID(task_id)
Call RunsQL(sql, rs)
if not rs.eof then
	pProjectID		= rs("Project_ID")
	pClientID		= rs("Client_ID")
	pDescription	= rs("Description")
	pPriority		= rs("Priority")
	pOrderedBy		= rs("OrderedBy")
	pAssignedTo		= rs("AssignedTo")
	pEstimatedHours	= rs("EstimatedHours")
	pDateCreated	= rs("DateCreated")
	if pDateCreated <> "" then
		if not isDate(pDateCreated) then
			pDateCreated = date()
		end if
	else
		pDateCreated = date()
	end if
	pDateDue		= rs("DateDue")
	if pDateDue <> "" then
		if not isDate(pDateDue) then
			pDateDue = date()
		end if
	else
		pDateDue = date()
	end if
	pDone = rs("Done")
end if
rs.close
set rs = nothing
%>

<!--#include file="../includes/main_page_open.asp"-->

<%
if session("errorMsg") <> "" then
	response.write "<p class='alert'>" & session("errorMsg") & "</p>"
	session("errorMsg") = ""
end if
%>

<form method="POST" action="edit2.asp" id="strForm" name="strForm">
<table cellpadding="2" cellspacing="2" border="0" align="center">
	<tr><td colspan="2" bgcolor="<%=gsColorHighlight%>" align="center"><b class="homeHeader"><%=dictLanguage("Task_Details")%></b></td></tr>
	<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Client")%>/<%=dictLanguage("Project")%>:</b><font class="alert">*</font></td>
		<td>
			<select name="client" class="formstyleXL">
<%
sql = sql_GetActiveProjectsWithClients()
Call runSQL(sql, rs)
while not rs.eof
	response.write"<option value='" & rs("cl_id") & "P" & rs("Project_ID") & "' "
	if (rs("cl_id")=pClientID) and (rs("Project_ID")=pProjectID) then
		response.write "Selected"
	end if
	response.write">" & rs("Client_Name") & "{" & rs("Description") & "}</option>"
	rs.movenext
wend
rs.close
set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td valign="top"><b class="bolddark"><%=dictLanguage("Description")%>:</b><font class="alert">*</font></td>
		<td><textarea name="description" rows="4" class="formstyleLong"><%=pDescription%></textarea></td>
	</tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Priority")%>:</b><font class="alert">*</font></td>
		<td>
			<select name="priority" class="formstyleMed">
				<option value="1" <%if pPriority=1 or pPriority=-1 or pPriority then response.write "Selected"%>><%=dictLanguage("Low")%></option>
				<option value="2" <%if pPriority=2 then response.write "Selected"%>><%=dictLanguage("Medium")%></option>
				<option value="3" <%if pPriority=3 then response.write "Selected"%>><%=dictLanguage("High")%></option>
				<option value="0" <%if pPriority=0 then response.write "Selected"%>><%=dictLanguage("On_Hold")%></option>
			</select>
		</td>
	</tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Created_By")%>:</b><font class="alert">*</font></td>
		<td>
			<select name="orderedBy" class="formstyleLong">
<%
sql = sql_GetActiveEmployees()
Call RunSQL(sql, rs)
while not rs.eof
	response.write "<option value='" & rs("Employee_ID") & "'"
	if rs("Employee_ID") = pOrderedBy then
		response.write " Selected "
	end if
	response.write ">" & rs("EmployeeName") & "</option>" 
	rs.movenext
wend
%>
			</select>
		</td>
	</tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Assigned_To")%>:</b><font class="alert">*</font></td>
		<td>
			<select name="assignedTo" class="formstylelong">
<%
if not (rs.bof and rs.eof) then
	rs.movefirst
end if
while not rs.eof
	response.write "<option value='" & rs("Employee_ID") & "'"
	if rs("Employee_ID") = pAssignedTo then
		response.write " Selected "
	end if
	response.write ">" & rs("EmployeeName") & "</option>" 
	rs.movenext
wend
rs.close
set rs = nothing
%>
			</select>
		</td>
	</tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Estimated_Hours")%>:</b><font class="alert">*</font></td>
		<td><input type="text" value="<%=pEstimatedHours%>" name="estimatedHours" class="formstyleShort" maxlength="10"></td>
	</tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Date_Created")%>:</b><font class="alert">*</font></td>
		<td>
			<input type="Text" name="dateCreated" value="<%=pDateCreated%>" class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','dateCreated',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Date_Due")%>:</b><font class="alert">*</font></td>
		<td>
			<input type="text" name="dateDue" value="<%=pDateDue%>" class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','dateDue',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td><b class="bolddark"><%=dictLanguage("Done")%>?</b></td>
		<td><input type="CheckBox" name="taskDone" <%if pDone = 1 or pDone = -1 or pDone then%> checked <%end if%> value="1"> (<%=dictLanguage("Yes")%>)</td>
	</tr>
	<%if session("permTasksEdit") or session("employee_id") = pOrderedBy then%>
	<tr><td colspan="2" align="center"><input type="submit" value="Submit" class="formButton"></td></tr>
	<%end if%>
</table>
<input type="hidden" name="id" value="<%=task_id%>">
</form>

<!--#include file="../includes/main_page_close.asp"-->