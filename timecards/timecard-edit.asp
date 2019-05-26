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
tID = Request("timecard_id")

if tID <> "" then
	sql = sql_GetTimecardsByID(tID)
	call RunSQL(sql, rsTimeCard)
	if not rsTimeCard.eof then
		client_ID = rsTimeCard("Client_ID")
		project_ID = rsTimeCard("Project_ID")
		emp = rsTimeCard("Employee_ID")
		if rsTimeCard("Task_ID")<>0 then
			task_ID = rsTimeCard("Task_ID")
		else 
			task_ID = 0
		end if
		code2 = trim(client_ID & "P" & project_ID & "T" & task_ID) %>

<!--#include file="../includes/main_page_open.asp"-->


<%		if session("strErrorMessage")<>"" then 
			Response.Write "<P class='alert'>" & session("strErrorMessage") & "<P>"
			Session("strErrorMessage") = ""
		end if %>

<form method="post" action="timecard-edit-processor.asp?timecard_id=<%=tID%>" name="strForm" id="strForm">
<table border="0" cellpadding="3" cellspacing="3" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Timecard_Details")%></td></tr>
	<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Client")%>:<font class="alert">*</font></td>
		<td>
			<select name="Client_ID" class="formstyleXL">
				<option value=""> - <%=dictLanguage("Select_One_Clients_Projects_Tasks")%> - </option>
<%		sql = sql_GetActiveTaskProjectClientList()
		Call RunSQL(sql, rsClientSelect)
		While Not rsClientSelect.EOF
			client_ID = rsClientSelect("Client_id")
			project_ID = rsClientSelect("Project_id")
			if rsClientSelect("Task_ID")<>"" then
				task_ID = rsClientSelect("Task_ID")
			else 
				task_ID = 0
			end if
			code1 = client_ID & "P" & project_ID & "T" & task_ID%>
	
				<option value="<%=code1%>"<%if code1=code2 then%> Selected <%end if%>><%=rsClientSelect("Client_Name")%> {<%=rsClientSelect("Description")%>} - <%=left(rsClientSelect("taskDesc"),15)%></option>
<%			rsClientSelect.MoveNext
		Wend%>

<%		sql = sql_GetActiveProjectsWithClients()
		call RunSQL(sql, rsClientSelect2)
		While Not rsClientSelect2.EOF
			client_ID = rsClientSelect2("cl_id")
			project_ID = rsClientSelect2("Project_ID")
			code1 = client_ID & "P" & project_ID & "T0"%>
	
				<option value="<%=code1%>"<%if code1=code2 then%> Selected <%end if%>><%=rsClientSelect2("Client_Name")%>{<%=rsClientSelect2("Description")%>}</option>
<%			rsClientSelect2.MoveNext
		Wend

		rsClientSelect.close
		Set rsClientSelect = Nothing
		rsClientSelect2.close
		set rsClientSelect2 = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Employee")%>:<font class="alert">*</font></td>
		<td>
			<select name="Employee" class="formstyleLong">
<%		'''''''''''''''''''''''''''''''''''''
		'Query the database and bring back a recordset
		'containing all of the employees' names
		'and their IDs
		'''''''''''''''''''''''''''''''''''''
		sql = sql_GetActiveEmployees()
		Call RunSQL(sql, rsEmployees)
		do while not rsEmployees.EOF %>
				<option value="<%=rsEmployees("Employee_ID")%>" <%if rsEmployees("Employee_ID") = rsTimeCard("Employee_ID") then %>selected<%end if%>><%=rsEmployees("EmployeeName")%></option>
<%			rsEmployees.MoveNext
		Loop
		rsEmployees.close
		set rsEmployees = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Project_Phase")%>:</td>
		<td>
			<select name="ProjectPhaseID" class="formstyleLong">
				<option value=""> - <%=dictLanguage("Select_One_If_App")%> - </option>
<%		sql = sql_GetProjectPhases()
		'response.write(sql)
		Call RunSQL(sql, rsPhase)
		While Not rsPhase.EOF
			projectphaseID = rsPhase("projectphaseID")
			projectphasename = rsPhase("ProjectphaseName") %>
				<option value="<%=projectphaseID%>" <%if rsTimeCard("projectphaseID") = projectphaseID then response.write "Selected"%>><%=projectPhaseName%></option>
<%			rsPhase.MoveNext
		Wend
		rsPhase.close
		Set rsPhase = Nothing%>
			</select>
		</td>
	</tr>  
    <tr>
		<td class="bolddark"><%=dictLanguage("Type_Of_Work")%>:<font class="alert">*</font></td>
		<td>
			<select name="TimeCardType_ID" class="formstyleLong">
				<option value=""> - <%=dictLanguage("Select_One")%> - </option>
<%		sql = sql_GetActiveTimecardTypes()
		call RunSQL(sql, rsTypeSelect)
		Do While Not rsTypeSelect.EOF %>
				<option 
<%			If rsTypeSelect("TimeCardType_ID") = rsTimeCard("TimeCardType_ID") then
				response.write "Selected"
			End If%>
				value="<%=rsTypeSelect("TimeCardType_ID")%>"><%=rsTypeSelect("TimeCardTypeDescription")%></option>
<%			rsTypeSelect.MoveNext
		Loop
		Set rsTypeSelect = Nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Hours")%>:<font class="alert">*</font></td>
		<td><input type="Text" name="TimeAmount" value="<%=rsTimeCard("TimeAmount")%>" class="formstyleTiny"></td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Non-Billable")%>:</td>
		<td><input type="CheckBox" name="Non-Billable" <%if rsTimeCard("Non-Billable") = "True" then%> checked <%end if%> value="yes">(<%=dictLanguage("Yes")%>)</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Reconciled")%>:</td>
		<td><input type="CheckBox" name="Reconciled" <%if rsTimeCard("Reconciled") = "True" then%> checked <%end if%> value="yes"></td>
	</tr>
<%		pDateWorked = rsTimeCard("DateWorked")
		if pDateWorked <> "" then
			if not isDate(pDateWorked) then
				pDateWorked = date()
			end if
		else
			pDateWorked = date()
		end if	%>
	<tr>
		<td class="bolddark"><%=dictLanguage("Date")%>:<font class="alert">*</font></td>
		<td>
			<input type="Text" name="DateWorked" value="<%=pDateWorked%>" class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','DateWorked',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td valign="top" class="bolddark"><%=dictLanguage("Description")%>:<font class="alert">*</font></td>
		<td><textarea name="WorkDescription" class="formstyleLong" rows="8"><%=rsTimeCard("WorkDescription")%></textarea></td>
	</tr> 
	<tr>
		<td colspan="2" align="center">
			<%if session("permTimecardsEdit") or session("employee_id")=rsTimeCard("Employee_ID") then%>
			<input type="Submit" name="Submit" value="Submit" class="formButton">
			<%end if
			  if session("permTimecardsDelete") or session("employee_id")=rsTimeCard("Employee_ID") then%>
			<input type="Submit" name="Submit" value="Delete" class="formButton" onClick="javascript: return confirm('Are you sure you want to delete this timecard?');">
			<%end if%>
		</td>
	</tr>
</table>
</form>
<%	end if 
	rsTimeCard.close
	set rsTimeCArd = nothing	
end if %>

<!--#include file="../includes/main_page_close.asp"-->
