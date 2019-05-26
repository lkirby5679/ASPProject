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
''''''''''''''''''''''''''''''''''''''''''''''''''''
'Establish any prepoulation values
'''''''''''''''''''''''''''''''''''''''
'if they're coming from tasks...
desc = request.querystring("desc")
from = request.querystring("from")
project = trim(request.querystring("project"))
client = trim(request.querystring("client"))
session("tsk") = request.querystring("task")
code2 = client & "P" & project & "T" & session("tsk")

''''''''''''''''''''''''''''''''''''''''''''''''''''
'double check to  make sure the data isn't coming from the session
'instead
''''''''''''''''''''''''''''''''''''''''''''''''''''
if left(code2,1) <> "P" then
	strClientCheck = code2
else
	strClientCheck = session("Client_ID")
end if

if desc <> "" then
	strDesc = desc
else
	strDesc = session("WorkDescription")
end if %>

<!--#include file="../includes/main_page_open.asp"-->

<%
if session("strErrorMessage")<>"" then 
	Response.Write "<P class='alert'>" & session("strErrorMessage") & "<P>"
	Session("strErrorMessage") = ""
end if %>

<form method="post" action="timecard-processor.asp" id="strForm" name="strForm">
<table border="0" cellpadding="3" cellspacing="3" align="center" style="zIndex: 25;">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeHeader"><%=dictLanguage("Timecard_Details")%></td></tr>
	<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Client/Project/Task")%>:<font class="alert">*</font></td>
		<td>
			<select name="Client_ID" class="formstyleXL">
				<option value=""> - <%=dictLanguage("Select_One_Clients_Projects_Tasks")%> - </option>
<%
sql = sql_GetActiveTaskProjectClientList()
'response.write sql & "<BR>"
Call RunSQL(sql, rsClientSelect)
While Not rsClientSelect.EOF
	client_ID = rsClientSelect("Client_ID")
	project_ID = rsClientSelect("Project_ID")
	task_ID = rsClientSelect("Task_ID")
	code1 = client_ID & "P" & project_ID & "T" & task_ID %>
				<option value="<%=code1%>"<%if code1 = strClientCheck then%> Selected <%end if%>><%=rsClientSelect("Client_Name")%> {<%=rsClientSelect("Description")%>} - <%=left(rsClientSelect("taskDesc"),15)%></option>
<%	rsClientSelect.MoveNext
Wend
rsClientSelect.close
set rsClientSelect = nothing

sql = sql_GetActiveProjectsWithClients()
'Response.Write sql & "<BR>"
Call RunSQL(sql, rsClientSelect)
While Not rsClientSelect.EOF
	client_ID = rsClientSelect("cl_id")
	project_ID = rsClientSelect("Project_ID")
	code1 = client_ID & "P" & project_ID %>
				<option value="<%=code1%>"><%=rsClientSelect("Client_Name")%> {<%=rsClientSelect("Description")%>}</option>
<%	rsClientSelect.MoveNext
Wend
rsClientSelect.close
Set rsClientSelect = Nothing %>
			</select>
		</td>
	</tr>
    <tr>
		<td class="bolddark"><%=dictLanguage("Project_Phase")%>:</td>
		<td>
			<select name="ProjectPhaseID" class="formstyleLong">
				<option value=""> - <%=dictLanguage("Select_One_If_App")%> - </option>
<%
sql = sql_GetProjectPhases()
'response.write(sql)
call RunSQL(sql, rsPhase)
While Not rsPhase.EOF
	projectphaseID = rsPhase("projectphaseID")
	projectphasename = rsPhase("ProjectphaseName") %>
				<option value="<%=projectphaseID%>"><%=projectPhaseName%></option>
<%	rsPhase.MoveNext
Wend
rsPhase.close
Set rsPhase = Nothing %>
			</select>
		</td>
	</tr>  
    <tr>
		<td class="bolddark"><%=dictLanguage("Type_Of_Work")%>:<font class="alert">*</font></td>
		<td>
			<select name="TimeCardType_ID" class="formstyleLong">
				<option value=""> - <%=dictLanguage("Select_One")%> - </option>
<%
sql = sql_GetActiveTimecardTypes()
Call RunSQL(sql, rsTypeSelect)
Do While Not rsTypeSelect.EOF
	intTypeOfWork = CINT(rsTypeSelect("TimeCardType_ID"))
	if isNumeric(session("TimeCardType_ID")) = True then
		intWorkType = CINT(session("TimeCardType_ID"))
	else
		intWorkType = 0
	end if
	if rsTypeSelect("TimeCardTypeDescription")<>"PTO" then%>
				<option value="<%=intTypeOfWork%>" <%if intTypeOfWork = intWorkType then%> selected <%end if%>><%=rsTypeSelect("TimeCardTypeDescription")%></option>
<%	end if
	rsTypeSelect.MoveNext
Loop
rsTypeSelect.close
Set rsTypeSelect = Nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Hours")%>:<font class="alert">*</font></td>
		<td><input type="Text" name="TimeAmount" value="<%=session("TimeAmount")%>" class="formstyleTiny"></td>
	</tr>
	<tr>
		<td class="bolddark"><%=dictLanguage("Date")%>:<font class="alert">*</font></td>
		<td>
			<input type="Text" name="DateWorked" value="<%=date()%>" class="formstyleShort" onKeyPress="txtDate_onKeypress();" maxlength="10">
			<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','DateWorked',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td valign="top" class="bolddark"><%=dictLanguage("Description")%>:<font class="alert">*</font></td>
		<td><textarea name="WorkDescription" rows="8" class="formstyleLong"><%=strDesc%></textarea></td>
	</tr> 
	<%if session("permTimecardsAdd") then%>
	<tr>
		<td colspan="2" align="center">
			<input type="Hidden" name="from" value="<%=from%>">
			<input type="Submit" name="Submit" value="Submit" class="formButton">
		</td>
	</tr>		
	<%end if%>
</table>
</form>

<p align="center">
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<!--#include file="../includes/main_page_close.asp"-->