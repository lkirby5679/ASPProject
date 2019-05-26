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
Function GetEmpName(ID)
	if ID <> "" then
		sql = sql_GetEmployeesByID(ID)
		'response.write sql
		Call RunSQL(sql, rs)
		if not rs.eof then
			empName = rs("EmployeeName")
		else
			empName = ""
		end if
		rs.close
		set rs = nothing
	else
		empName = ""
	end if
	GetEmpName = empName
End Function

Function GetClientName(ID)
	if ID <> "" then
		sql = sql_GetClientsByID(ID)
		'response.write sql
		Call RunSQL(sql, rs)
		if not rs.eof then
			clientName = rs("Client_Name")
		else
			clientName = ""
		end if
		rs.close
		set rs = nothing
	else
		clientName = ""
	end if
	GetClientName = clientName
End Function

Function GetProjectName(ID)
	if ID <> "" then
		sql = sql_GetProjectsByID(ID)
		'response.write sql
		Call RunSQL(sql, rs)
		if not rs.eof then
			projectName = rs("description")
		else
			projectName = ""
		end if
		rs.close
		set rs = nothing
	else
		projectName = ""
	end if
	GetProjectName = projectName
End Function

for each i in request.form 
    session(i) = request.form(i)
next
if session("description") = "" then
    session("errorMsg") = dictLanguage("Error_TaskNoName") & "<BR>"
end if
if session("priority") = "" then
    session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskNoPriority") & "<BR>"
end if
if session("orderedBy") = "" then
    session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskNoAssignedBy") & "<BR>"
end if
if session("assignedTo") = "" then
    session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskAssignedTo") & "<BR>"
end if
if session("estimatedHours") = "" then
    session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskEstHours") & "<BR>"
else
	if not isNumeric(session("estimatedHours")) then
		session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskInvalidHours") & "<br>"
	end if	
end if
if session("client") = "" then
    session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskNoClient") & "<BR>"
Else 'make sure this client is active
	intClientId = session("client")
	intClientId = mid(intClientId,1,inStr(LCase(intClientId),"p") - 1)
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Establish Database Connection
	''''''''''''''''''''''''''
	strClientSQL = sql_GetActiveClientsByID(intClientId)
	Call runSQL(strClientSQL, rsClient)
	if rsClient.EOF then 'client not active
		session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskInvalidClient") & "<br>"
	end if
	rsClient.close
	set rsClient = nothing
End if
if session("dateCreated") = "" then
    session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskNoDateCreated") & "<BR>"
else
	if not isDate(session("dateCreated")) then
		session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskInvalidDateCreated") & "<br>"
	end if	
end if
if session("dateDue") = "" then
    session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskNoDateDue") & "<BR>"
else
	if not isDate(session("dateDue")) then
		session("errorMsg") = session("errorMsg") & dictLanguage("Error_TaskInvalidDateDue") & "<br>"
	end if	
end if

if session("errorMsg") <> "" then
     response.redirect "new.asp"
end if

temp2 = split(session("client"),"P")
client_ID = temp2(0)
project_ID = temp2(1)

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

sql = sql_InsertTask( _
	SQLEncode(session("description")), _
	session("priority"), _
	session("orderedBy"), _
	session("assignedTo"), _
	session("estimatedHours"), _
	client_ID, _
	project_ID, _
	session("dateCreated"), _
	time(), _
	0, _
	1, _
	session("dateDue"))
'Response.Write (sql)
Call DoSQL(sql)

sqlemailto = sql_GetEmployeesByID((session("assignedTo")))
Call RunSQL(sqlemailto, rsemailto)
if not rsemailto.eof then
	emailto = rsemailto("emailaddress")

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'send the email
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	msg = dictLanguage("Task_Inst_3")
	if gsTaskEmails then
		Call SendEmail("", emailto, "", gsAdminEmail, "New Task", msg, "", "", _
			"", "", "", FALSE)
	end if
end if
rsemailto.close
set rsemailto = nothing %>

<!--#include file="../includes/main_page_open.asp"-->

<table cellpadding="3" cellspacing="3" border=0 align="center">
	<tr><td colspan="2" bgcolor="<%=gsColorHighlight%>" align="center" class="homeheader"><%=dictLanguage("Task_Details")%>: <%=dictLanguage("Added_Task")%></td></tr>
	<tr>
		<td>
			<table cellpadding="3" cellspacing="3" border="0" style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Client")%>:</td>
					<td valign="top"><b><%=GetClientName(client_ID)%></b></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Project")%>:</td>
					<td valign="top"><b><%=GetProjectName(project_ID)%></b></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Description")%>:</td>
					<td valign="top"><%=replace(session("description"),vbcrlf,"<BR>")%></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Created_By")%>:</td>
					<td valign="top"><%=GetEmpName(session("orderedBy"))%></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Assigned_To")%>:</td>
					<td valign="top"><%=GetEmpName(session("assignedTo"))%></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Estimated_Hours")%>:</td>
					<td valign="top"><%=session("estimatedHours")%></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Date_Created")%>:</td>
					<td valign="top"><%=session("dateCreated")%></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Date_Due")%>:</td>
					<td valign="top"><%=session("dateDue")%></td>
				</tr>
			</table>
			<p align="center">
			<a href="new.asp"><%=dictLanguage("Add_Another_Task")%></a><br>
			<a href="<%=gsSiteRoot%>projects/project-status.asp"><%=dictLanguage("Project_Status")%></a><br>
			<a href="<%=gsSiteRoot%>tasks/view_mine.asp"><%=dictLanguage("View_My_Tasks")%></a><br>
			<a href="view_my_assigned.asp"><%=dictLanguage("View_Tasks_I_Assigned")%></a><br>
			<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
			</p>
		</td>
	</tr>
</table>

<%for each i in Request.Form
	session(i) = ""
  next%>

<!--#include file="../includes/main_page_close.asp"-->