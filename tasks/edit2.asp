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

temp2 = split(request.form("client"),"P")
client_ID = temp2(0)
project_ID = temp2(1)

for each i in Request.Form
	session(i) = Request.Form(i)
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
     response.redirect "edit.asp?id=" & session("id") & ""
end if

if session("taskDone") = "1" then
	session("taskDone") = 1
else
	session("taskDone") = 0
end if

'insert changed data into database
sql = sql_UpdateTasks( _
	client_ID, _
	project_ID, _
	SQLEncode(session("description")), _
	session("priority"), _
	session("orderedBy"), _
	session("assignedTo"), _
	session("estimatedHours"), _
	session("dateCreated"), _
	session("dateDue"), _
	session("taskDone"), _
	session("id"))
'response.write sql
Call DoSQL(sql)

Select Case session("priority")
	case 0
		session("priority") = dictLanguage("On_Hold")
	case 1
		session("priority") = dictLanguage("Low")
	case 2
		session("priority") = dictLanguage("Medium")
	case 3
		session("priority") = dictLanguage("High")
	case else
		session("priority") = dictLanguage("Undefined")
end select

Select Case session("taskDone")
	case 0
		session("taskDone") = dictLanguage("No")
	case 1 
		session("taskDone") = dictLanguage("Yes")
end select						

sql = sql_GetEmployeesByID(session("assignedTo"))
Call runSQL(sql, rsemailto)
if not rsemailto.eof then
	emailto = rsemailto("emailaddress")
end if
rsemailto.close
set rsemailto = nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'send the email
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
msg = dictLanguage("Task_Inst_2") & VBCRLF & VBCRLF & dictLanguage("Description") & ":" &  VBCRLF & session("description") & VBCRLF & ""

if gsTaskEmails then	
	Call SendEmail("", emailto, "", gsAdminEmail, "New Task", msg, "", "", _
		"", "", "", FALSE)
end if
%>

<!--#include file="../includes/main_page_open.asp"-->

<table cellpadding="3" cellspacing="3" border="0" align="center">
	<tr><td bgcolor="<%=gsColorHighlight%>" align="center"><b class="homeheader"><%=dictLanguage("Task_Details")%>: <%=dictLanguage("Updated_Task")%></td></tr>
	<tr>
		<td>
			<table border="0" cellpadding="3" cellspacing="3" style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Client")%>:</td>
					<td valign="top"><b><%=GetClientName(client_ID)%></b></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Project")%>:</td>
					<td valign="top"><b><%=GetProjectName(project_ID)%></b></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Description")%>:</td>
					<td valign="top"><%=replace(session("description"),vbcrlf,"<BR>")%></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Priority")%>:</td>
					<td valign="top"><%=session("priority")%></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Created_By")%>:</td>
					<td valign="top"><%=GetEmpName(session("orderedBy"))%></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Assigned_To")%>:</td>
					<td valign="top"><%=GetEmpName(session("assignedTo"))%></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Estimated_Hours")%>:</td>
					<td valign="top"><%=session("estimatedHours")%></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Date_Created")%>:</td>
					<td valign="top"><%=session("dateCreated")%></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Date_Due")%>:</td>
					<td valign="top"><%=session("dateDue")%></td>
				</tr>
				<tr>
					<td valign="top" nowrap class="bolddark"><%=dictLanguage("Done")%>:</td>
					<td valign="top"><%=session("taskDone")%></td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<p align="center">
<a href="new.asp"><%=dictLanguage("Add_Another_Task")%></a><br>
<a href="<%=gsSiteRoot%>projects/project-status.asp"><%=dictLanguage("Project_Status")%></a><br>
<a href="<%=gsSiteRoot%>tasks/view_mine.asp"><%=dictLanguage("View_My_Tasks")%></a><br>
<a href="view_my_assigned.asp"><%=dictLanguage("View_Tasks_I_Assigned")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->

<%
for each i in Request.Form
	session(i) = ""
next
%>