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
''''''''''''''''''''''''''''''''''''''
'Validate Form Data
''''''''''''''''''''''''''''''''''''''

If (not request.form("start_date")="") then
	if not isdate(request.form("start_date")) then
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ReportInvalidStart")
	end if
End If

If (not request.form("end_date")="") then
	if not isdate(request.form("end_date")) then
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ReportInvalidEnd")
	end if
End If

If Session("strErrorMessage") <> "" then
	response.redirect "client_work_log_for_client.asp"
End If

'''''''''''''''''''''''''''
'Sessionize variables
'''''''''''''''''''''''''''

for each i in request.form
	session(i) = request.form(i)
next

'''''''''''''''''''''''''''
'load in employee id lookup table
'''''''''''''''''''''''''''
sql = sql_GetAllEmployees()
Call runSQL(sql, rs)
dim intCount
do while not rs.eof
	if intCount < cint(rs("Employee_ID")) then
		intCount = cint(rs("Employee_ID"))
	end if
	rs.movenext 
loop
rs.movefirst
redim employee(intCount)

do while not rs.eof
	for x=0 to UBound(employee)
		if x=cint(rs("Employee_ID")) then
			employee(x)="<a href='employees/default.asp?employeeID=" & cint(rs("Employee_ID")) & "'>" & rs("EmployeeName") & "</a>"
			exit for
		end if
	next
	rs.movenext
loop
rs.close
set rs = nothing

'''''''''''''''''''''''''''
'build query from form info or querystring info
'''''''''''''''''''''''''''
If Request.Querystring <> "" then
	strWhereClause = Request.QueryString("strWhereClause")
Else
	If session("client_id") <> "" Then
		temp=split(session("Client_ID"),"P")
		Client_ID = temp(0)
		Project_ID = temp(1)
		strCompanyClause = "tbl_Clients.Client_ID = " & Client_ID & " AND " & "tbl_Projects.Project_ID = " & Project_ID
		boolCompanyClause = True
		boolWhereClause = True
	End If
	If session("start_date") <> "" Then
		strStartClause = "tbl_timecards.dateworked >= " & DB_DATEDELIMITER & MediumDate(session("start_date")) & DB_DATEDELIMITER & " "
		boolStartDate = True
		boolHasDate = True
		boolWhereClause = True
	End If
	If session("end_date") <> "" Then
		strEndClause = "tbl_timecards.dateworked <= " & DB_DATEDELIMITER & MediumDate(session("end_date")) & DB_DATEDELIMITER & " "
		boolEndDate = True
		boolHasDate = True
		boolWhereClause = True
	End If
	If session("emp_id") <> "" Then
		strEmployeeClause = "tbl_timecards.employee_id = " & session("emp_id") & ""
		boolEmployeeClause = True
		boolWhereClause = True
	End If
	If session("reconciled") = "Yes" then
		strReconciledClause = "tbl_timecards.reconciled = " & prepareBit(1) & " "
		boolReconciledClause = TRUE
		boolWhereClause = True
	elseif session("reconciled") = "No" then
		strReconciledClause = "tbl_timecards.reconciled = 0 "
		boolReconciledClause = TRUE
		boolWhereClause = True
	End if		
	If session("typeofwork_id") <> "" Then
		strTypeofWorkClause = "tbl_TimeCards.TimeCardType_ID = " & session("typeofwork_id") & ""
		boolTypeofWorkClause = True
		boolWhereClause = True
	End If
	
	'put the clauses together
	If boolWhereClause Then
		strWhereClause = " WHERE("
		If boolCompanyClause Then
			strWhereClause = strWhereClause & strCompanyClause
			boolStarted = True
		End If
		If boolStartDate Then
			If boolStarted Then
				strWhereClause = strWhereClause & " AND "
			End If
			strWhereClause = strWhereClause & strStartClause
			boolstarted = true
		End If
		If boolEndDate Then
			If boolStarted Then
				strWhereClause = strWhereClause & " AND "
			End If
			strWhereClause = strWhereClause & strEndClause
			boolStarted = TRUE
		End If
		If boolEmployeeClause Then
			If boolStarted Then
				strWhereClause = strWhereClause & " AND "
			End If
			strWhereClause = strWhereClause & strEmployeeClause
			boolStarted = TRUE
		End If
		If boolReconciledClause then
			If boolStarted Then
				strWhereClause = strWhereClause & " AND " 
			End if
			strWhereClause = strWhereClause & strReconciledClause
		End if			
		If boolTypeofWorkClause Then
			If boolStarted Then
				strWhereClause = strWhereClause & " AND "
			End If
			strWhereClause = strWhereClause & strTypeofWorkClause
			boolStarted = TRUE
		End If
		strWhereClause = strWhereClause & ")"
		session("strWhereClause") = strWhereClause
	End If

End if

strSQL = "SELECT tbl_Clients.Client_Name, tbl_Clients.Rep_ID,tbl_Projects.project_management_ID, " & _
	"tbl_Employees.EmployeeLogin, tbl_TimeCards.TimeAmount, tbl_TimeCards.DateWorked, " & _
	"tbl_TimeCardTypes.TimeCardTypeDescription, tbl_TimeCards.[Non-Billable], " & _
	"tbl_TimeCards.Reconciled, tbl_timecards.timecard_id, tbl_timecards.workdescription, " & _
	"tbl_Projects.Description, tbl_TimeCards.Project_ID " & _
	"FROM (tbl_Clients INNER JOIN tbl_Projects ON tbl_Clients.Client_ID = tbl_Projects.Client_ID) " & _
	"INNER JOIN (tbl_TimeCardTypes INNER JOIN (tbl_Employees INNER JOIN tbl_TimeCards ON " & _
	"tbl_Employees.Employee_ID = tbl_TimeCards.Employee_ID) ON " & _
	"tbl_TimeCardTypes.TimeCardType_ID = tbl_TimeCards.TimeCardType_ID) ON " & _
	"tbl_Projects.Project_ID = tbl_TimeCards.Project_ID" & _
	strWhereClause & _
	" ORDER BY tbl_Clients.Client_Name, tbl_TimeCards.DateWorked,tbl_Employees.EmployeeLogin"
'Response.Write (strSQL)
Call RunSQL(strSQL, adoRS)
session("strWhereClause") = strWhereClause
%>


<!--#include file="../includes/main_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border="0" width="100%">
	<tr><td class="homeheader"><%=dictLanguage("Work_Log_For_Client")%>&nbsp;
<%
'Write the dates of the report if there are any specified.
Response.Write "("
If boolHasDate = True Then
	If boolStartDate = True Then
		If boolEndDate = False Then
			Response.Write dictLanguage("Since") & " " & Request.Form("start_date")
		Else
			Response.Write dictLanguage("Since") & " " & Request.Form("start_date") & " "
		End If
	End If
	If boolEndDate = True Then
		Response.Write dictLanguage("Through") & " " & Request.Form("end_date")
	Else
		Response.Write dictLanguage("Through") & " " & Date
	End If
Else
	Response.Write dictLanguage("Through") & " " & Date
End If
Response.Write ")"
Response.Write "</td></tr></table>"


If adoRS.EOF Then
	Response.Write "<p>" & dictLanguage("No_Report_Work") & "</p>"
End If
intTotal = 0
intBillable = 0
intNonBillable = 0
strProjectName = ""
strClientName = ""
boolFirst = True

Do While Not adoRS.EOF
	boolDataExists = True
	If adoRS("Client_Name") <> strClientName or adoRS("Description") <> strProjectName Then		
		'does not write unless the table header has been written
		If Not boolFirst Then %>

	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="9"></td></tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td class="bolddark"><%=dictLanguage("Billable")%>:</td>
		<td><b><%=intBillable%></b></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td class="bolddark"><%=dictLanguage("Non-Billable")%>:</td>
		<td><b><%=intNonBillable%></b></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td class="bolddark"><%=dictLanguage("Total")%>:</td>
		<td><b><%=intTotal%></b></td>
	</tr>
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="9"></td></tr>
</Table>
<%
		intTotal = 0
		intBillable = 0
		intNonBillable = 0
		End If
			
		'Checks to see if client or project has changed
		'If so write out the new header
		If adoRS("client_name") <> strClientName or adoRS("Description") <> strProjectName Then
		%><br><b><%=adoRS("client_name")%> {<%=adoRS("Description")%>}</b>
		|| <%=dictLanguage("Account_Rep")%>: <b><% x=adoRS("Rep_ID")
		response.write employee(x)%></b>
		<%End If%>
<table border="0" width=100% cellpadding="2" cellspacing="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td align="center" class="columnheader" nowrap><%=dictLanguage("Date")%></td>
		<td align="center" class="columnheader" nowrap><%=dictLanguage("Employee")%></td>
		<td align="center" class="columnheader" nowrap><%=dictLanguage("Time")%></td>
		<td align="center" class="columnheader" nowrap><%=dictLanguage("Billable")%></td>
		<td align="center" class="columnheader" nowrap><%=dictLanguage("Reconciled")%></td>
		<td align="center" class="columnheader" nowrap><%=dictLanguage("Activity")%></td>
		<td align="center" class="columnheader" nowrap width="100%"><%=dictLanguage("Comments")%></td>
	</tr>
<%
		strClientName = adoRS("client_name")
		strProjectName = adoRS("Description")
		boolFirst = False
	End If
	if strBGColor = "#FFFFFF" then
		strBGColor = "#EEEEEE"
	else
		strBGColor = "#FFFFFF"
	end if %>
	<tr bgcolor="<%=strBGColor%>">
		<td align="center" class="small"><a href="<%=gsSiteRoot%>timecards/timecard-edit.asp?timecard_id=<%=Server.URLEncode(adoRS("timeCard_ID"))%>"><%=adoRS("dateworked")%></a></td>
		<td align="center" class="small"><%=adoRS("employeelogin")%></td>
		<td align="center" class="small"><%=adoRS("timeamount")%></td>
		<td align="center" class="small">
<%	if adoRS("Non-Billable") = "True" then
		response.write dictLanguage("No")
	else
		response.write dictLanguage("Yes")
	end if%></td>
		<td align="center" class="small">
<%	if adoRS("Reconciled") = "True" then
		response.write dictLanguage("Yes")
	else
		response.write dictLanguage("No")
	end if%></td>	
		<td align="center" class="small"><%=adoRS("TimeCardTypeDescription")%></td>
		<td class="small"><%=adoRS("workdescription")%></td>
	</tr>
<%	if adoRS("Non-Billable") = "True" then
    	intNonBillable = intNonBillable + adoRS("timeamount")
	else
		intBillable = intBillable + adoRS("timeamount")
	End If
		intTotal = intTotal + adoRS("timeamount")
	adoRS.MoveNext
Loop
'Write out end data for last table
If boolDataExists Then
%>
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="9"></td></tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td class="bolddark"><%=dictLanguage("Billable")%>:</td>
		<td><b><%=intBillable%></b></td>
		<td colspan=5></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td class="bolddark"><%=dictLanguage("Non-Billable")%>:</td>
		<td><b><%=intNonBillable%></b></td>
		<td colspan=5></td>
	</tr>
	<tr>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td class="bolddark"><%=dictLanguage("Total")%>:</td>
		<td><b><%=intTotal%></b></td>
		<td colspan=5></td>	
	</tr>
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="9"></td></tr>
</Table>

<p align="center">
<a href="client_work_log_for_client.asp"><%=dictLanguage("Return_Work_Log_Form")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<%	
End If
adoRS.Close

for each i in request.form
	session(i) = ""
next
%>

<!--#include file="../includes/main_page_close.asp"-->