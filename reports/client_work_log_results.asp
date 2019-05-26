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
	response.redirect "client_work_log.asp"
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
'dim employee(150)
'dim employeeBillableTotal(150)
'dim employeeNonBillableTotal(150)
'dim employeeCostTotal(150)
'dim employeeRateTotal(150)

sql = sql_GetAllEmployees()
Call RunSQL(sql, rs)

dim intEmpCount
do while not rs.eof
	if intEmpCount < cint(rs("Employee_ID")) then
		intEmpCount = cint(rs("Employee_ID"))
	end if
	rs.movenext 
loop
rs.movefirst
redim employee(intEmpCount)
redim employeeBillableTotal(intEmpCount)
redim employeeNonBillableTotal(intEmpCount)
redim employeeCostTotal(intEmpCount)
redim employeeRateTotal(intEmpCount)

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

dim activity(150)
dim activityBillableTotal(150)
dim activityNonBillableTotal(150)
dim activityCostTotal(150)
dim activityRateTotal(150)

strSQL = sql_GetActiveTimecardTypes()
Call runSQL(strSQL, rs)
do while not rs.eof
	for x=0 to UBound(activity)
		if x=cint(rs("timecardtype_id")) then
			activity(x)=rs("timecardtypedescription")
			exit for
		end if
	next
	rs.movenext
loop
rs.close
set rs = nothing

dim dept(150)
dim deptBillableTotal(150)
dim deptNonBillableTotal(150)
dim deptCostTotal(150)
dim deptRateTotal(150)

strSQL = sql_GetEmployeeTypes()
Call runSQL(strSQL, rs)
do while not rs.eof
	for x=0 to UBound(dept)
		if x=cint(rs("employeetype_id")) then
			dept(x)=rs("employeetype")
			exit for
		end if
	next
	rs.movenext
loop
rs.close
set rs= nothing

dim phase(150)
phase(0) = "No Phase Indicated"
dim phaseBillableTotal(150)
dim phaseNonBillableTotal(150)
dim phaseCostTotal(150)
dim phaseRateTotal(150)

strSQL = sql_GetProjectPhases()
Call runSQL(strSQL, rs)
do while not rs.eof
	for x=0 to UBound(phase)
		if x=cint(rs("projectphaseid")) then
			phase(x)=rs("projectphasename")
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
		temp=split(session("client_ID"),"P")
		Client_ID = temp(0)
		Project_ID = temp(1)
		strCompanyClause = "tbl_Clients.Client_ID = " & Client_ID & " AND " & "tbl_Projects.Project_ID = " & Project_ID
		boolCompanyClause = True
	End If	
	If session("start_date") <> "" Then
		strStartClause = "tbl_timecards.dateworked >= " & DB_DATEDELIMITER & MediumDate(session("start_date")) & DB_DATEDELIMITER & " "
		boolStartDate = True
		boolHasDate = True
	End If
	If session("end_date") <> "" Then
		strEndClause = "tbl_timecards.dateworked <= " & DB_DATEDELIMITER & MediumDate(session("end_date")) & DB_DATEDELIMITER & " "
		boolEndDate = True
		boolHasDate = True
	End If
	If session("proj_type_id") <> "" Then
		strProjTypeClause = "tbl_Projects.ProjectType_ID = " & session("proj_type_id")
		boolProjTypeClause = True
	End If
	If session("work_type_id") <> "" Then
		strWorkTypeClause = "tbl_timecards.TimeCardType_ID = " & session("work_type_id")
		boolWorkTypeClause = True
	End If
	If session("emp_id") <> "" Then
		strEmployeeClause = "tbl_timecards.employee_id = " & session("emp_id")
		boolEmployeeClause = True
	End If
	If session("department") <> "" Then
		strDeptClause = "tbl_employees.department = '" & session("department") & "'"
		boolDeptClause = True
	End If
	If session("reconciled") = "Yes" then
		strReconciledClause = "tbl_timecards.reconciled = " & prepareBit(1) & " "
		boolReconciledClause = TRUE
	elseif session("reconciled") = "No" then
		strReconciledClause = "tbl_timecards.reconciled = 0 "
		boolReconciledClause = TRUE
	End if	
	If session("rep_id") <> "" Then
		strRepClause = "(tbl_Clients.Rep_ID = " & session("rep_id") & " or tbl_Projects.AccountExec_ID = " & session("rep_id") & ")"
		boolRepClause = True
	End If

	'put the clauses together
	'strWhereClause = " WHERE tbl_Projects.ProjectType_ID = tbl_ProjectTypes.ProjectType_ID and tbl_Clients.Client_ID = tbl_Projects.Client_ID and tbl_Employees.Employee_ID = tbl_TimeCards.Employee_ID and (tbl_TimeCardTypes.TimeCardType_ID = tbl_TimeCards.TimeCardType_ID or tbl_Timecards.timecardtype_id=0) and (tbl_Projects.Project_ID = tbl_TimeCards.Project_ID or tbl_timecards.project_id=0) "
	strWhereClause = " WHERE 1 = 1 "
	If boolCompanyClause Then
		strWhereClause = strWhereClause & " AND " & strCompanyClause
	End If
	If boolStartDate Then
		strWhereClause = strWhereClause & " AND " & strStartClause
	End If
	If boolEndDate Then
		strWhereClause = strWhereClause & " AND " & strEndClause
	End If
	If boolProjTypeClause Then
		strWhereClause = strWhereClause & " AND " & strProjTypeClause
	End If
	If boolWorkTypeClause Then
		strWhereClause = strWhereClause & " AND " & strWorkTypeClause
	End If
	If boolEmployeeClause Then
		strWhereClause = strWhereClause & " AND " & strEmployeeClause
	End If
	If boolDeptClause Then
		strWhereClause = strWhereClause & " AND " & strDeptClause
	End If	
	If boolReconciledClause then
		strWhereClause = strWhereClause & " AND " & strReconciledClause
	End if	
	If boolRepClause Then
		strWhereClause = strWhereClause & " AND " & strRepClause
	End If
	session("strWhereClause") = strWhereClause
End if

strSQL = "SELECT tbl_Clients.Client_Name, tbl_Clients.Rep_ID, " & _
	"tbl_Employees.EmployeeLogin, tbl_Employees.Employeetype_id, tbl_TimeCards.TimeAmount, " & _
	"tbl_TimeCards.DateWorked, tbl_TimeCards.Cost, tbl_TimeCards.Rate, tbl_TimeCards.Employee_ID, " & _
	"tbl_Timecards.timecardtype_ID, tbl_TimeCardTypes.TimeCardTypeDescription, " & _
	"tbl_TimeCards.[Non-Billable], tbl_TimeCards.Reconciled, tbl_Timecards.TimeCard_ID, " & _
	"tbl_TimeCards.WorkDescription, tbl_Projects.Description, tbl_Projects.ProjectType_ID, " & _
	"tbl_timecards.projectphaseID, tbl_ProjectTypes.ProjectTypeDescription as ProjType, " & _
	"tbl_TimeCards.Project_ID "
	'"FROM tbl_ProjectTypes, tbl_Clients, tbl_projects, tbl_TimeCardTypes, tbl_timeCards, tbl_Employees " & _
strSQL = strSQL & "FROM (tbl_TimeCardTypes RIGHT JOIN (tbl_ProjectTypes RIGHT JOIN (tbl_Employees RIGHT JOIN (tbl_TimeCards LEFT JOIN tbl_Projects ON tbl_TimeCards.Project_ID = tbl_Projects.Project_ID) ON tbl_Employees.Employee_ID = tbl_TimeCards.Employee_ID) ON tbl_ProjectTypes.ProjectType_ID = tbl_Projects.ProjectType_ID) ON tbl_TimeCardTypes.TimeCardType_ID = tbl_TimeCards.TimeCardType_ID) LEFT JOIN tbl_Clients ON tbl_TimeCards.Client_ID = tbl_Clients.Client_ID"
strSQL = strSQL & strWhereClause & _
	" ORDER BY tbl_Projects.ProjectType_ID, tbl_Clients.Client_Name, tbl_Projects.Description, " & _
	"tbl_timecards.[non-billable] desc, tbl_Employees.EmployeeLogin, tbl_TimeCards.DateWorked"
'response.write strSQL & "<BR>"
Call runSQL(strSQL, adoRS)
%>

<!--#include file="../includes/main_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border="0" width="100%">
	<tr><td class="homeheader"><%=dictLanguage("Work_Log_Internal")%>&nbsp; 
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
	Response.Write dictLanguage("No_Report_Work")
End If

intTotal = 0
intGrandTotal = 0
intBillable = 0
intBillableTotal = 0
intNonBillable = 0
intNonBillableTotal = 0

strProjectName = "xiudmadfpoi"
strClientName = "asdifpjjpmmmiy"
boolFirst = True

strBGColor = "#FFFFFF"
Do While Not adoRS.EOF
	boolDataExists = True
	If adoRS("Client_Name") <> strClientName or adoRS("Description") <> strProjectName then
		'does not write unless the table header has been written
		If Not boolFirst Then
%>
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
<%			intTotal = 0
			intBillable = 0
			intNonBillable = 0
		End If
			
		'Checks to see if client or project has changed
		'If so write out the new header
		If adoRS("client_name") <> strClientName or adoRS("Description") <> strProjectName then
			Response.Write "<br><b>" & adoRS("client_name") & " {" & adoRS("Description") & "}</b><br>"
			Response.Write dictLanguage("Account_Rep") & ": " 
			if adoRS("rep_id")<>"" then
				x = int(adoRS("Rep_ID"))
				response.write "<b>" & employee(x) & "</b>"
			else
				Response.Write "<b>None</b>"
			end if 
			Response.write " -- "
			Response.write dictLanguage("Project_Type") & ": " 
			if adoRS("ProjType")<>"" then
				Response.Write "<b>" & adoRS("ProjType") & "</b>"
			else
				Response.Write "<b>None</b>"
			end if
	    end if %>
<table border="0" width=100% cellspacing="2" cellpadding="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td class="columnheader"><!--this cell will contain the edit link--></td>
		<td class="columnheader" align="center" width=50><%=dictLanguage("Date")%></td>
		<td class="columnheader" align="center" width=50><%=dictLanguage("Employee")%></td>
		<td class="columnheader" align="center" width=50><%=dictLanguage("Time")%></td>
		<td class="columnheader" align="center" width=50><%=dictLanguage("Type")%></td>
		<td class="columnheader" align="center" width=50><%=dictLanguage("Billable")%></td>
		<td class="columnheader" align="center" width=50><%=dictLanguage("Reconciled")%></td>
		<td class="columnheader" align="center" width=50><%=dictLanguage("Phase")%></td>
		<td class="columnheader" align="center" width="100%"><%=dictLanguage("Comments")%></td>
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
		<td align="center" class="small" valign="top"><a href="<%=gsSiteRoot%>timecards/timecard-edit.asp?timecard_id=<%=Server.URLEncode(adoRS("timeCard_ID"))%>"><%=dictLanguage("Edit")%></a></td>
		<td align="center" class="small" valign="top"><%=adoRS("dateworked")%></td>
		<td align="center" class="small" valign="top"><%=adoRS("employeelogin")%></td>
		<td align="center" class="small" valign="top"><%=adoRS("timeamount")%></td>
		<td align="center" class="small" valign="top"><%=adoRS("timecardtypedescription")%></td>
		<td align="center" class="small" valign="top">
<%	if adoRS("Non-Billable") = "True" then
		response.write dictLanguage("No")
	else
		response.write dictLanguage("Yes")
	end if%></td>
		<td align="center" class="small" valign="top">
<%	if adoRS("Reconciled") = "True" then
		response.write dictLanguage("Yes")
	else
		response.write dictLanguage("No")
	end if%></td>
		<td align="center" class="small" valign="top">
<%	if adoRS("ProjectPhaseID")="0" then 
		response.write dictLanguage("None")
	elseif adoRS("ProjectPhaseID")<>"" then
		response.write GetProjectPhaseName(adoRS("ProjectPhaseID"))
	else
		response.write "&nbsp;"
	end if %></td>
		<td class="small" valign="top"><%=adoRS("workdescription")%></td>
	</tr>
<%
'Update Counters for bottom of the page

if adoRS("Non-Billable") = "True" then
    intNonBillable = intNonBillable + adoRS("timeamount")
    intNonBillableTotal = intNonBillableTotal + adoRS("timeamount")
	if adoRS("employee_id")<>"" then
		employeeNonBillableTotal(cint(adoRS("employee_id")))=employeeNonBillableTotal(cint(adoRS("employee_id")))+adoRS("timeamount")
	end if
	if adoRS("timecardtype_id")<>"" then
		activityNonBillableTotal(cint(adoRS("timecardtype_id")))=activityNonBillableTotal(cint(adoRS("timecardtype_id")))+adoRS("timeamount")
	end if
	if adoRS("employeetype_id")<>"" then
		deptNonBillableTotal(cint(adoRS("employeetype_id")))=deptNonBillableTotal(cint(adoRS("employeetype_id")))+adoRS("timeamount")
	end if
	if adoRS("projectphaseid")<>"" then
		phaseNonBillableTotal(cint(adoRS("projectphaseid")))=phaseNonBillableTotal(cint(adoRS("projectphaseid")))+adoRS("timeamount")
	end if
else
	intBillable = intBillable + adoRS("timeamount")
    intBillableTotal = intBillableTotal + adoRS("timeamount")
	if adoRS("rate")<>"" then
		intRateTotal = intRateTotal + adoRS("rate")
	end if
	if adoRS("employee_id")<>"" then
		employeeBillableTotal(cint(adoRS("employee_id")))=employeeBillableTotal(cint(adoRS("employee_id")))+adoRS("timeamount")
		employeeRateTotal(cint(adoRS("employee_id")))=employeeRateTotal(cint(adoRS("employee_id")))+adoRS("rate")
	end if
	if adoRS("timecardtype_id")<>"" then
		activityBillableTotal(cint(adoRS("timecardtype_id")))=activityBillableTotal(cint(adoRS("timecardtype_id")))+adoRS("timeamount")
		activityRateTotal(cint(adoRS("timecardtype_id")))=activityRateTotal(cint(adoRS("timecardtype_id")))+adoRS("rate")
	end if
	if adoRS("employeetype_id")<>"" then
		deptBillableTotal(cint(adoRS("employeetype_id")))=deptBillableTotal(cint(adoRS("employeetype_id")))+adoRS("timeamount")
		deptRateTotal(cint(adoRS("employeetype_id")))=deptRateTotal(cint(adoRS("employeetype_id")))+adoRS("rate")
	end if
	if adoRS("projectphaseID")<>"" then
		phaseBillableTotal(cint(adoRS("projectphaseid")))=phaseBillableTotal(adoRS("projectphaseid"))+adoRS("timeamount")
		phaseRateTotal(cint(adoRS("projectphaseid")))=phaseRateTotal(cint(adoRS("projectphaseid")))+adoRS("rate")
	end if
End If
	intTotal = intTotal + adoRS("timeamount")
    intGrandTotal = intGrandTotal + adoRS("timeamount")
	intCostTotal = intCostTotal + adoRS("cost")
	if adoRS("employee_id")<>"" then
		employeeCostTotal(cint(adoRS("employee_id")))=employeeCostTotal(cint(adoRS("employee_id")))+adoRS("cost")
	end if
	if adoRS("timecardtype_id")<>"" then
		activityCostTotal(cint(adoRS("timecardtype_id")))=activityCostTotal(cint(adoRS("timecardtype_id")))+adoRS("cost")
	end if
	if adoRS("employeetype_id")<>"" then
		deptCostTotal(cint(adoRS("employeetype_id")))=deptCostTotal(cint(adoRS("employeetype_id")))+adoRS("cost")
	end if
	if adoRS("projectphaseID")<>"" then	
		phaseCostTotal(cint(adoRS("projectphaseid")))=phaseCostTotal(cint(adoRS("projectphaseid")))+adoRS("cost")
	end if
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
		<td colspan=5>&nbsp;</td>	
	</tr>
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="9"></td></tr>		
</Table>

<br>

<table border="0" cellpadding="2" cellspacing="2" width="100%">
	<tr>
		<td>

			<table border="0" cellpadding="2" cellspacing="2" width="600">
				<tr bgcolor="<%=gsColorHighlight%>">
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Type")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Non-Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Total_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Cost")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Value")%></td>
				</tr>		
<%	strBGColor = "#FFFFFF"
	for x=0 to 150
		if activity(x)<>"" then 
			if strBGColor = "#FFFFFF" then
				strBGColor = "#EEEEEE"
			else
				strBGColor = "#FFFFFF"
			end if		%>
				<tr bgcolor="<%=strBGColor%>">
					<td nowrap><%=activity(x)%>&nbsp;</font></td>
					<td align="right"><%=activityBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=activityNonBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=(activityBillableTotal(x) + activityNonBillableTotal(x))%>&nbsp;</td>
					<td align="right"><%=formatNumber(activityCostTotal(x),2)%>&nbsp;</td>
					<td align="right"><%=formatNumber(activityRateTotal(x),2)%>&nbsp;</td>
				</tr>			
<%		end if
	next
%>
			</table>

			<br>

			<table border="0" cellpadding="2" cellspacing="2" width="600">
				<tr bgcolor="<%=gsColorHighlight%>">
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Department")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Non-Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Total_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Cost")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Value")%></td>
				</tr>		
<%	strBGColor = "#FFFFFF"
	for x=0 to 150
		if Dept(x)<>"" then 
			if strBGColor = "#FFFFFF" then
				strBGColor = "#EEEEEE"
			else
				strBGColor = "#FFFFFF"
			end if		%>
				<tr bgcolor="<%=strBGColor%>">
					<td><%=dept(x)%>&nbsp;</font></td>
					<td align="right"><%=deptBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=deptNonBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=(deptBillableTotal(x) + deptNonBillableTotal(x))%>&nbsp;</td>
					<td align="right"><%if deptCostTotal(x)<>"" then Response.write formatNumber(deptCostTotal(x),2) else Response.Write "0.00"%>&nbsp;</td>
					<td align="right"><%=formatNumber(deptRateTotal(x),2)%>&nbsp;</td>
				</tr>			
<%		end if
	next
%>
			</table>

			<br>

			<table border="0" cellpadding="2" cellspacing="2" width="600">
				<tr bgcolor="<%=gsColorHighlight%>">
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Employee")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Non-Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Total_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Cost")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Value")%></td>
				</tr>		
<%	strBGColor = "#FFFFFF"
	for x=0 to intEmpCount
		if EmployeeBillableTotal(x) > 0 or EmployeeNonBillableTotal(x) > 0 then 
			if strBGColor = "#FFFFFF" then
				strBGColor = "#EEEEEE"
			else
				strBGColor = "#FFFFFF"
			end if	%>
				<tr bgcolor="<%=strBGColor%>">
					<td><%=employee(x)%>&nbsp;</td>
					<td align="right"><%=EmployeeBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=EmployeeNonBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=(EmployeeBillableTotal(x) + EmployeeNonBillableTotal(x))%>&nbsp;</td>
					<td align="right"><%if EmployeeCostTotal(x)<>"" then Response.write formatNumber(EmployeeCostTotal(x),2) else Response.Write "0.00"%>&nbsp;</td>
					<td align="right"><%=formatNumber(EmployeeRateTotal(x),2)%>&nbsp;</td>
				</tr>			
<%		end if
	next
%>
			</table>

			<br>

			<table border="0" cellpadding="2" cellspacing="2" width="600">
				<tr bgcolor="<%=gsColorHighlight%>">
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Phase")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Non-Billable_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Total_Hours")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Cost")%></td>
					<td align="center" class="columnheader" nowrap><%=dictLanguage("Value")%></td>
				</tr>		
<%	strBGColor = "#FFFFFF"
	for x=0 to 150
		if PhaseBillableTotal(x) > 0 or PhaseNonBillableTotal(x) > 0 then 
			if strBGColor = "#FFFFFF" then
				strBGColor = "#EEEEEE"
			else
				strBGColor = "#FFFFFF"
			end if	%>
				<tr bgcolor="<%=strBGColor%>">
					<td nowrap><%=phase(x)%>&nbsp;</font></td>
					<td align="right"><%=PhaseBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=PhaseNonBillableTotal(x)%>&nbsp;</td>
					<td align="right"><%=(PhaseBillableTotal(x) + PhaseNonBillableTotal(x))%>&nbsp;</td>
					<td align="right"><%=formatNumber(PhaseCostTotal(x),2)%>&nbsp;</td>
					<td align="right"><%=formatNumber(PhaseRateTotal(x),2)%>&nbsp;</td>
				</tr>			
<%		end if
	next
%>
			</table>
		</td>
		<td align="right" valign="top">

			<table border="0" cellpadding="2" cellspacing="2" style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
				<tr>
					<td>&nbsp;</td>
					<td class="bolddark"><%=dictLanguage("Billable_Total")%>:</td>
					<td align="right"><b><%=intBillableTotal%></b></td>
				</tr>
				<tr>
					<td>&nbsp;</td>	
					<td class="bolddark"><%=dictLanguage("Non-Billable_Total")%>:</td>
					<td align="right"><b><%=intNonBillableTotal%></b></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td class="bolddark"><%=dictLanguage("Grand_Total")%>:</td>
					<td align="right"><b><%=intGrandTotal%></b></td>	
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td class="bolddark"><%=dictLanguage("Total_Cost")%>:</td>
					<td align="right"><b><%if intCostTotal<>"" then Response.Write formatNumber(intCostTotal,2) else Response.Write "0.00"%></b></td>	
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td class="bolddark"><%=dictLanguage("Total_Value")%>:</td>
					<td align="right"><b><%=formatNumber(intRateTotal,2)%></b></td>	
				</tr>
			</table>
		</td>
	</tr>
</table>

<%	
End If
adoRS.Close
set adoRS = nothing
%>
<p align="center">
<a href="client_work_log.asp"><%=dictLanguage("Return_Work_Log_Form")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<% 
for each i in request.form
	session(i) = ""
next

function GetProjectPhaseName(id)
	sql = sql_GetProjectPhasesByID(id)
	set rsPhase = Conn.execute(sql)
	if not rsPhase.eof then
		rsPhase.movefirst
		phaseName = rsPhase("projectphaseName")
	else
		phaseName = "None"
	end if
	rsPhase.close
	set rsPhase = nothing
	GetProjectPhaseName = phaseName
end function
%>

<!--#include file="../includes/main_page_close.asp"-->