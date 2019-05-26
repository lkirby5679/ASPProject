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
if Request.Form("Submit") = "Delete" then
	response.redirect "timecard-delete.asp?timecard_id=" & request.Querystring("timecard_id")
end if

''''''''''''''''''''''''''''''''''''''
'Validate Form Data
''''''''''''''''''''''''''''''''''''''

If request.form("Client_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoClient")
End If

If request.form("TimeCardType_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoWorkType")
End If

If request.form("TimeAmount") = "" or _
   IsNumeric(request.form("TimeAmount")) = False then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoTime")
End If

If request.form("DateWorked") = "" or _
   IsDate(request.form("DateWorked")) = False then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoDate")
End If

'New function - 4-26-99 DSH
If ((Month(request.form("DateWorked")) <> Month(Date) and DateDiff("d", request.form("DateWorked"), Date())>1) or Year(request.form("DateWorked")) <> Year(Date)) then
	'Session("strErrorMessage") = Session("strErrorMessage") & "<br>You Must Enter a Date within the current Month"
End if

If request.form("WorkDescription") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoDesc")
End If

If Request.Form("Employee") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoEmp")
End If

If Session("strErrorMessage") <> "" then
	response.redirect "timecard-edit.asp?timecard_id=" & request.querystring("timecard_id")
End If

Reconciled = 0    'Reconciled defaults to False
if Request.Form("Reconciled") = "yes" then
	Reconciled = 1
end if

employee_id = Request.Form("Employee")
sql = sql_GetEmployeesWithTypeByID(employee_id)
Call RunSQL(sql, rsUserInfo)
If Not rsUserInfo.EOF then
	lngRate = rsUserInfo("Rate")
	lngCost = rsUserInfo("Cost")
End If
rsUserInfo.close
set rsUserInfo = nothing

if instr(request.form("Client_ID"), "P") then
	clientProject=split(Request.Form("Client_ID"),"P")
	Client_ID = clientProject(0)
	Project_ID = clientProject(1)
	if instr(Project_ID,"T") then
		clientProject2=split(Project_ID,"T")
		Project_ID = clientProject2(0)
		Task_ID = clientProject2(1)
	else
		Task_ID = 0
	end if
else
	Client_ID = Request.Form("Client_ID")
	Project_ID = 0
	Task_ID = 0
end if

NonBillable = 0    'Non-Billable defaults to False
if Request.Form("Non-Billable") = "yes" then
	NonBillable = 1 
end if

if Client_ID = 1 and gsHostNonBillable then
	NonBillable = 1
end if

if request.form("projectphaseID")="" then
	projectphaseID = 0
else 
	projectphaseID = request.form("projectphaseID")
end if
	
lngTimeCardRate = lngRate * request.form("TimeAmount")
lngTimeCardCost = lngCost * request.form("TimeAmount")

sql = sql_UpdateTimecards( _
	Client_ID, _
	Project_ID, _
	Task_ID, _
	Employee_ID, _
	request.form("TimeCardType_ID"), _
	request.form("TimeAmount"), _
	SQLEncode(request.form("WorkDescription")), _
	request.form("DateWorked"), _
	date(), _
	session("Employee_ID"), _
	lngTimeCardRate, _
	lngTimeCardCost, _
	projectPhaseID, _
	NonBillable, _
	Reconciled, _
	request("timecard_id"))
'response.write(sql)
Call DoSQL(sql) %>

<!--#include file="../includes/main_page_open.asp"-->

<p align="center">
<%=dictLanguage("Updated_Timecard")%><br><br>
<a href="timecard-view.asp"><%=dictLanguage("View_My_Timecards")%></a><br>
<a href="timecard.asp"><%=dictLanguage("Add_Timecard")%></a><br>
<a href="<%=gsSiteRoot%>reports/client_work_log.asp"><%=dictLanguage("Work_Log_Internal")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<!--#include file="../includes/main_page_close.asp"-->
