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
''''''''''''''''''''''''''''''''''''''''''''''
'session-ize the form variables
''''''''''''''''''''''''''''''''''''''''''''''
for each i in request.form
	session(i) =  request.form(i)
next

from = session("from")

If session("Client_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoClient")
Else
	intClientID = session("Client_ID")
	intClientID = mid(intClientID,1,inSTR(intClientID,"P") - 1)
	sql = sql_GetActiveClientsByID(intClientID)
	Call RunSQL(sql, rsClient)
	if rsClient.EOF then 'client not active
		session("strErrorMessage") = session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardInvalidClient")
	end if
	rsClient.close
	set rsClient = nothing
End If

If session("TimeCardType_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoWorkType")
End If

If session("TimeAmount") = "" or IsNumeric(session("TimeAmount")) = False then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoTime")
End If

If session("DateWorked") = "" or IsDate(session("DateWorked")) = False then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoDate")
End If

If Month(session("DateWorked")) <> Month(Date) or Year(session("DateWorked")) <> Year(Date) then
	'Session("strErrorMessage") = Session("strErrorMessage") & "<br>You Must Enter a Date within the current Month"
End if

If session("WorkDescription") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_TimecardNoDesc")
End If

If Session("strErrorMessage") <> "" then
	response.redirect "timecard.asp"
End If

if session("projectphaseID")="" then
	projectphaseID = 0
else 
	projectphaseID = session("projectphaseID")
end if

strEmployeeID = session("employee_id")

sql	= sql_GetEmployeesWithTypeByID(strEmployeeID)
'Response.Write sql & "<BR>"
Call runSQL(sql, rsUserInfo)
If Not rsUserInfo.EOF then
	lngRate = rsUserInfo("Rate")
	lngCost = rsUserInfo("Cost")
End If
rsUserInfo.close
set rsUserInfo = nothing
	
if instr(session("Client_ID"),"P") then
	temp=split(session("Client_ID"),"P")
	Client_ID = temp(0)
	Project_ID = temp(1)
	if instr(Project_ID,"T") then
		temp2=split(Project_ID,"T")
		Project_ID = temp2(0)
		Task_ID = temp2(1)
	else
		Task_ID = 0
	end if
else
	Client_ID = session("Client_ID")
	Project_ID = 0
	Task_ID = 0
end if

NonBillable = 0

lngTimeCardRate = lngRate * session("TimeAmount")
lngTimeCardCost = lngCost * session("TimeAmount")

sql = sql_InsertTimecard( _
	Client_ID, _
	Project_ID, _
	Task_ID, _
	projectPhaseID, _
	strEmployeeID, _
	session("TimeCardType_ID"), _
	session("TimeAmount"), _
	SQLEncode(session("WorkDescription")), _
	date(), _
	session("DateWorked"), _
	date(), _
	session("Employee_ID"), _
	lngTimeCardRate, _
	lngTimeCardCost, _
	0, _
	NonBillable)
'response.write sql
Call DoSQL(sql)

'''''''''''''''''
'clear out the form's session variables so the
'form is clear when they return
''''''''''''''
for each i in request.form
	session(i) = ""
next
%>

<!--#include file="../includes/connection_close.asp"-->

<%
Session("strErrorMessage") = dictLanguage("Added_Timecard")
if from = "tasks" then
	response.redirect gsSiteRoot & "tasks/view_mine.asp"
else
	response.redirect "timecard.asp"
end if
%>

