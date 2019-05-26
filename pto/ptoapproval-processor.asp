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
''''''''''''''''''''''''''''''''''''''''''''''
'session-ize the form variables
''''''''''''''''''''''''''''''''''''''''''''''
for each i in request.form
	session(i) =  SQLEncode(request.form(i))
next

if session("Paid") = 1 or session("Paid") then
	session("Paid") = 1 
else 
	session("Paid") = 0
end if

if session("Approved") = 1 or session("Approved") then
	session("Approved") = 1 
else 
	session("Approved") = 0
end if

if session("Excused") = 1 or session("Excused") then
	session("Excused") = 1 
else 
	session("Excused") = 0
end if
	
If Session("strErrorMessage") <> "" then
	response.redirect "ptoform.asp"
End If

pto_ID = Session("pto_ID")

sqlPTO = sql_UpdatePTO(Session("Employee"), Session("Start_Date"), Session("Start_Time"), _
	Session("End_Date"), Session("End_Time"), Session("Total_Hours"), _
	Session("Paid"), Session("Balance"), Session("Reason"), _
	Session("Approved"), Session("Excused"), 1, pto_ID)
 Call DoSQL(sqlPTO)

'''''''''''''''''''''''''''''''''''''''''''''
' Get info about employee
'''''''''''''''''''''''''''''''''''''''''''''
sqlInfo	= sql_GetEmployeesByID(Session("employee"))
'response.write(sqlInfo)
Call RunSQL(sqlInfo, rsInfo)
if not rsInfo.eof then
	emailAddress = rsInfo("EMailAddress")
	empName = rsInfo("EmployeeName")
end if

if Session("Approved") = 1 then
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' add a timecard for person taking off that day
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sqlInsertTimeCard = sql_InsertTimecard(1, 0, 0, 0, session("Employee"), _
		0, session("Total_Hours"), "PTO", date(), session("Start_Date"), date(), _
		session("Employee"), 0, 0, 0, 1)
	Call DoSQL(sqlInsertTimeCard)

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'send approval email to person taking PTO  
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if emailAddress<>"" then
		strBody = empName & dictLanguage("PTO_Approval1")
		strBody = strBody & Session("Start_Date") & " - " & Session("Total_Hours") & " " 
		strBody = strBody & dictLanguage("PTO_Approval2")
		Call SendEmail(empName, emailAddress, gsAdminEmail, gsAdminEmail, "PTO System", strBody, "", "", "", "", "", FALSE)
	end if
else
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'send dissaproval email to person taking PTO  
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if emailAddress<>"" then	
		strBody = empName & dictLanguage("PTO_Approval3") & Session("Start_Time") & ":" & Session("Start_Date") & " - " & Session("End_Time") & ":" & Session("End_Date") & dictLanguage("PTO_Approval4")
		Call SendEmail(empName, emailAddress, gsAdminEmail, gsAdminEmail, "PTO System", strBody, "", "", "", "", "", FALSE)
	end if
end if

for each i in request.form
	session(i) =  ""
next
%>

<table border="0" cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("PTO_Form")%></b></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td colspan="2" align="center"><%=dictLanguage("PTO_Thankyou1")%></td></tr>
</table>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<br>

<!--#include file="../includes/main_page_close.asp"-->