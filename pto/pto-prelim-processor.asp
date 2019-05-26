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

If Session("strErrorMessage") <> "" then
	response.redirect "pto-prelimform.asp"
End If

sqlPTO = sql_InsertPTO(Session("employee"), Session("Start_Date"), Session("Start_Time"), _
		Session("End_Date"), Session("End_Time"), session("Total_Hours"), 0, 0, 0, "")
Call DoSQL(sqlPTO)

sql = sql_GetLatestPTO()
Call RunSQL(sql, rs)
if not rs.EOF then
	pto_ID = rs("max_pto")
end if
rs.close
set rs = nothing
if pto_ID = "" then
	pto_ID = 0
end if

'''''''''''''''''''''''''''''''''''''''''''''
' Get info about employee
'''''''''''''''''''''''''''''''''''''''''''''
sqlInfo	= sql_GetEmployeesByID(Session("employee"))
Call RunSQL(sqlInfo, rsInfo)
if not rsInfo.eof then	
	emailAddress = rsInfo("EMailAddress")
	empName = rsInfo("EmployeeName")
end if
rsInfo.close
set rsInfo = nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'send email to person pto is for
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
strBody = empName & dictLanguage("PTO_PrelimNotice1") & gsSiteURL & "/pto/ptoform.asp?pto_ID=" & pto_ID & " "
strBody = strBody & dictLanguage("PTO_PrelimNotice2") & Session("Start_Date") & " - " & Session("End_Date") & "."

Call SendEmail(empName, emailAddress, gsAdminEmail, gsAdminEmail, "PTO System", strBody, "", "", "", "", "", FALSE)

for each i in request.form
	session(i) =  ""
next
%>

<table border="0" cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("Preliminary_PTO_Form")%></b></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td colspan="2" align="center"><%=dictLanguage("PTO_Thankyou1")%></td></tr>
</table>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<br>

<!--#include file="../includes/main_page_close.asp"-->
