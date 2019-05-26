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

if session("Paid")=1 or session("Paid") then
	session("Paid") = 1
else
	session("Paid") = 0
end if

If request.form("employee") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & dictLanguage("PTO_Error1") & ""
End If
	
If Session("strErrorMessage") <> "" then
	response.redirect "ptoform.asp"
End If

if Session("action") = "new" then
	sqlPTO = sql_InsertPTO(Session("employee"), Session("Start_Date"), Session("Start_Time"), _
		Session("End_Date"), Session("End_Time"), session("Total_Hours"), _
		Session("Paid"), session("Balance"), 1, Session("Reason"))
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

elseif Session("action") = "update" then
	pto_ID = Session("pto_ID")
	sqlPTO = sql_UpdatePTO(Session("Employee"), Session("Start_Date"), _
		Session("Start_Time"), Session("End_Date"), Session("End_Time"), _
		Session("Total_Hours"), Session("Paid"), Session("Balance"), Session("Reason"), 0, 0, 1, pto_ID)
	Call DoSQL(sqlPTO)

end if
'''''''''''''''''''''''''''''''''''''''''''''
' Get info about employee
'''''''''''''''''''''''''''''''''''''''''''''
sqlInfo	= sql_GetEmployeesByID(Session("employee"))
'response.write(sqlInfo)
Call RunSQL(sqlInfo, rsInfo)
if not rsInfo.eof then
	emailAddress = rsInfo("EMailAddress")
	empName = rsInfo("EmployeeName")
	reportsto = rsInfo("ReportsTo")
end if
rsInfo.close
set rsInfo = nothing

if reportsto <> "" then
	sqlManager = sql_GetEmployeesByID(reportsto)
	Call RunSQL(sqlManager, rsManager)
	if not rsManager.eof then
		managerName = rsManager("EmployeeName")
		managerEmail = rsManager("EMailAddress")
	end if
	rsManager.close
	set rsManager = nothing
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'send email to manager
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if managerEmail <> "" then
	strBody = empName & dictLanguage("PTO_Notice1") & Session("Start_Time") & ":" & Session("Start_Date") & " - " & Session("End_Time") & ":" & Session("End_Date") & "."
	strBody = strBody & dictLanguage("PTO_Notice2") & gsSiteURL & "/pto/ptoapproval.asp?pto_ID=" & pto_ID & dictLanguage("PTO_Notice3") & ""
	Call SendEmail(managerName, managerEmail, gsAdminEmail, gsAdminEmail, "PTO System", strBody, "", "", "", "", "", FALSE)
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

