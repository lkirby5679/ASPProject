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
id = Request.Form("id")
empID = Request.Form("empID")
''''''''''''''''''''''''''
'This page updates the time sheet record from edit.asp
''''''''''''''''''''''''''
for each i in Request.Form
	session(i) = Request.Form(i)
next

if session("datehourslogged")="" then
	Session("strErrorMessage") = session("strErrorMessage") & "<br>" & dictLanguage("Error_TimesheetNoDate")
elseif not isDate(session("datehourslogged")) then
	Session("strErrorMessage") = session("strErrorMessage") & "<br>" & dictLanguage("Error_TimesheetInvalidDate")
end if

If Session("strErrorMessage") <> "" then
	response.redirect "edit.asp?id=" & id & "&empID=" & empID & ""
End If

session("in1") = session("in1Hour") & ":" & session("in1Minute") & ":00 " & session("in1AMPM")
session("in2") = session("in2Hour") & ":" & session("in2Minute") & ":00 " & session("in2AMPM")
session("out1") = session("out1Hour") & ":" & session("out1Minute") & ":00 " & session("out1AMPM")
session("out2") = session("out2Hour") & ":" & session("out2Minute") & ":00 " & session("out2AMPM")

sql = sql_UpdateTimeSheets( _
	session("datehourslogged"), _
	session("in1"), _
	session("in2"), _
	session("out1"), _
	session("out2"), _
	id)
'Response.Write sql & "<BR>"
Call DoSQL(sql)

session("in1") = ""
session("in2") = ""
session("out1") = ""
session("out2") = ""
for each i in Request.Form
	session(i) = ""
next
%>

<!--#include file="../includes/connection_close.asp"-->

<%response.redirect "view2.asp?employee=" & empID & ""%>
