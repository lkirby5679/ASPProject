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
<!--#include file="includes/main_page_header.asp"-->

<%
'''''''''''''''''''''''''''''''''''''''
'This page logs information to the timeSheets
'DB Table for hourly employees
'''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''
'Execute a DB query to bring back today's timesheet record
'for user
'''''''''''''''''''''''''''''''''''
sql = sql_GetTimesheetsByDate(Date(), session("employee_id"))
Call RunSQL(sql, rs)

'''''''''''''''''''''''''''''''''''
'If this is their initial log-on for today
'''''''''''''''''''''''''''''''''''
If rs.eof then
	sql = sql_InsertTimesheets( _
		session("employee_id"), _
		Date(), _
		Time(), _
		1)
	Call DoSQL(sql)
	session("loginAction") = "Login"

'''''''''''''''''''''''''''''''''''
'If they already have a record for today
'''''''''''''''''''''''''''''''''''
else
	sql = sql_UpdateTimesheetsNow( _
		rs("In2"), _
		rs("Out1"), _
		rs("Out2"), _
		session("employee_id")) 
	'Response.Write sql
	if sql <> "0" then
		Call DoSQL(sql)
	else
		session("error_message") = dictLanguage("Error_LoggedTwice") & "<br>"
		response.redirect("main.asp")
	end if

	sql = sql_InsertLoginAction( _
		session("loginAction"), _
		session("employee_id"), _
		Date(), _
		Request.ServerVariables("REMOTE_ADDR"))
	'response.write(sql)
	Call DoSQL(sql)
	session("loginAction") = ""
	
end if

rs.close
%>

<!--#include file="includes/connection_close.asp"-->

<% response.redirect("main.asp")%>
