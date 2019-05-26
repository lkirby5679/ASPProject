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
<%
strUserid = sqlencode(request.form("userid"))
strPassword = sqlencode(request.form("password"))
%>

<!--#include file="includes/SiteConfig.asp"-->
<!--#include file="includes/connection_open.asp"-->

<%
if gsAutoLogin then
	if request.Cookies("shark_user_id") <> "" and request.Cookies("shark_password") <> "" and request.form("ccokie") <> "false" then
		strUserid = request.Cookies("shark_user_id")
		strPassword = request.Cookies("shark_password")
		boolFromCookie = True
	end if
end if

'-----------------------------------------------------
'-----------------------------------------------------
'Level 1 Basic Validation Testing
'-----------------------------------------------------
'-----------------------------------------------------

'Test for Entry of Member ID prior to Running Search
'-----------------------------------------------------
if (strUserID="") then ' Member ID Entered
  session("msg") = "<br><font size=2><center>" & dictLanguage("Error_NoUserID") & "</center></h3>"
  Response.Redirect "default.asp?opt=1"

'Run Search
'-----------------------------------------------------
else 'Run Database Lookup
  sql = sql_GetEmployeesByLogin(strUserID)
  'response.write sql & "<BR>"
  Set rs = Server.CreateObject("ADODB.RecordSet")
  rs.open sql, conn, 3, 3
end if

'-----------------------------------------------------
'-----------------------------------------------------
'Level 2 Validation Testing
'-----------------------------------------------------
'-----------------------------------------------------

'Member ID Entered Now Run Search for Record
'-------------------------------------------
If rs.eof Then 'No Record Found, Bad ID
  rs.close
  session("msg") = "<br><font size=2><center>" & dictLanguage("Error_NoUser") & "&nbsp;<em>" & strUserid & "</em>. " & dictLanguage("Please_Try_Again") & "</center></font>"
  Response.Redirect "default.asp?opt=1"

'Customer Record Found Now Check Password
'-----------------------------------------
elseif not Ucase(strPassword) = Ucase(rs("password")) then
	rs.close
  session("msg") = "<br><font size=2><center>" & dictLanguage("Error_InvalidPass") & "<br>" & dictLanguage("Please_Try_Again") & "</center></font>"
  Response.Redirect "default.asp?opt=1"

'Check for inactive status
'-------------------------------------------
elseif rs("active") <> 1 and not rs("active") then
     rs.close
     session("msg") = "<br><font size=2><center>" & dictLanguage("Error_InactiveAcct") & "</center></font>"
     Response.Redirect "default.asp?opt=1"

'-----------------------------------------------------
'-----------------------------------------------------
' Step 3. Success Logon Approved
'-----------------------------------------------------
'-----------------------------------------------------

else
	'check their permissions
	sql = sql_GetEmployeesPermissionsByID(rs("Employee_ID"))
	Call RunSQL(sql, rsPerms)
	if not rsPerms.eof then
		'give them the permissions they have been assigned
		if rsPerms("permAll")=1 or rsPerms("permAll") then
			'let them do everything
			session("permAdmin")				= TRUE
			session("permAdminCalendar")		= TRUE
			session("permAdminDatabaseSetup")	= TRUE
			session("permAdminEmployees")		= TRUE
			session("permAdminEmployeesPerms")	= TRUE
			session("permAdminFileRepository")	= TRUE
			session("permAdminForum")			= TRUE
			session("permAdminNews")			= TRUE
			session("permAdminResources")		= TRUE
			session("permAdminSurveys")			= TRUE
			session("permAdminThoughts")		= TRUE
		
			session("permClientsAdd")			= TRUE
			session("permClientsEdit")			= TRUE
			session("permClientsDelete")		= TRUE
			session("permProjectsAdd")			= TRUE
			session("permProjectsEdit")			= TRUE
			session("permProjectsDelete")		= TRUE
			session("permTasksAdd")				= TRUE
			session("permTasksEdit")			= TRUE
			session("permTasksDelete")			= TRUE
			session("permTimecardsAdd")			= TRUE
			session("permTimecardsEdit")		= TRUE
			session("permTimecardsDelete")		= TRUE
			session("permTimesheetsEdit")		= TRUE			
			session("permForumAdd")				= TRUE
			session("permRepositoryAdd")		= TRUE
			session("permPTOAdmin")				= TRUE
		elseif rsPerms("permAdminAll")=1 or rsPerms("permAdminAll") then 
			'give them permissions to all admin tasks
			session("permAdmin")				= TRUE
			session("permAdminCalendar")		= TRUE
			session("permAdminDatabaseSetup")	= TRUE
			session("permAdminEmployees")		= TRUE
			session("permAdminEmployeesPerms")	= TRUE
			session("permAdminFileRepository")	= TRUE
			session("permAdminForum")			= TRUE
			session("permAdminNews")			= TRUE
			session("permAdminResources")		= TRUE
			session("permAdminSurveys")			= TRUE
			session("permAdminThoughts")		= TRUE
		end if
		'get individual admin permissions
		session("permAdmin")					= rsPerms("permAdmin")
		session("permAdminCalendar")			= rsPerms("permAdminCalendar")
		session("permAdminDatabaseSetup")		= rsPerms("permAdminDatabaseSetup")
		session("permAdminEmployees")			= rsPerms("permAdminEmployees")
		session("permAdminEmployeesPerms")		= rsPerms("permAdminEmployeesPerms")
		session("permAdminFileRepository")		= rsPerms("permAdminFileRepository")
		session("permAdminForum")				= rsPerms("permAdminForum")
		session("permAdminNews")				= rsPerms("permAdminNews")
		session("permAdminResources")			= rsPerms("permAdminResources")
		session("permAdminSurveys")				= rsPerms("permAdminSurveys")
		session("permAdminThoughts")			= rsPerms("permAdminThoughts")
		
		session("permClientsAdd")				= rsPerms("permClientsAdd")
		session("permClientsEdit")				= rsPerms("permClientsEdit")
		session("permClientsDelete")			= rsPerms("permClientsDelete")
		session("permProjectsAdd")				= rsPerms("permProjectsAdd")
		session("permProjectsEdit")				= rsPerms("permProjectsEdit")
		session("permProjectsDelete")			= rsPerms("permProjectsDelete")
		session("permTasksAdd")					= rsPerms("permTasksAdd")
		session("permTasksEdit")				= rsPerms("permTasksEdit")
		session("permTasksDelete")				= rsPerms("permTasksDelete")
		session("permTimecardsAdd")				= rsPerms("permTimecardsAdd")
		session("permTimecardsEdit")			= rsPerms("permTimecardsEdit")
		session("permTimecardsDelete")			= rsPerms("permTimecardsDelete")
		session("permTimesheetsEdit")			= rsPerms("permTimesheetsEdit")
		session("permForumAdd")					= rsPerms("permForumAdd")
		session("permRepositoryAdd")			= rsPerms("permRepositoryAdd")
		session("permPTOAdmin")					= rsPerms("permPTOAdmin")
		
		for each i in session.Contents
			if session(i)<>"" then
				if len(session(i))>4 then
					if left(session(i),4) = "perm" then
						if session(i) = 1 or session(i) = -1 then
							session(i) = TRUE
						elseif session(i) = 0 then
							session(i) = FALSE
						end if
					end if
				end if
			end if
		next
		
		'check to see if they have any admin permissions at all, if so give them access to 
		'   the admin folder
		if session("permAdminCalendar") or session("permAdminDatabaseSetup") or _
			session("permAdminEmployees") or session("permAdminFileRepository") or _
			session("permAdminForum") or session("permAdminNews") or session("permAdminResources") or _
			session("permAdminEmployeesPerms") or session("permAdminSurveys") or session("permAdminThoughts") then
				session("permAdmin") = TRUE
		end if
	else
		session("permAdmin")					= FALSE
		session("permAdminCalendar")			= FALSE
		session("permAdminDatabaseSetup")		= FALSE
		session("permAdminEmployees")			= FALSE
		session("permAdminEmployeesPerms")		= FALSE
		session("permAdminFileRepository")		= FALSE
		session("permAdminForum")				= FALSE
		session("permAdminNews")				= FALSE
		session("permAdminResources")			= FALSE
		session("permAdminSurveys")				= FALSE
		session("permAdminThoughts")			= FALSE
		
		session("permClientsAdd")				= FALSE
		session("permClientsEdit")				= FALSE
		session("permClientsDelete")			= FALSE
		session("permProjectsAdd")				= FALSE
		session("permProjectsEdit")				= FALSE
		session("permProjectsDelete")			= FALSE
		session("permTasksAdd")					= FALSE
		session("permTasksEdit")				= FALSE
		session("permTasksDelete")				= FALSE
		session("permTimecardsAdd")				= TRUE   'we want them to be able to add time right?
		session("permTimecardsEdit")			= FALSE
		session("permTimecardsDelete")			= FALSE
		session("permTimesheetsEdit")			= FALSE		
		session("permForumAdd")					= FALSE
		session("permRepositoryAdd")			= FALSE
		session("permPTOAdmin")					= FALSE
	end if
	rsPerms.close
	set rsPerms = nothing
	
	session("logonstatus") = 1
	session("userid") = Ucase(rs("EmployeeLogin"))
	session("employee_id") = rs("Employee_ID")
	if gsAutoLogin then
		if not boolFromCookie then
			Response.Cookies("shark_user_id") = session("userid")
			Response.Cookies("shark_user_id").expires = #12/31/2010 00:00:00#
			Response.Cookies("shark_employee_id") = session("userid")
			Response.Cookies("shark_employee_id").expires = #12/31/2010 00:00:00#
			Response.Cookies("shark_password") = request.form("password")
			Response.Cookies("shark_password").expires = #12/31/2010 00:00:00#
		else
			session("cookieMessage") = dictLanguage("User") & " " & request.Cookies("shark_user_id") & " " & dictLanguage("logged_on") & " [" & Now() & "]<br><font size=-1>" & dictLanguage("click") & " <a href=" & gsSiteRoot & "killall.asp><font color=" & CHR(34) & "black" & CHR(34) & ">" & dictLanguage("here") & "</font></a> " & dictLanguage("logon_diff_user") & ".</font>"	
		end if
	end if
	sql = sql_UpdateEmployeesLastLogin(now(), strUserID)
	Call DoSQL(sql)
	rs.close
	response.redirect "main.asp"
end if
rs.close
%>
<!--#include file="includes/connection_close.asp"-->