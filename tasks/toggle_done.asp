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
<!--#include file="../includes/SiteConfig.asp"-->
<!--#include file="../includes/mail.asp"-->
<!--#include file="../includes/connection_open.asp"-->

<%''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This page toggles an employee's task to "done" or "not done".
'It then sends an email informing the producer that the task is done
'and then shuttles them back form whence they came
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Toggle task completion field
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if request.querystring("value") = "no" then
	boolDone = "1"
else
	boolDone = "0"
end if
sql = sql_CompleteTasks(request.querystring("id"), boolDone)
'response.write sql
Call DoSQL(sql)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Collect all the info necessary for the mailing and build the email
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql = sql_GetTaskDoneEmailInfoByProjectID(request.querystring("id"))
Call RunSQL(sql, rsMailInfo) 
if not rsMailInfo.eof then
	strWorker = rsMailInfo("worker")
	if rsMailInfo("Done") = True then
		strMsg = strMsg & dictLanguage("has_completed")
		strSubject = dictLanguage("Intranet_Task") & ": " & dictLanguage("Task_Completed") 
	else
		strMsg = strMsg & dictLanguage("has_reopened")
		strSubject = dictLanguage("Intranet_Task") & ": " & dictLanguage("Task_Reopened")
	end if
	strMsg = strMsg & "task #" & rsMailInfo("task_id") & ":" & CHR(10) & CHR(13) & CHR(10)
	strMsg = strMsg & rsMailInfo("client_name") & "--" & CHR(10)
	strMsg = strMsg & CHR(9) & rsMailInfo("description") & CHR(10) & CHR(13) & CHR(10)
	if boolDone = "1" then
		strMsg = strMsg & dictLanguage("Task_Inst_4") & session("userid") & CHR(10) & CHR(13) & CHR(10)
	else 
		strMsg = strMsg & dictLanguage("Task_Inst_5") & session("userid") & CHR(10) & CHR(13) & CHR(10)
	end if

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'send the email and close all objects to free system resources
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if gsTaskEmails then
		Call SendEmail("", rsMailInfo("emailaddress"), "", rsMailInfo("worker_email"), _
			strSubject, strMsg, "", "", "", "", "", FALSE)
	end if
end if %>

<!--#include file="../includes/connection_close.asp"-->

<%response.redirect request.querystring("referring_page") & "?sql=" & Server.URLEncode(request("sql"))%>

