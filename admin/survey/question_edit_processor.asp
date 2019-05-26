<%@language="VBSCript"%>
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
<!--#include file="../../includes/main_page_header.asp"-->

<% 	
poll = trim(request("pollid"))
question = trim(request("question"))
question_text = SQLEncode(trim(request("question_text")))
int_type = trim(request("type"))
old_type = trim(request("old_type"))

if question <> "" then
	sql = sql_UpdateSurveyQuestion(question_text, int_type, question)
	call DoSQL(sql)

	if old_type = 3 and int_type <> 3 then
		sql = sql_DeleteSurveyAnswerByQuestionID(question)
		Call DOSQL(sql)
	end if		
	if int_type = 3 or int_type=0 then
		sql = sql_DeleteSurveyAnswerByQuestionID(question)
		Call DoSQL(sql)
	end if
	if int_type = 3 then	
		sql = sql_InsertPollAnswer(question, 0, 0)
		Call DoSQL(sql)
	end if

end if
response.redirect "survey_edit.asp?pollid=" & poll & ""
%>

<!--#include file="../../includes/connection_close.asp"-->