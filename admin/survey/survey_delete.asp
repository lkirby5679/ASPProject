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
<!--#include file="../../includes/main_page_open.asp"-->

<%
thispage = "survey_delete.asp"

poll = trim(request("pollid"))

sql = sql_GetAllSurveys()
Call RunSQL(sql, rsPolls)

if poll = "" then
	sql = sql_GetCurrentSurvey()
	Call RunSQL(sql, rsActive)
	if not rsActive.eof then
		response.write "<center>" & dictLanguage("Active_Survey_Is") & ": <br><b>" & rsActive("fdPollTopic") & "</b></center>"
	else
		response.write "<center>" & dictLanguage("No_Active_Survey") & "</center>"
	end if
	rsActive.close
	set rsActive = nothing 
	Response.Write "<form method=""post"" action=""" & thispage & """ id=form1 name=form1>"
	Response.Write "<br><center>"
	Response.Write "<Select Name=""pollid"" size=""1"" class=""formstylelong"">"
	Response.Write "<option value=''>--" & dictLanguage("Choose_Survey") & "--</option>"
	do While Not rsPolls.EOF
		Response.Write "<option value=""" & rsPolls("fdPollID") & """"
		if poll = trim(rsPolls("fdPollID")) then
			Response.Write " Selected"
		end if
		Response.Write ">" & rsPolls("fdPollTopic") & "</option>"
		rsPolls.MoveNext
	loop 
	rsPolls.close
	set rsPolls = nothing
	Response.Write "</Select>"
	Response.Write "<input type=""Submit"" value=""Delete"" class=""formButton"" onClick=""return confirm('" & dictLanguage("Confirm_Delete_Survey") & "')"" id=Submit1 name=Submit1>"
	Response.Write "</form>"
end if

if poll <> "" then
	sql = sql_GetQuestionsByPollID(poll)
	Call RunSQL(sql, rsQuestions)
	while not rsQuestions.eof
		sql = sql_DeleteSurveyAnswerByQuestionID(rsQuestions("fdPollQuestionID"))
		Call DoSQL(sql)
		rsQuestions.movenext
	wend
	rsQuestions.close
	set rsQuestions = nothing

	' Delete all the questions
	sql = sql_DeleteSurveyQuestionByPollID(poll)
	Call DoSQL(sql)

	' Delete the results
	sql = sql_DeleteSurveyResultsByPollID(poll)
	Call DoSQL(sql)

	' Delete the poll
	sql = sql_DeleteSurvey(poll)
	Call DoSQL(sql)

	response.write ("<center>" & dictLanguage("Survey_Is_Deleted") & "</center>")	
end if%>


<br>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>				
</p>


<!--#include file="../../includes/main_page_close.asp"-->
