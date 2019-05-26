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
thispage = "survey_edit.asp"

poll = trim(request("pollid"))

sql = sql_GetAllSurveys()
Call RunSQL(sql, rsPolls)

if poll = "" then
	sql = sql_GetCurrentSurvey()
	Call RunSQL(sql, rsActive)
	if not rsActive.eof then
		response.write "<center>" & dictLanguage("Active_Survey_Is") & ": <br><b>" & rsActive("fdPollTopic") & "</b><br>" & dictLanguage("SurveyInst_3") & "</center>"
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
	Response.Write "<input type=""Submit"" value=""Edit"" id=Submit1 name=Submit1 class=""formbutton"">"
	Response.Write "</form>"
end if

if poll <> "" then
	sql = sql_GetPollByID(poll)
	Call RunSQl(sql, rsTopic)
	if not rsTopic.eof then
		sql = sql_GetQuestionsByPollID(poll)
		Call RunSQL(sql, rsQuestions)
		Response.Write "<center><br>"
		Response.Write "<a href=""question_add.asp?pollid=" & poll & """><b>" & dictLanguage("Add_Question") & "</b></a><br><br>"
		if not rsQuestions.eof then 
			Response.Write "<table border=""0"" cellpadding=""2"" cellspacing=""2"" align=""center"">"	
			do while not rsQuestions.eof
				Response.Write "<tr><td><a href=""question_edit.asp?question=" & rsQuestions("fdPollQuestionID") & "&pollid=" & poll & """>"
				Response.Write "<font color=""black""><b>&quot;" & rsQuestions("fdPollQuestion") & "&quot;</b></font></a></td></tr>"
				Response.Write "<tr><td>"
				sql = sql_GetAnswersByQuestionID(rsQuestions("fdPollQuestionID"))
				Call RunSQL(sql, rsAnswers)
				int_maxLength = 0
				if not rsAnswers.eof then
					while not rsAnswers.eof				
						if len(rsAnswers("fdPollAnswer")) > int_maxLength then
							int_maxLength = len(rsAnswers("fdPollAnswer"))
						end if
						rsAnswers.Movenext
					wend
					rsAnswers.MoveFirst
				end if 
				if rsQuestions("fdPollQuestionType") <> 3 then 
					Response.Write "<center><a href=""answer_add.asp?question=" & rsQuestions("fdPollQuestionID") & "&pollid=" & poll & """>"
					Response.Write dictLanguage("Add_Answer") & "</a></center>"
				end if
				do while not rsAnswers.eof
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					if rsQuestions("fdPollQuestionType") = 1 then
						Response.Write "<input type=""Radio"" name=""Question" & rsQuestions("fdPollQuestionID") & """ value=""" & rsAnswers("fdPollAnswerID") & """>"
					elseif rsQuestions("fdPollQuestionType") = 2 then
						Response.Write "<input type=""CheckBox"" name=""Question" & rsQuestions("fdPollQuestionID") & """ value=""" & rsAnswers("fdPollAnswerID") & """>"
					elseif rsQuestions("fdPollQuestionType") = 3 then
						Response.Write "<textarea name=""Question" & rsQuestions("fdPollQuestionID") & """ cols=""45"" rows=""3""></textarea>"
					end if 
					Response.Write "<a href=""answer_edit.asp?answer=" & rsAnswers("fdPollAnswerID") & "&pollid=" & poll & """><font color=""black"">"
					response.write rsAnswers("fdPollAnswer")
					response.write "</font></a>"
					if (int_maxLength <> 0 and int_maxLength < 3 and bool_showGraph <> TRUE) then
						'do nothing
					else
						response.write "<br>"
					end if
				    rsAnswers.movenext
				loop
				rsAnswers.close 
				Response.Write "</td></tr>"
				rsQuestions.movenext
			loop
			response.write "</table></center>"
		else
			response.write "<br>" & dictLanguage("No_Questions_Entered") & "<br>"
		end if
		rsQuestions.close
		set rsQuestions = nothing
	end if 
	rsTopic.close
	set rsTopic = nothing
end if%>


<br>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>				
</p>

<!--#include file="../../includes/main_page_close.asp"-->

