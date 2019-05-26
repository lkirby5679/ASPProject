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
thispage = "view_surveys.asp"
poll = trim(request("pollid"))
entry = trim(request("entry"))

sql = sql_GetAllSurveys()
Call RunSQL(sql, rsPolls)

if entry = "" then
	Response.Write "<form method=""post"" action=""" & thispage & """>"
	Response.Write "<br><br><center>"
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
	Response.Write "</Select><input type=""Submit"" value=""View"" class=""formbutton"">"
	Response.Write "</form>"
end if

if poll <> "" then
	sql = sql_GetSurveyResultsByPollID(poll)
	Call RunSQL(sql, rsPollResults) 
	Response.Write "<table border=""0"" cellpadding=""2"" cellspacing=""2"" align=""center"">"
	Response.Write "<tr><td width=""60""><b>" & dictLanguage("Entry") & "</b></td>"
	Response.Write "<td width=""90""><b>" & dictLanguage("Entry_Date") & "</b></td>"
	Response.Write "<td width=""110""><b>" & dictLanguage("Origin_IP_Address") & "</b></td>"
	Response.Write "<td width=""110""><b>" & dictLanguage("Email_Address") & "</b></td></tr>"
	do while not rsPollResults.eof
		Response.Write "<tr><td><a href=""" & thispage & "?entry=" & rsPollResults("fdPollResultsIPID") & """>" & rsPollResults("fdPollResultsIPID") & "</a></td>" 
		Response.Write "<td>" & rsPollResults("fdPollResultsIPDate") & "</td>"
		Response.Write "<td>" & rsPollResults("fdPollResultsIP") & "</td>" 
		Response.Write "<td>" & rsPollResults("fdPollResultsEmail") & "</td></tr>"
		rsPollResults.movenext
	loop
	Response.write "</table>"
end if

if entry <> "" then
	sql = sql_GetSurveyResultsByID(entry)
	Call RunSQL(sql, rsPollResult)	
	sql = sql_GetQuestionsByPollID(rsPollResult("fdPollID"))
	Call RunSQL(sql, rsQuestions)
	
	answers = split(rsPollResult("fdPollResults"), "|")
	answer_num = 0
	Response.Write "<table border=""0"" cellpadding=""2"" cellspacing=""2"" align=""center"">"
	Response.Write "<tr><td><b>" & dictLanguage("Entry_Date") & ":</b>" & rsPollResult("fdPollResultsIPDate") & "</td>"
	Response.Write "<td><b>" & dictLanguage("Origin_IP_Address") & ":</b>" & rsPollResult("fdPollResultsIP") & "</td></tr></table>"
	Response.Write "<table border=""0"" cellpadding=""2"" cellspacing=""2"" align=""center"">"
	if not rsQuestions.eof and not rsQuestions.bof then
		do while not rsQuestions.eof
			answer = right(answers(answer_num), len(answers(answer_num)) - instr(answers(answer_num), ":"))
			found = 0
			Response.Write "<tr><td><b>&quot;" & rsQuestions("fdPollQuestion") & "&quot;</b><br></td></tr>"
			Response.Write "<tr><td>" 
			sql = sql_GetAnswersByQuestionID(rsQuestions("fdPollQuestionID"))
			Call RunSQL(sql, rsAnswers)
			do while not rsAnswers.eof
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				if rsQuestions("fdPollQuestionType") = 1 then
					Response.Write "<input type=""Radio"" name=""Question" & rsQuestions("fdPollQuestionID") & """ value=""" & rsAnswers("fdPollAnswerID") & """"
					if trim(rsAnswers("fdPollAnswerID")) = trim(answer) then
						Response.Write " Checked"
					end if
					Response.Write ">"
				elseif rsQuestions("fdPollQuestionType") = 2 then
					multiple_answers = split(answer, ",") 
					for im = 0 to ubound(multiple_answers)
						if trim(rsAnswers("fdPollAnswerID")) = trim(multiple_answers(im)) then
							found = 1
						end if
					next
					Response.Write "<input type=""CheckBox"" name=""Question" & rsQuestions("fdPollQuestionID") & """ value=""" & rsAnswers("fdPollAnswerID") & """"
					if found = 1 then
						Response.Write " Checked"
					end if
					Response.Write ">"
					found = 0
				elseif rsQuestions("fdPollQuestionType") = 3 then
					Response.Write "<textarea name=""Question" & rsQuestions("fdPollQuestionID") & """ cols=""45"" rows=""3"">" & answer & "</textarea>"
				end if
				Response.Write rsAnswers("fdPollAnswer") & "<br>"
			rsAnswers.movenext
			loop	
			rsAnswers.close
			set rsAnswers = nothing
			Response.Write "</td></tr>"
			rsQuestions.movenext
		answer_num = answer_num + 1
		loop
		Response.Write "</table><br>"
	end if
	rsQuestions.close
	set rsQuestions = nothing
end if
%>		

<br>
<p align="Center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>			
</p>

<!--#include file="../../includes/main_page_close.asp"-->
