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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:  This page is to be used as an include page.  The
'				following variables should be defined in the parent.
' ThisPage = The name of the parent in which the poll is included.
' LimitNumofVotes = TRUE/FALSE:  If set to True, limits votes by IP 
'				Addresses and by setting a Cookie Overturns RecordInCookie if it 
'				is set to False
' RecordInCookie = TRUE/FALSE:  If set to True, creates cookie without limiting
'				votes if limitNumofVotes is set to False
' WindowSize:#	Determines the size of the window in which the poll is displayed.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

if request("mode") = "vote" then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Vote Mode
'	Records the results of the submitted form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if request("polltopic") = "" then
		response.redirect (Request.ServerVariables("HTTP_REFERER"))
	else
		polltopic = trim(request("polltopic"))
	end if
	
	SQL = sql_GetQuestionsByPollID(polltopic)
	Call RunSQL(sql, rs)
	
	do while not rs.eof
		answerID = request("Question" & rs("fdPollQuestionID"))
		if answerID <> "" then
			if rs("fdPollQuestionType") = 1 then
				sql = sql_UpdatePollResults(answerID)
				Call DoSQL(sql)
			elseif rs("fdPollQuestionType") = 2 then
				ArrayQuestions = split(answerID, ",")
				for i = 0 to ubound(ArrayQuestions)
					sql = sql_UpdatePollResults(trim(ArrayQuestions(i)))
					Call DoSQL(sql)
					response.write ArrayQuestions(i)
				next
			end if
			sql = sql_UpdatePollQuestionResults(rs("fdPollQuestionID"))
			Call DoSQL(sql)
			answers = TRUE
		end if
		PollResult = PollResult & rs("fdPollQuestionID") & ":" & answerID & "|"
	rs.movenext
	loop
	
	rs.close
	set rs = nothing
	
	if answers <> TRUE then
		response.redirect (Request.ServerVariables("HTTP_REFERER"))
	end if
		sql = sql_InsertPollResults( _
			polltopic, _
			Request.ServerVariables("REMOTE_ADDR"), _
			date(),  _
			SQLEncode(PollResult), _
			SQLEncode(request("recpt")))
		Call DoSQL(sql)
		
		'msg = VBCRLF & "Thank you for taking the time to fill out our survey."
		'Call SendEmail("", request("recpt"), "", "dan@diggersolutions.com", _
		'	"Intranet Survey", _
		'	msg, _
		'	"", "", "", "", "", FALSE)
		
	if RecordInCookie = TRUE or LimitNumofVotes = TRUE then
		Response.Cookies("poll" & polltopic) = date()
		Response.Cookies("poll" & polltopic).expires = #12/31/10 00:00:00#
	end if
	
	response.redirect thispage & "?mode=results"

else
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Default Mode
'	Dynamically generates the form to be submitted or
'	displays results
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if request("pollid") <> "" then
		PollQuestion = request("pollid")
	elseif PollQuestion = "" then
		PollQuestion = 1
	end if
	
	if LimitNumofVotes = TRUE then
		sql = sql_CheckVoterIPByPollID(Request.ServerVariables("REMOTE_ADDR"), PollQuestion)
		'Response.Write sql & "<BR>"
		Call RunSQL(sql, rsCheckIP)
		if not rsCheckIP.eof then
			IPFound = TRUE
		end if
		if request.Cookies("poll" & PollQuestion) <> "" then
			CookieFound = TRUE
		end if
		rsCheckIP.close
		set rsCheckIP = nothing
	end if
	
	sql = sql_GetPollByID(PollQuestion)
	Call RunSQL(sql, rs)
	if not rs.eof then
		sql = sql_GetQuestionsByPollID(PollQuestion)
		Call RunSQL(sql, rsQuestions)
		
		Response.Write "<table border=""0"" cellpadding=""2"" cellspacing=""2"" align=""center"">"
		if not rsQuestions.eof and not rsQuestions.bof then
		  	if IPFound <> TRUE AND CookieFound <> TRUE AND request("mode") <> "results" then
				Response.Write "<form method=""post"" action=""" & thispage & "?mode=vote"" id=form1 name=form1>"
				Response.Write "<input type=""hidden"" name=""recpt"" value=""" & server.HTMLEncode(request("recpt")) & """>"
				Response.Write "<input type=""hidden"" name=""polltopic"" value=""" & rs("fdPollID") & """>"
			end if
			do while not rsQuestions.eof
				TotalVotes = rsQuestions("fdPollQuestionResult")
				Response.Write "<tr><td class=""homecontent"">&quot;" & rsQuestions("fdPollQuestion") & "&quot;"
				if ((request("mode") = "results") or IPFound or CookieFound) then
					if rsQuestions("fdPollQuestionType")<>"0" then
						Response.Write "<font class=""small""> - " & rsQuestions("fdPollQuestionResult") & " " & dictLanguage("Replies") & "</font>"
					end if
				end if
				Response.Write "</td></tr>"
				Response.Write "<tr><td class=""smallhome"">"
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
					rsAnswers.movefirst
				end if
				do while not rsAnswers.eof
					VoteForAnswer = rsAnswers("fdPollAnswerResult")
					if TotalVotes > 0 and VoteForAnswer > 0 then
						Percentage = formatnumber(100 * (VoteForAnswer/TotalVotes), 0)
					end if
					if IPFound <> TRUE AND CookieFound <> TRUE AND request("mode") <> "results" then
						if rsQuestions("fdPollQuestionType") = 1 then
							Response.Write "<input type=""Radio"" name=""Question" & rsQuestions("fdPollQuestionID") & """ value=""" & rsAnswers("fdPollAnswerID") & """>"
						elseif rsQuestions("fdPollQuestionType") = 2 then
							Response.Write "<input type=""CheckBox"" name=""Question" & rsQuestions("fdPollQuestionID") & """ value=""" & rsAnswers("fdPollAnswerID") & """>"
						elseif rsQuestions("fdPollQuestionType") = 3 then
							Response.Write "<textarea name=""Question" & rsQuestions("fdPollQuestionID") & """ rows=""3"" class=""formstylelong""></textarea>"
						end if					
					end if
					Response.Write rsAnswers("fdPollAnswer")
					if (IPFound = TRUE OR CookieFound = TRUE OR request("mode") = "results" ) AND Percentage <> "" then
						bool_showGraph = TRUE
						Response.Write "&nbsp;" & Percentage & "%<br>&nbsp;&nbsp;&nbsp;"
						cntShowImage = Int(Percentage/4)
						if cntShowImage < 1 then
							cntShowImage = 1
						end if
						for i = 1 to cntShowImage
							Response.Write "<img src=""" & gsSiteRoot & "images/graph_unit.gif"" border=""0"" valign=""middle"" WIDTH=""5"" HEIGHT=""5"">"
						next
					Percentage = ""
					end if
					if (int_maxLength <> 0 and int_maxLength < 3 and bool_showGraph <> TRUE) then
					else
						response.write "<br>"
					end if
				    rsAnswers.movenext
				loop
				rsAnswers.close 
				set rsAnswers = nothing 
				Response.Write "</td></tr>"
				rsQuestions.movenext
			loop
			Response.write "</table>"
			if IPFound <> TRUE AND CookieFound <> TRUE AND request("mode") <> "results" then
				Response.Write "<table border=""0"" align=""center""><tr>"
				Response.Write "<td><input type=""Submit"" value=""Vote"" id=""Submit1"" name=""Submit1"" class=""formButton""></td></form>"
				Response.Write "<form method=""post"" action=""" & thispage & "?mode=results"" id=""form2"" name=""form2"">"
				Response.Write "<td align=""center""><input type=""Submit"" value=""Results"" id=""Submit2"" name=""Submit2"" class=""formButton""></td></form>"
				Response.Write "</tr></table>"
			end if
		end if
		rsQuestions.close
		set rsQuestions = nothing
		rs.close
		set rs = nothing
	else
		response.write "<br>" & dictLanguage("No_Active_Survey")
	end if
end if%>
