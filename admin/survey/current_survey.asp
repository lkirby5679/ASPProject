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
thispage = "current_survey.asp"
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
	Response.Write "<center><form method=""post"" action=""" & thispage & """>"
	Response.Write "<br>"
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
	Response.Write "</Select><input type=""Submit"" value=""Submit"" class=""formbutton""></form></center>"
end if

if poll <> "" then
	sql = sql_GetCurrentSurveyRecords()
	Call RunSQL(sql, rsActive)
	if not rsActive.eof then
		sql = sql_UpdateCurrentSurvey(poll)
		Call DoSQL(sql)
	else 
		sql = sql_InsertCurrentSurvey(poll)
		Call DoSQL(sql)
	end if
	rsActive.close
	set rsActive = nothing
	response.write "<center>" & dictLanguage("Survey_Is_Active") & "</center>"	
end if
%>


<br>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>				
</p>

<!--#include file="../../includes/main_page_close.asp"-->