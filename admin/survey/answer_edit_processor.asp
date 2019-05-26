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
answer = trim(request("answer"))
answer_text = SQLEncode(trim(request("answer_text")))

if answer <> "" then
	sql = sql_UpdateAnswer(answer_text, answer)
	Call DoSQL(sql)
end if

response.redirect "survey_edit.asp?pollid=" & poll & ""
%>

<!--#include file="../../includes/connection_close.asp"-->
