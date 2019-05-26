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
if session("userid") = ""  or session("employee_id")="" then
	response.redirect gsSiteRoot & "default.asp"
end if
session.Timeout = 1200
if session("logonstatus") <> 1 then
      msg = "<center>" & dictLanguage("Error_NotLoggedIn") & "</center>"
      session("msg") = msg
      Response.Redirect gsSiteRoot & "default.asp"
end if
%>