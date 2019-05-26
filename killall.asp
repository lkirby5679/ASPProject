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
   'This page resets session variables
   session.abandon

	Response.Cookies("shark_user_id") = ""
	Response.Cookies("shark_user_id").expires = #12/31/2000 00:00:00#
	Response.Cookies("shark_employee_id") = ""
	Response.Cookies("shark_employee_id").expires = #12/31/2000 00:00:00#
	Response.Cookies("shark_password") = ""
	Response.Cookies("shark_password").expires = #12/31/2000 00:00:00#

   newpage = "default.asp?opt=1"
   response.redirect newpage

%>