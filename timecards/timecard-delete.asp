<%@ LANGUAGE="VBSCRIPT" %>
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
<!--#include file="../includes/main_page_header.asp"-->

<%
tID = request("timecard_id")
if tID <> "" then
	sql = sql_DeleteTimecard(tID)
	Call DoSQL(sql)
end if
%>

<!--#include file="../includes/main_page_open.asp"-->

<p align="center">
<%=dictLanguage("Deleted_Timecard")%>
<br><br>
<a href="timecard-view.asp"><%=dictLanguage("View_My_Timecards")%></a><br>
<a href="timecard.asp"><%=dictLanguage("Add_Timecard")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<!--#include file="../includes/main_page_close.asp"-->



