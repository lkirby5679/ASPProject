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
<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->


<table border="0" cellspacing="2" cellpadding="2" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Surveys")%></td></tr>
	<tr><td><a href="survey.asp"><%=dictLanguage("Active_Survey")%></a></td></tr>
	<tr><td><a href="survey.asp?mode=results"><%=dictLanguage("Active_Survey_Results")%></a></td></tr>
	<tr><td><a href="view_surveys.asp"><%=dictLanguage("Detailed_Survey_Results")%></a></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td><a href="current_survey.asp"><%=dictLanguage("Change_Active_Survey")%></a></td></tr>
	<tr><td><a href="survey_add.asp"><%=dictLanguage("Add_Survey")%></a></td></tr>
	<tr><td><a href="survey_edit.asp"><%=dictLanguage("Edit_A_Survey")%></a></td></tr>
	<tr><td><a href="survey_delete.asp"><%=dictLanguage("Delete_Survey")%></a></td></tr>
</table>

<br><br>
<p align="Center"><a href=".."><%=dictLanguage("Return_Admin_Home")%></a><br></p>

<!--#include file="../../includes/main_page_close.asp"-->
