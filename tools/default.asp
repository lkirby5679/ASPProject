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
<!--#include file="../includes/main_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border=0 align="center">
	<tr><td align="center" bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("Tools")%></b></td></tr>
	<tr><td>&nbsp;</td></tr>
<%	if gsCalendar then%>
	<tr><td nowrap><a href="javascript: popup('<%=gsSiteRoot%>calendar/');"><%=dictLanguage("Calendar")%></a></td></tr>
<%	end if%>
<%	if gsDiscussion then%>
	<tr><td nowrap><a href="<%=gsSiteRoot%>forum/"><%=dictLanguage("Discussion_Forum")%></a>
<%	end if%>
<%	if gsEmployeeDirectory then%>
	<tr><td nowrap><a href="<%=gsSiteRoot%>employees/"><%=dictLanguage("Employee_Directory")%></a></td></tr>
<%	end if%>
<%	if gsFileRepository then %>
	<tr><td nowrap><a href="<%=gsSiteRoot%>repository/"><%=dictLanguage("Document_Repository")%></a></td></tr>
<%	end if%>
<%	if gsNews then%>
	<tr><td nowrap><a href="javascript: popup('<%=gsSiteRoot%>news/');"><%=dictLanguage("News")%></a></td></tr>
<%	end if%>
<%	if gsOptions then%>
	<tr><td nowrap><%=dictLanguage("Options")%></td></tr>
<%	end if%>
	<tr><td colspan="3">&nbsp;</td></tr>
	<td colspan="3" align="center"><a href="main.asp"><%=dictLanguage("Return_Business_Console")%></a></td></tr>
</table>

<!--#include file="../includes/main_page_close.asp"-->