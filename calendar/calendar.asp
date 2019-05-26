<%@ LANGUAGE="VBSCRIPT" %>
<Title>Calendar</title>
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
<!--#include file="../includes/popup_page_open.asp"-->

<%
strView = request("view")
if strView = "" then
	strView = "month"
end if
%>


<table cellpadding="2" cellspacing="2" border="0" width="100%" align="center" bgcolor="<%=gsColorHighlight%>">
	<tr>
		<td class="home" align="right">
			<a href="calendar.asp?view=month&date=<%=request("date")%>" class="homeheader">Month View</a>&nbsp&nbsp|
			<a href="calendar.asp?view=week&date=<%=request("date")%>" class="homeheader">Week View</a>&nbsp&nbsp|
			<a href="calendar.asp?view=day&date=<%=request("date")%>" class="homeheader">Day View</a>&nbsp&nbsp
		</td>
	</tr>
	<tr>
		<td>
<%	select case strView 
		case "week" %>
		<!--#include file="../includes/calendar_week.asp"-->
<%		case "day" %>
		<!--#include file="../includes/calendar_day.asp"-->
<%		case else %>
		<!--#include file="../includes/calendar_month.asp"-->
<%	end select %>
		</td>
	</tr> 
</table>	



<!--#include file="../includes/popup_page_close.asp"-->
