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
	<tr><td align="center" bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("Reports")%></b></td></tr>
	<tr><td>&nbsp;</td></tr>
<%	if gsProductionReportGrid then%>
	<tr><td nowrap><a href="<%=gsSiteRoot%>reports/production_report_grid.asp"><%=dictLanguage("Production_Report_Grid")%></a></td></tr>
<%	end if%>
	<tr><td nowrap><a href="<%=gsSiteRoot%>reports/client_work_log.asp"><%=dictLanguage("Work_Log_Internal")%></a></td></tr>
	<tr><td nowrap><a href="<%=gsSiteRoot%>reports/client_work_log_for_client.asp"><%=dictLanguage("Work_Log_For_Client")%></a></td></tr>
<%	if gsTimeSheets then%>
	<tr><td nowrap><a href="<%=gsSiteRoot%>timesheets/view.asp"><%=dictLanguage("Hourly_Timesheets")%></a></td></tr>
<%	end if%>
<%	if gsProjects then%>
	<tr><td nowrap><a href="<%=gsSiteRoot%>projects/project-status.asp"><%=dictLanguage("Project_Status")%></a></td></tr>
<%	end if%>
<%	if gsClientContactLog then%>
	<tr><td nowrap><a href="<%=gsSiteRoot%>clients/view_log.asp"><%=dictLanguage("View_Client_Contact_Log")%></a></td></tr>
<%	end if%>
<%	if gsPTO and session("permPTOAdmin") then %>
	<tr><td nowrap><a href="<%=gsSiteRoot%>pto/view.asp"><%=dictLanguage("View_PTO_Requests")%></a></td></tr>
<%	end if %>
	<tr><td colspan="3">&nbsp;</td></tr>
	<td colspan="3" align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></td></tr>
</table>

<!--#include file="../includes/main_page_close.asp"-->