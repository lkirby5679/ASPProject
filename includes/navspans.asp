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

<%if gsDHTMLNavigation then
	intSpanNumber = 0
	if gsTimecards then%>
<span name="navSubTimecards" id="navSubTimecards" style="background-color:<%=gsColorBackground%>; height:1; width:1; position: absolute; left: 0; top: 0; display: none; zIndex: 10;" onMouseOver="this.style.cursor='hand';">
	<table cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td nowrap bgcolor="<%=gsColorHighlight%>">
<%		if session("permTimecardsAdd") then %>
				<a href="<%=gsSiteRoot%>timecards/timecard.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("New_Timecard")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 
		end if %>
				<a href="<%=gsSiteRoot%>timecards/timecard-view.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_My_Timecards")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
				<a href="<%=gsSiteRoot%>reports/client_work_log.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Work_Log_Internal")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
				<a href="<%=gsSiteRoot%>reports/client_work_log_for_client.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Work_Log_For_Client")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
<%		if gsTimesheets then%>
				<a href="<%=gsSiteRoot%>timesheets/view.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Hourly_Timesheets")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
			</td>
		</tr>
	</table>
</span>
<%	end if%>
<%	if gsTasks then%>
<span name="navSubTasks" id="navSubTasks" style="background-color:<%=gsColorBackground%>; height:1; width:1; position: absolute; left: 0; top: 0; display: none; zIndex: 10;">
	<table cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td nowrap bgcolor="<%=gsColorHighlight%>">
<%		if session("permTasksAdd") then %>
				<a href="<%=gsSiteRoot%>tasks/new.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("New_Task")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 
		end if %>
				<a href="<%=gsSiteRoot%>tasks/view.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_Tasks")%></div></a>				
<%		intSpanNumber = intSpanNumber + 1 %>
				<a href="<%=gsSiteRoot%>tasks/view_mine.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_My_Tasks")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
				<a href="<%=gsSiteRoot%>tasks/view_my_assigned.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_Tasks_I_Assigned")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
			</td>
		</tr>
	</table>
</span>
<%	end if%>
<%	if gsProjects then%>
<span name="navSubProjects" id="navSubProjects" style="background-color:<%=gsColorBackground%>; height:1; width:1; position: absolute; left: 0; top: 0; display: none; zIndex: 10;">
	<table cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td nowrap bgcolor="<%=gsColorHighlight%>">
<%		if session("permProjectsAdd") then%>
				<a href="<%=gsSiteRoot%>projects/project-add.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("New_Project")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 
		end if %>
				<a href="<%=gsSiteRoot%>projects/project-view.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_Projects")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
				<a href="<%=gsSiteRoot%>projects/project-status.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Project_Status")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
<%		if gsProjectDashboard then%>
				<a href="<%=gsSiteRoot%>projects/dashboard.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Project_Dashboard")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 
		end if %>
			</td>
		</tr>
	</table>
</span>
<%	end if%>
<%	if gsClients then%>
<span name="navSubClients" id="navSubClients" style="background-color:<%=gsColorBackground%>; height:1; width:1; position: absolute; left: 0; top: 0; display: none; zIndex: 10;">
	<table cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td nowrap bgcolor="<%=gsColorHighlight%>">
<%		if session("permClientsAdd") then %>
				<a href="<%=gsSiteRoot%>clients/client-add.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("New_Client")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 
		end if %>
				<a href="<%=gsSiteRoot%>clients/client-view.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_Clients")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
				<a href="<%=gsSiteRoot%>clients/view_mine.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_My_Clients")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
<%		if gsClientContactLog then%>
				<a href="<%=gsSiteRoot%>clients/view_log.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_Client_Contact_Log")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
				<a href="javascript:popup('<%=gsSiteRoot%>clients/quickAdd.asp')"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Add_Log_Entry")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
			</td>
		</tr>
	</table>
</span>
<%	end if%>
<%	if gsReports then%>
<span name="navSubReports" id="navSubReports" style="background-color:<%=gsColorBackground%>; height:1; width:1; position: absolute; left: 0; top: 0; display: none; zIndex: 10;">
	<table cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td nowrap bgcolor="<%=gsColorHighlight%>">
<%		if gsProductionReportGrid then%>
				<a href="<%=gsSiteRoot%>reports/production_report_grid.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Production_Report_Grid")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
				<a href="<%=gsSiteRoot%>reports/client_work_log.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Work_Log_Internal")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
				<a href="<%=gsSiteRoot%>reports/client_work_log_for_client.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Work_Log_For_Client")%></div></a>
<%		intSpanNumber = intSpanNumber + 1 %>
<%		if gsTimeSheets then%>
				<a href="<%=gsSiteRoot%>timesheets/view.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Hourly_Timesheets")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>	
<%		end if%>
<%		if gsProjects then%>
				<a href="<%=gsSiteRoot%>projects/project-status.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Project_Status")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsClientContactLog then%>
				<a href="<%=gsSiteRoot%>clients/view_log.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_Client_Contact_Log")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsPTO and session("permPTOAdmin") then%>
				<a href="<%=gsSiteRoot%>pto/view.asp"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("View_PTO_Requests")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
			</td>
		</tr>
	</table>
</span>
<%	end if%>
<%	if gsOther or gsPTO then%>
<span name="navSubOther" id="navSubOther" style="background-color:<%=gsColorBackground%>; height:1; width:1; position: absolute; left: 0; top: 0; display: none; zIndex: 10;">
	<table cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td nowrap bgcolor="<%=gsColorHighlight%>">
<%		if gsCalendar then%>
				<a href="javascript: popupcalendar('<%=gsSiteRoot%>calendar/calendar.asp');"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Calendar")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsDiscussion then%>
				<a href="<%=gsSiteRoot%>forum/"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Discussion_Forum")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsEmployeeDirectory then%>
				<a href="<%=gsSiteRoot%>employees/"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Employee_Directory")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsFileRepository then %>
				<a href="<%=gsSiteRoot%>repository/"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Document_Repository")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsInventory then 'and permAdminInventory then%>
				<a href="<%=gsSiteRoot%>inventory/"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Inventory")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsNews then%>
				<a href="javascript: popup('<%=gsSiteRoot%>news/');"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("News")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsPTO then%>
				<a href="<%=gsSiteRoot%>pto/"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("PTO")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		if gsResources then%>
				<a href="javascript: popup('<%=gsSiteRoot%>resources/');"><div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Resources")%></div></a>
<%			intSpanNumber = intSpanNumber + 1 %>
<%		end if%>
<%		'if gsOptions then%>
				<!--div class="nav" name="subnav" id="subnav" onMouseOver="javascript: rollOnSubNav(<%=intSpanNumber%>);" onMouseOut="javascript: rollOffSubNav(<%=intSpanNumber%>);"><%=dictLanguage("Options")%></div-->
<%		'	intSpanNumber = intSpanNumber + 1 %>
<%		'end if%>
			</td>
		</tr>
	</table>
</span>
<%	end if
  end if%>