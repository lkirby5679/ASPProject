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
<!--#include file="includes/main_page_header.asp"-->

<%
'find out who this person is
strEmpID = session("employee_id")
if strEmpID = "" then
	Response.Redirect "default.asp"
end if

if gsHomeThoughts then
	'get the random thought
	sql = sql_GetLeastUsedThought()
	Call RunSQL(sql,rsThought)
	if not rsThought.eof then
		strThought = rsThought("description")
		strThoughtId = rsThought("ID")
	else
		strThought = dictLanguage("Error_NoThoughtsExist")
	end if
	rsThought.close
	set rsThought = nothing
	if strThoughtID <> "" then
		sql = sql_IncrementThoughtByID(strThoughtID)
		Call DoSQL(sql)
	end if
end if

if gsHomeProductionReport then
	'figure team's monthly production goal
	sql = sql_GetTeamMonthlyProductionGoal()
	Call RunSQL(sql,rs)
	if not rs.eof then
		if rs("SumOfproductionGoal")<>"" then
			strTeamProductionGoal = FormatNumber(rs("SumOfproductionGoal"),0)
		else
			strTeamProductionGoal = 0
		end if
	else
		strTeamProductionGoal = 0
	end if
	rs.close
	set rs = nothing

	'figure team's current hours
	sql = sql_GetTeamMonthlyBillableHours(date())
	'Response.Write sql & "<BR>"
	Call RunSQL(sql,rs)
	if not rs.eof then
		if rs("WorkThisMonth")<>"" then
			strTeamWorkThisMonth = rs("WorkThisMonth")
		else
			strTeamWorkThisMonth = 0
		end if
	else
		strTeamWorkThisMonth = 0
	end if
	strTeamWorkThisMonth = FormatNumber(strTeamWorkThisMonth,2)
	rs.close
	set rs = nothing

	'figure personal production goal
	sql = sql_GetEmployeeMonthlyProductionGoal(strEmpID)
	Call runSQL(sql,rs)
	if not rs.eof then
		if rs("productionGoal")<>"" then
			strPersonalProductionGoal = rs("productionGoal")
		else
			strPersonalProductionGoal = 0
		end if
	else
		strPersonalProductionGoal = 0
	end if	
	rs.close
	set rs = nothing

	'figure personal current hours
	sql = sql_GetEmployeeMonthlyBillableHours(strEmpID, date())
	'Response.Write sql & "<BR>"
	Call RunSQL(sql,rs)
	if not rs.eof then
		if rs("WorkThisMonth")<>"" then
			strPersonalWorkThisMonth = rs("WorkThisMonth")
		else
			strPersonalWorkThisMonth = 0
		end if
	else
		strPersonalWorkThisMonth = 0
	end if
	strPersonalWorkThisMonth = FormatNumber(strPersonalWorkThisMonth,2)
	rs.close
	set rs = nothing

	'figure hours for today
	sql = sql_GetEmployeeDailyBillableHours(strEmpID, date())
	'response.write sql
	Call RunSQL(sql,rs)
	if not rs.eof then
		if rs("WorkToday")<>"" then
			strPersonalWorkToday = rs("WorkToday")
		else
			strPersonalWorkToday = 0
		end if
	else
		strPersonalWorkToday = 0
	end if
	strPersonalWorkToday = FormatNumber(strPersonalWorkToday,2)
	rs.close
	set rs = nothing

	'figure hours for yesterday
	sql = sql_GetEmployeeDailyBillableHours(strEmpID, dateadd("d", -1, date()))	
	set rs=conn.execute(sql)
	if not rs.eof then
		if rs("WorkToday")<>"" then
			strPersonalWorkYesterday = rs("WorkToday")
		else
			strPersonalWorkYesterday = 0
		end if
	else
		strPersonalWorkYesterday = 0
	end if
	strPersonalWorkYesterday = FormatNumber(strPersonalWorkYesterday,2)
	rs.close
	set rs = nothing
end if

if gsTimesheets then
	'is	the person hourly, if so then generate the timeclock
	'need to go back in and set up tbl_timesheets for this to work
	boolHourly = FALSE
	sql = sql_GetHourlyEmployeesByID(strEmpID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		boolHourly = TRUE
		sql = sql_GetTimesheetsByDate(date(), strEmpID)
		Call RunSQL(sql, rs2)
		if not rs2.eof then
			if rs2("LoggedIn") = "True" then
				strTimeClockValue = dictLanguage("Clock_Out")
			else
				strTimeClockValue = dictLanguage("Clock_In")
			end if
		else
			strTimeClockValue = dictLanguage("Clock_In")
		end if		
		rs2.close
		set rs2 = nothing
	end if
	rs.close
	set rs = nothing
end if
%>


<!--#include file="includes/main_page_open.asp"-->
	
<table border="0" width="100%" cellpadding="0" cellspacing="1">
	<tr>
		<td valign="top">
			<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=gsColorBackground%>" width="100%"><tr><td valign="top">		
<%	if gsHomeThoughts then %>
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>">
<%		if gsHomeMinimize then
			if Request.Cookies("thoughts")<>"FALSE" then 
				strDisplay = "block" %>
						<img align="right" name="thoughts_option" id="thoughts_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('thoughts');" onmouseover="this.style.cursor='hand'">
<%			else 
				strDisplay = "none" %>
						<img align="right" name="thoughts_option" id="thoughts_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('thoughts');" onmouseover="this.style.cursor='hand'">
<%			end if 
		end if %>
  						<b class="homeheader"><%=dictLanguage("THINK_OF_THIS")%>...</b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="thoughts" id="thoughts">
					<td class="homecontent"><%=strThought%></td>
				</tr>
			</table>
<%	end if %>
			</td></tr></table>			
			<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=gsColorBackground%>" width="100%"><tr><td valign="top">
<%	if gsHomeNews then %>
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>" colspan="2">
<%			if gsHomeMinimize then
				if Request.Cookies("news")<>"FALSE" then 
					strDisplay = "block" %>
						<img align="right" name="news_option" id="news_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('news');" onmouseover="this.style.cursor='hand'">
<%				else
					strDisplay = "none" %>
						<img align="right" name="news_option" id="news_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('news');" onmouseover="this.style.cursor='hand'">
<%				end if 
			end if %>
						<b class="homeheader"><%=dictLanguage("NEWS")%></b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="news" id="news">
					<td><!--#include file="includes/news.asp"--></td>
				</tr>
			</table>			
<%	end if %>						
			</td>
			<td valign="top" width="300">
<%	if gsHomeDiscussion then %>
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>">
<%		if gsHomeMinimize then 
			if Request.Cookies("discussion")<>"FALSE" then 
				strDisplay = "block" %>
						<img align="right" name="discussion_option" id="discussion_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('discussion');" onmouseover="this.style.cursor='hand'">
<%			else
				strDisplay = "none" %>
						<img align="right" name="discussion_option" id="discussion_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('discussion');" onmouseover="this.style.cursor='hand'">
<%			end if 
		end if %>
						<b class="homeheader"><%=dictLanguage("NEW_DISCUSSION_TOPICS")%></b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="discussion" id="discussion">
					<td>
						<table border="0" cellpadding="2" cellspacing="2" width="100%">
							<tr>
								<td class="homecontent"><!--#include file="includes/forums.asp"--></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>			
<%	end if %>			
			</td></tr></table>
			<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=gsColorBackground%>" width="100%"><tr><td valign="top">
<%  if gsHomeProductionReport then %>
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>" colspan="2">
<%		if gsHomeMinimize then
			if Request.Cookies("summary")<>"FALSE" then 
				strDisplay = "block" %>
						<img align="right" name="summary_option" id="summary_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('summary');" onmouseover="this.style.cursor='hand'">
<%			else 
				strDisplay = "none" %>
						<img align="right" name="summary_option" id="summary_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('summary');" onmouseover="this.style.cursor='hand'">
<%			end if
		end if %>
						<b class="homeheader"><%=dictLanguage("PRODUCTION_REPORT")%></b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="summary" id="summary">
					<td class="homecontent">
						<%=dictLanguage("Team_Monthly_Prod_Goal")%>:
						<b><%=strTeamProductionGoal%></b> <%=dictLanguage("Hours")%>.<br>
						<%=dictLanguage("We_Are_Currently_At")%> <b><%=strTeamWorkThisMonth%></b>.					
<%		if strTeamWorkThisMonth - strTeamProductionGoal =< 0 then%>
						<%=dictLanguage("We_Need")%> <b><%=FormatNumber(strTeamProductionGoal-strTeamWorkThisMonth,2)%></b> <%=dictLanguage("more")%>!<br>
<%		else%>
						<%=dictLanguage("We_Are")%> <b><%=FormatNumber(strTeamWorkThisMonth-strTeamProductionGoal,2)%></b> <%=dictLanguage("Over_Our_Goal")%>!<br>
<%		end if%>

						<p class="homecontent">
						<%=dictLanguage("Personal_Monthly_Prod_Goal")%>:
						<b><%=strPersonalProductionGoal%></b> <%=dictLanguage("Hours")%>.<br>
						<%=dictLanguage("You_Are_Currently_At")%> <b><%=strPersonalWorkThisMonth%></b>.

<%		if strPersonalWorkThisMonth > strPersonalProductionGoal then%>
						<%=dictLanguage("You_Need")%> <b><%=FormatNumber(strPersonalProductionGoal-strPersonalWorkThisMonth,2)%></b> <%=dictLanguage("more")%>!<br>
<%		else%>
						<%=dictLanguage("You_Are")%> <b><%=FormatNumber(strPersonalWorkThisMonth-strPersonalProductionGoal,2)%></b> <%=dictLanguage("Over_Your_Goal")%>!<br>
<%		end if%>
						<br>
						<%=dictLanguage("You_Are_Currently_At")%> <b><%=strPersonalWorkToday%></b> <%=dictLanguage("Hours_Today")%>, 
						<b><%=strPersonalWorkYesterday%></b> <%=dictLanguage("Hours_Yesterday")%>.</p>
					</td>
<%		if gsProductionReportGrid then%>
					<td valign="middle"><a href="<%=gsSiteRoot%>reports/production_report_grid.asp" class="home"><%=dictLanguage("Production_Report_Grid")%></a></td>
<%		end if%>
				</tr>
			</table>
<%	end if %>
			</td></tr></table>
			
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<%if gsTimecards then%><td bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("TIMECARDS")%></b></td><%end if%>
					<%if gsTasks then%><td bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("TASKS")%></b></td><%end if%>
					<%if gsProjects then%><td bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("PROJECTS")%></b></td><%end if%>
					<%if gsClients then%><td bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("CLIENTS")%></b></td><%end if%>
					<%if gsPTO then%><td bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("PTO")%></b></td><%end if%>
					<%if gsOther then%><td bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("TOOLS")%></b></td><%end if%>
				</tr>
				<tr>
					<%if gsTimecards then%>
					<td valign="top">
					<%	if session("permTimecardsAdd") then %>
						<a href="<%=gsSiteRoot%>timecards/timecard.asp" class="home"><%=dictLanguage("New")%></a><br>
					<%	end if %>
						<a href="<%=gsSiteRoot%>timecards/timecard-view.asp" class="home"><%=dictLanguage("View_My_Timecards")%></a><br>
						<a href="<%=gsSiteRoot%>reports/client_work_log.asp" class="home"><%=dictLanguage("Work_Log_Internal")%></a><br>
						<a href="<%=gsSiteRoot%>reports/client_work_log_for_client.asp" class="home"><%=dictLanguage("Work_Log_For_Client")%></a><br>
						<a href="<%=gsSiteRoot%>timeSheets/view.asp" class="home"><%=dictLanguage("Hourly_Timesheets")%></a></font><br><br>
					</td>
					<%end if%>
					<%if gsTasks then%>
					<td valign="top">
					<%	if session("permTasksAdd") then %>
						<a href="<%=gsSiteRoot%>tasks/new.asp" class="home"><%=dictLanguage("New")%></a><br>
					<%  end if %>
						<a href="<%=gsSiteRoot%>tasks/view.asp" class="home"><%=dictLanguage("View_Tasks")%></a><br>
						<a href="<%=gsSiteRoot%>tasks/view_mine.asp" class="home"><%=dictLanguage("View_My_Tasks")%></a><br>
						<a href="<%=gsSiteRoot%>tasks/view_my_assigned.asp" class="home"><%=dictLanguage("View_Tasks_I_Assigned")%></a><br>
					</td>
					<%end if%>
					<%if gsProjects then%>
					<td valign="top">
					<%	if session("permProjectsAdd") then %>
						<a href="<%=gsSiteRoot%>projects/project-add.asp" class="home"><%=dictLanguage("New")%></a><br>
					<%	end if %>	
						<a href="<%=gsSiteRoot%>projects/project-view.asp" class="home"><%=dictLanguage("View_Projects")%></a><br>
						<a href="<%=gsSiteRoot%>projects/project-status.asp" class="home"><%=dictLanguage("Project_Status")%></a><br>
					<%	if gsProjectDashboard then %>
						<a href="<%=gsSiteRoot%>projects/dashboard.asp" class="home"><%=dictLanguage("Project_Dashboard")%></a>
					<%  end if %>
					</td>
					<%end if%>
					<%if gsClients then%>
					<td valign="top">
					<%	if session("permClientsAdd") then %>
						<a href="<%=gsSiteRoot%>clients/client-add.asp" class="home"><%=dictLanguage("New")%></a><br>
					<%	end if %>	
						<a href="<%=gsSiteRoot%>clients/client-view.asp" class="home"><%=dictLanguage("View_Clients")%></a><br>
						<a href="<%=gsSiteRoot%>clients/view_mine.asp" class="home"><%=dictLanguage("View_My_Clients")%></a><br><br>
						<%if gsClientContactLog then%>
						<b class="boldalthighlight">MESSAGE CENTER:</b><br>
						<a href="<%=gsSiteRoot%>clients/view_log.asp" class="home"><%=dictLanguage("Client_Contact_Log")%></a><br>
						<a HREF="javascript:popup('<%=gsSiteRoot%>clients/quickAdd.asp')" class="home"><%=dictLanguage("Add_Log_Entry")%></a>
						<%end if %>
					</td>
					<%end if%>
					<%if gsPTO then%>
					<td valign="top">
						<a href="<%=gsSiteRoot%>pto/ptoform.asp" class="home"><%=dictLanguage("Initiate_PTO_Request")%></a><br>
						<a href="<%=gsSiteRoot%>pto/pto-prelimform.asp" class="home"><%=dictLanguage("Preliminary_PTO_Form")%></a><br>
					<%	if session("permPTOAdmin") then%>
						<br>
						<a href="<%=gsSiteRoot%>pto/view.asp" class="home"><%=dictLanguage("View_PTO_Requests")%></a><br>
					<%	end if%>
					<%end if%>						
					<%if gsOther then%>
					<td valign="top">
						<%if gsCalendar then%><a href="javascript: popupcalendar('<%=gsSiteRoot%>calendar/calendar.asp');" class="home"><%=dictLanguage("Calendar")%></a><br><%end if%>
						<%if gsDiscussion then%><a href="<%=gsSiteRoot%>forum/" class="home"><%=dictLanguage("Discussion_Forum")%></a><br><%end if%>
						<%if gsEmployeeDirectory then%><a href="<%=gsSiteRoot%>employees/" class="home"><%=dictLanguage("Employee_Directory")%></a><br><%end if%>
						<%if gsFileRepository then%><a href="<%=gsSiteRoot%>repository/" class="home"><%=dictLanguage("Document_Repository")%></a><br><%end if%>
						<%if gsNews then%><a href="javascript: popup('<%=gsSiteRoot%>news/');" class="home"><%=dictLanguage("News")%></a><br><%end if%>					
						<%if gsResources then%><a href="javascript: popup('<%=gsSiteRoot%>resources/');" class="home"><%=dictLanguage("Resources")%></a><br><%end if%>
					</td>				
					<%end if%>
				<tr>
			</table>

			<!-- put in perfomance tracker -->
			<!-- put in opportunity tracker -->
			<!-- put in pto system -->
	
		</td>
		<td valign="top" width="270">			
<%	if gsHomeCalendar then %>
			<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=gsColorBackground%>" width="100%"><tr><td valign="top">
			<table border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="<%=gsColorBackground%>" width="100%">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>" width="100%">
<%		if gsHomeMinimize then
			if Request.Cookies("calendar")<>"FALSE" then 
				strDisplay = "block" %>
						<img align="right" name="calendar_option" id="calendar_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('calendar');" onmouseover="this.style.cursor='hand'">
<%			else
				strDisplay = "none" %>
						<img align="right" name="calendar_option" id="calendar_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('calendar');" onmouseover="this.style.cursor='hand'">
<%			end if 
		end if %>
						<b class="homeHeader"><%=dictLanguage("CALENDAR")%></b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="calendar" id="calendar">
					<td><!--#include file="includes/calendar.asp"--></td>
				</tr>
				<tr>
					<td width="100%" align="center">
						<a href="javascript: popupcalendar('<%=gsSiteRoot%>calendar/calendar.asp');" class="home"><%=dictLanguage("Click_For_Full_View")%></a>
					</td>
				</tr>				
			</table>
			</td></tr></table>	
<%	end if %>


<%	if gsHomeSurvey then %>
			<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=gsColorBackground%>" width="100%" style="margin-top: 1px;"><tr><td valign="top">
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>">
<%		if gsHomeMinimize then 
			if Request.Cookies("survey")<>"FALSE" then 
				strDisplay = "block" %>
						<img align="right" name="survey_option" id="survey_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('survey');" onmouseover="this.style.cursor='hand'">
<%			else
				strDisplay = "none" %>
						<img align="right" name="poll_option" id="poll_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('survey');" onmouseover="this.style.cursor='hand'">
<%			end if 
		end if %>
						<b class="homeheader"><%=dictLanguage("SURVEY")%></b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="survey" id="survey">
					<td>
<%		thispage = Request.ServerVariables("URL")
		LimitNumofVotes = TRUE
		RecordInCookie = TRUE 

		sql = sql_GetCurrentSurvey
		Call runSQL(sql, rsPolls)
		if not rsPolls.eof then
			PollQuestion = rsPolls("fdPollID")%>
	
				<!--#include file="includes/survey.asp"-->

<% 		else response.write ("<center><br>" & dictLanguage("No_Active_Survey") & ".<br></center>")
		end if
		rsPolls.close
		set rsPolls = nothing %>	
					</td>
				</tr>
			</table>
			</td></tr></table>
<%	end if %>

<%  if gsHomeResources then %>
			<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=gsColorBackground%>" width="100%" style="margin-top: 1px;"><tr><td valign="top">
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>">
<%		if gsHomeMinimize then 
			if Request.Cookies("resources")<>"FALSE" then 
				strDisplay = "block" %>
						<img align="right" name="resources_option" id="resources_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('resources');" onmouseover="this.style.cursor='hand'">
<%			else
				strDisplay = "none" %>
						<img align="right" name="resources_option" id="resources_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('resources');" onmouseover="this.style.cursor='hand'">
<%			end if 
		end if %>
						<b class="homeheader"><%=dictLanguage("RESOURCES")%></b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="resources" id="resources">
					<td>
						<table border="0" cellpadding="2" cellspacing="2" width="100%">
							<tr>
								<td class="homecontent"><!--#include file="includes/resources.asp"--></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			</td></tr></table>
<%  end if %>

<%  if gsHomeExtNews then %>
			<table cellpadding="0" cellspacing="0" border="0" bgcolor="<%=gsColorBackground%>" width="100%" style="margin-top: 1px;"><tr><td valign="top">
			<table border="0" width="100%" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>">
<%		if gsHomeMinimize then 
			if Request.Cookies("extnews")<>"FALSE" then 
				strDisplay = "block" %>
						<img align="right" name="extnews_option" id="extnews_option" src="<%=gsSiteRoot%>images/x.gif" border="0" alt="<%=dictLanguage("Collapse_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('extnews');" onmouseover="this.style.cursor='hand'">
<%			else
				strDisplay = "none" %>
						<img align="right" name="extnews_option" id="extnews_option" src="<%=gsSiteRoot%>images/maximize.gif" border="0" alt="<%=dictLanguage("Open_This_Window")%>" width="18" height="15" onClick="javascript: toggleWindow('extnews');" onmouseover="this.style.cursor='hand'">
<%			end if 
		end if %>
						<b class="homeheader"><%=dictLanguage("NEWS_FEED")%></b>
					</td>
				</tr>
				<tr style="display: <%=strDisplay%>;" name="extnews" id="extnews">
					<td>
						<table border="0" cellpadding="2" cellspacing="2" width="100%">
							<tr>
								<td class="homecontent"><!--#include file="includes/extnews.asp"--></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			</td></tr></table>
<%  end if %>
		</td>
	</tr>
</table>

<!--#include file="includes/main_page_close.asp"-->

