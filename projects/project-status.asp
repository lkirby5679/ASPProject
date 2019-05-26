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
Function GetEmpName(ID)
	if ID <> "" then
		sql = sql_GetEmployeesByID(ID)
		'response.write sql
		Call RunSQL(sql, rs)
		if not rs.eof then
			empName = rs("EmployeeName")
		else
			empName = ""
		end if
		rs.close
		set rs = nothing
	else
		empName = ""
	end if
	GetEmpName = empName
End Function

Function GetHoursUsed(Project, EmpType)
	sql = sql_GetProjectHoursByEmpType(Project,EmpType)
	Call RunSQL(sql, rs)
	if not rs.eof then
		if rs("HoursUsed")<>"" then
			GetHoursUsed = rs("HoursUsed")
		else
			GetHoursUsed = 0
		end if
	else
		GetHoursUsed = 0
	end if
	rs.close
	set rs = nothing
End Function

Function GetPhaseHoursUsed(Project, Phase)
	sql = sql_GetProjectHoursByPhase(Project, Phase)
	'response.write sql	
	Call RunsQL(sql, rs)
	if not rs.eof then
		if rs("HoursUsed")<>"" then
			GetPhaseHoursUsed = rs("HoursUsed")
		else
			GetPhaseHoursUsed = 0
		end if
	else
		GetPhaseHoursUsed = 0
	end if
	rs.close
	set rs = nothing
End Function

active = request.querystring("active")
if active = "" or active = "Active" then 
	active="Active"
	activeView = dictLanguage("Active")
elseif active="All" then
	active="All"
	activeView = dictLanguage("All")
end if
sortby = request.querystring("sortby")
if sortby = "" then 
	sortby="ByClient" 
end if
typeProject = request.querystring("typeProject")
if typeProject = "" then
	typeProject = "Fixed-Price"
end if
if typeProject = "all" then
	boolAllTypes = TRUE
end if

sql = sql_GetProjectTypes()
Call RunSQL(sql, rsProjectTypes)

sql = sql_GetEmployeeTypes()
Call RunsQL(sql, rsRates)

dim phase(150)
phase(0) = "No Phase Indicated"
sql = sql_GetProjectPhases()
Call RunSQL(sql, rsPhases)
do while not rsPhases.eof
	for x=0 to UBound(phase)
		if x=cint(rsPhases("projectphaseid")) then
			phase(x)=rsPhases("projectphasename")
			exit for
		end if
	next
	rsPhases.movenext
loop
rsPhases.close
set rsPhases = nothing

sql = sql_GetProjectView(typeProject, boolAllTypes, active)
'response.write sqlViewProjects
Call RunSQL(sql, rsViewProjects)
%>

<!--#include file="../includes/main_page_open.asp"-->

<font face="Verdana,Arial,Helvetica" size="2">
	<b class="homeheader"><%=dictLanguage("View")%>&nbsp;<%=activeView%>&nbsp;<%=dictLanguage("Projects")%></b> 
<%if active="Active" then%>
	(<%=dictLanguage("All_Active_Projects")%>)
<%end if%>
	<br>
	<table border="0" cellpadding="2" cellspacing="2" width="100%">
		<tr>
			<td>
<%if active="Active" then%>
				<a href="project-status.asp?active=All&typeProject=<%=typeProject%>" class="small"><%=dictLanguage("View_All_Projects")%></a>
<%else%>
				<a href="project-status.asp?typeProject=<%=typeProject%>" class="small"><%=dictLanguage("View_Active_Projects")%></a>
<%end if%>
				<br><br>
			</td>
			<td align="right">
				<table border="0">
					<tr>
<%while not rsProjectTypes.eof %>
						<td><a href="project-status.asp?active=<%=active%>&typeProject=<%=server.urlencode(rsProjectTypes("ProjectTypeDescription"))%>" class="small"><%=rsProjectTypes("ProjectTypeDescription")%></a></td>	
<%	rsProjectTypes.movenext 
  wend
  rsProjectTypes.close
  set rsProjectTypes = nothing %>
					</tr>
				</table>
			</td>
		</tr>
	</table>

<%totalProjectedHours = 0
  grandTotalProjectedHours = 0
  totalCurrentHours = 0
  grandTotalCurrentHours = 0
  projectType = ""  
  intRowCounter = 0
  firstTime = "True"

  if not rsViewProjects.eof then
	  while not rsViewProjects.EOF
		color = rsViewProjects("color")
		intRowCounter = intRowCounter + 1
	
		if projectType <> rsViewProjects("ProjectTypeDescription") then
	  		projectType = rsViewProjects("ProjectTypeDescription")
	  		if firstTime = "True" then
				firstTime = "False"
			else
				Response.Write "<BR>"
			end if 
		end if 
		if intRowCounter MOD 2 = 1 then 
			strRowColor = gsColorWhite
		else 
			strRowColor = "#FFFFFF"
		end if %>

	<table width="100%" border=0 cellpadding=2 cellspacing=2 style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
		<tr bgcolor="<%=strRowColor%>">
			<td valign=top>
				<table border="0" cellpadding="2" cellspacing="2" width="100%">
					<tr>
						<td width="40%" valign="top">
							<table border="0" cellpadding="2" cellspacing="2" width="100%">
								<tr><td colspan="2" class="small"><%=dictLanguage("Client")%>: <b><a href="<%=gsSiteRoot%>clients/client-edit.asp?client_id=<%=rsViewProjects("cl_id")%>"><%=rsViewProjects("Client_Name")%></a></b></td></tr>
								<tr><td colspan="2" class="small"><%=dictLanguage("Project")%>: <b><a href="project-edit.asp?project_id=<%=rsViewProjects("Project_ID")%>"><%=rsViewProjects("Description")%></a></b></td></tr>
<%		if rsViewProjects("forum_id")<>"" and gsDiscussion then %>
								<tr><td colspan="2" class="small"><b><a href="<%=gsSiteRoot%>forum/default.asp?fid=<%=rsViewProjects("forum_id")%>"><%=dictLanguage("Project_Discussion")%></a></b></td></tr> 
<%		end if %>	
<%		if rsViewProjects("folder_id")<>"" and gsFileRepository then %>
								<tr><td colspan="2" class="small"><b><a href="<%=gsSiteRoot%>repository/default.asp?folder=<%=rsViewProjects("folder_id")%>"><%=dictLanguage("Project_Documents")%></a></b></td></tr> 
<%		end if %>	
								<tr><td colspan="2" class="small"><%=dictLanguage("Work_Order")%>: <b><%=rsViewProjects("WorkOrder_Number")%></b></td></tr>
								<tr>
									<td class="small">Start Date: <b><%=rsViewProjects("Start_Date")%></b></td>
									<td align="right" class="small"><%=dictLanguage("End_Date")%>: <b><%=rsViewProjects("Launch_Date")%></b></td>
								</tr>		
								<tr><td colspan="2" class="small"><%=dictLanguage("Account_Rep")%>: <b><%=GetEmpName(rsViewProjects("AccountExec_ID"))%></b></td></tr>
								<tr><td colspan="2">&nbsp;</td></tr>
								<tr><td colspan="2" align="center"><a href="javascript://" onClick="window.open('project-status_single.asp?project_id=<%=rsViewProjects("project_id")%>','','height=400,width=700,toolbar=no,menubar=yes,scrollbars=no,resizable=yes');" class="small"><b><%=dictLanguage("Print_This_Summary")%></b></a></td></tr>
							</table>
						</td>
						<td width="*" valign="top">
							<table border="0" cellpadding="2" cellspacing="2" width="100%" style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
								<tr>
									<td align="center" bgcolor="<%=color%>" class="small"><b><%=dictLanguage("Status")%>: <%=color%></b></td>
									<td align="center" class="small"><b><%=dictLanguage("B")%></b></td>
									<td align="center" class="small"><b><%=dictLanguage("A")%></b></td>
									<td align="center" class="small"><b><%=dictLanguage("R")%></b></td>		
									<td align="center" class="small"><b><%=dictLanguage("Core_Team")%></b></td>
								</tr>
<%		rsRates.movefirst
		budgetTotal = 0
		usedTotal = 0
		remainingTotal = 0
		sql = sql_GetProjectBudgetsByID(rsViewProjects("Project_id"))
		'Response.Write sql & "<BR>"
		Call runSQL(sql, rsBudget)
		while not rsRates.eof 
			budget = 0
			used = 0
			remaining = 0
			boolThisBudget = FALSE
			if not rsBudget.eof then
				if trim(rsBudget("employeetype_id")) = trim(rsRates("employeetype_id")) then 
					boolThisBudget = TRUE
					budget = rsBudget("Hours")
				end if
			end if			
			if isNull(budget) then
				budget = 0
			end if
	 		budgetTotal = budgetTotal + budget
			used = GetHoursUsed(rsViewProjects("project_id"),rsRates("EmployeeType_ID"))
			if isNull(used) then
				used = 0
			end if
			usedTotal = usedTotal + used 
			remaining = budget - used
			remainingTotal = remainingTotal + remaining
			if budget<>0 or used<>0 then %>
	
								<tr>
									<td class="small"><%=rsRates("EmployeeType")%></td>
									<td align="right" class="small"><%=budget%></td>
									<td align="right" class="small"><%=used%></td>
									<td align="right" class="small"><%=formatnumber(remaining,2,-1,-1)%></td>
									<td class="small"><%if boolThisBudget then %><%=GetEmpName(rsBudget("Employee_ID"))%><%else response.write "&nbsp;"%><%end if %></td>
								</tr>
<%	 		end if
			if boolThisBudget then
				rsBudget.movenext
			end if
			rsRates.movenext 
		wend 
		rsBudget.close
		set rsBudget = nothing %>
								<tr>
									<td class="small"><b><%=dictLanguage("Total")%></b></td>
									<td align="right" class="small"><b><%=budgetTotal%></b></td>
									<td align="right" class="small"><b><%=usedTotal%></b></td>
									<td align="right" class="small"><b><%=formatnumber(remainingTotal,2,-1,-1)%></b></td>
									<td>&nbsp;</td>
								</tr>
							</table>
							<br>
							<table border="0" cellpadding="2" cellspacing="2" width="100%" style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
								<tr>
									<td align="center" class="small"><b><%=dictLanguage("Phase")%></b></td>
									<td align="center" class="small"><b><%=dictLanguage("Hours_Used")%></b></td>		
								</tr>
<%		for x = 0 to 25 
			if phase(x)<>"" then
				phaseHoursUsed = GetPhaseHoursUsed(rsViewProjects("project_id"),x) 
				if phaseHoursUsed <> 0 then%>		
								<tr>
									<td class="small"><%=phase(x)%></td>
									<td align="right" class="small"><%=phaseHoursUsed%></td>		
								</tr>
<%				end if
			end if
		next %>
							</table>		
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<% 
		rsViewProjects.MoveNext
	  wend
   else %>
   
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr>
			<td align="Center">
				<br><%=dictLanguage("No_Projects_For_Type")%> &quot;<%=typeProject%>&quot;.  Select another type.<br><br><br>
			</td>
		</tr>
	</table>   

<% end if 
rsRates.close
set rsRates = nothing
rsViewProjects.close
set rsViewProjects = nothing
%>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<!--#include file="../includes/main_page_close.asp"-->

