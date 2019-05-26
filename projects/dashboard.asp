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

<% 
pid = request("pid")
if pid <> "" then
	sql = sql_GetProjectViewByID(pid)
	Call RunSQL(sql, rs)
	if not rs.eof then
		pName     = rs("Description")
		cID       = rs("cl_id")
		cName     = rs("client_name")
		pforumID  = rs("forum_id")
		pmainforumID = pforumID
		pfolderID = rs("folder_ID")
	end if
	rs.close
	set rs = nothing
end if
folder = request("folder")
if folder = "" then
	folder = pfolderID
end if 

if pid = "" then %>


<table cellpadding="2" cellspacing="2" border=0 align="center">
	<tr>
		<td>
			<form method="POST" action="dashboard.asp" id=form1 name=form1>
			<input type="hidden" name="active" value="<%=active%>"> 
				<table border="0" cellpadding="2" cellspacing="2" align="Center">
					<tr bgcolor="<%=gsColorHighlight%>"><td align="center" colspan="2" class="homeheader"><%=dictLanguage("Select_Project_For_Dashboard")%></td></tr>
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr>
						<td><%=dictLanguage("Select_Project")%>:</td>
						<td>
							<select name="pid" class="formstyleXL">
<%
sql = sql_GetAllProjectsWithClients()
Call RunSQL(sql, rs)
while not rs.eof
	response.write "<option value='" & rs("Project_ID") & "'>" & rs("Client_Name") & ": " & rs("Description") & " "
	rs.movenext
wend
rs.close
set rs = nothing
%>
							</select>
							<input type="submit" value="Submit" class="formButton" id=submit1 name=submit1>
						</td>
					</tr>
				</table>
			</form>
		</td>
	</tr>
</table>


<%
else
%>

<table border="0" cellpadding="2" cellspacing="2" width="100%">
<tr><td colspan="2">
		<table cellpadding="2" cellspacing="2" border="0">
			<tr>
				<td><b><%=dictLanguage("Project")%>: </b></td>
				<td><b><%=cName%> - <%=pName%></b>&nbsp;&nbsp;<a href="clients/client-edit.asp?client_id=<%=cid%>"><small>(<%=dictLanguage("Edit_Client")%>)</small></a>&nbsp;&nbsp;<a href="project-edit.asp?project_id=<%=pid%>"><small>(<%=dictLanguage("Edit_Project")%>)</small></a></td>
			</tr>
		</table>
	</td>
</tr>	
<tr>
	<td valign="top">
		<table border="0" cellpadding="2" cellspacing="0">
			<tr bgcolor="<%=gsColorHighlight%>"><td class="columnheader"><%=dictLanguage("Project_Status")%></td></tr>		
			<tr><td>
				<%			
				project_id = pid
				if project_id <> "" then
					sql = sql_GetProjectTypes()
					Call RunSQL(sql, rsProjectTypes)

					sql = sql_GetEmployeeTypes()
					Call RunSQL(sql, rsRates)

					dim phase(150)
					phase(0) = "No Phase Indicated"
					sql = sql_GetProjectPhases()
					Call RunSQL(sql, rsPhases)
					while not rsPhases.eof
						for x=0 to UBound(phase)
							if x=cint(rsPhases("projectphaseid")) then
								phase(x)=rsPhases("projectphasename")
								exit for
							end if
						next
						rsPhases.movenext
					wend
					rsPhases.close
					set rsPhases = nothing

					sql = sql_GetProjectViewByID(project_id)
					'response.write sql
					call RunSQL(sql, rsViewProjects)

					totalProjectedHours = 0
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
								strRowColor = "#FFFFFF"
							else
								strRowColor = gsColorWhite 							
							end if %>

				<table width="100%" border="0" cellpadding="2" cellspacing="0" style="border-width: 2px; border-style: solid; border-color: <%=gsColorHighlight%>">
					<tr bgcolor="<%=strRowColor%>">
						<td valign=top>
							<table border="0" cellpadding="2" cellspacing="2" width="100%">
								<tr>
									<td width="40%">
										<table border="0" cellpadding="2" cellspacing="2" width="100%">
											<tr><td colspan="2" class="small"><%=dictLanguage("Client")%>: <b><a href="<%=gsSiteRoot%>clients/client-edit.asp?client_id=<%=rsViewProjects("cl_id")%>"><%=rsViewProjects("Client_Name")%></a></b></td></tr>
											<tr><td colspan="2" class="small"><%=dictLanguage("Project")%>: <b><a href="project-edit.asp?project_id=<%=rsViewProjects("Project_ID")%>"><%=rsViewProjects("Description")%></a></b></td></tr>
				<%		if rsViewProjects("forum_id")<>"" and gsDiscussion then %>
												<tr><td colspan="2" class="small"><b><a href="<%=gsSiteRoot%>forum/default.asp?fid=<%=rsViewProjects("forum_id")%>">Project Discussion Forum</a></b></td></tr> 
				<%		end if %>	
				<%		if rsViewProjects("folder_id")<>"" and gsFileRepository then %>
												<tr><td colspan="2" class="small"><b><a href="<%=gsSiteRoot%>repository/default.asp?folder=<%=rsViewProjects("folder_id")%>">Project File Repository</a></b></td></tr> 
				<%		end if %>	
											<tr><td colspan="2" class="small"><%=dictLanguage("Work_Order")%>: <b><%=rsViewProjects("WorkOrder_Number")%></b></td></tr>
											<tr>
												<td class="small"><%=dictLanguage("Start_Date")%>: <b><%=rsViewProjects("Start_Date")%></b></td>
												<td align="right" class="small"><%=dictLanguage("End_Date")%>: <b><%=rsViewProjects("Launch_Date")%></b></td>
											</tr>
											<tr><td colspan="2" class="small"><%=dictLanguage("Account_Rep")%>: <b><%=GetEmpName(rsViewProjects("AccountExec_ID"))%></b></td></tr>
											<tr><td colspan="2" class="small">&nbsp;</td></tr>
											<tr><td colspan="2" align="center" class="small">&nbsp;</td></tr>
										</table>
									</td>
									<td width="*">
										<table border="0" cellpadding="2" cellspacing="2" width="100%" style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
											<tr>
												<td align="center" bgcolor="<%=color%>" class="small"><b><%=dictLanguage("Status")%>: <%=color%></b></td>
												<td align="center" class="small"><b><%=dictLanguage("B")%></b></td>
												<td align="center" class="small"><b><%=dictLanguage("A")%></b></td>
												<td align="center" class="small"><b><%=dictLanguage("R")%></b></td>		
												<td align="center" class="small"><b><%=dictLanguage("Core_Team")%></b></td>
											</tr>
				<%				rsRates.movefirst
								budgetTotal = 0
								usedTotal = 0
								remainingTotal = 0
								sqlBudget = sql_GetProjectBudgetsByID(rsViewProjects("Project_id"))
								Call runSQL(sqlBudget, rsBudget)
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
												<td align="right" class="small"><%=formatnumber(remaining,1,-1,-1)%></td>
												<td class="small"><%if boolThisBudget then%><%=GetEmpName(rsBudget("Employee_ID"))%><%else Response.Write "&nbsp;"%><%end if%></td>
											</tr>
				<% 					end if
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
												<td align="right" class="small"><b><%=formatnumber(remainingTotal,1,-1,-1)%></b></td>
												<td class="small">&nbsp;</td>
											</tr>
										</table>
										<br>
										<table border="0" cellpadding="2" cellspacing="2" width="100%" style="border-width: 2; border-style: solid; border-color: <%=gsColorBackground%>">
											<tr>
												<td align="center" class="small"><b><%=dictLanguage("Phase")%></b></td>
												<td align="center" class="small"><b><%=dictLanguage("Hours_Used")%></b></td>		
											</tr>
				<%				for x = 0 to 25 
									if phase(x)<>"" then
										phaseHoursUsed = GetPhaseHoursUsed(rsViewProjects("project_id"),x) 
										if phaseHoursUsed <> 0 then%>		
											<tr>
												<td class="small"><%=phase(x)%></td>
												<td align="right" class="small"><%=phaseHoursUsed%></td>		
											</tr>
				<%						end if
						 			end if
								next %>
										</table>			
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
				<%			rsViewProjects.MoveNext
						wend
					else 
						Response.Write dictLanguage("Invalid_Project")
					end if
				else	
					Response.Write dictLanguage("Invalid_Project")
				end if
				rsRates.close
				set rsRates = nothing
				rsViewProjects.close
				set rsViewProjects = nothing %>					
			</td></tr>
		</table>
	</td>




	<td valign="top">
		<table cellpadding="2" cellspacing="0" border="0">

<% if gsProjectQuotes then %>
			<tr bgcolor="<%=gsColorHighlight%>">
				<td class="columnheader" colspan="2"><%=dictLanguage("Project_Quotes")%></td>
			</tr>
			<tr>
				<td colspan="2">
					<table width="100%" border="0" cellpadding="2" cellspacing="0" style="border=2px; border-color=<%=gsColorHighlight%>; border-style: solid;">
						<tr>
							<td width="50%" valign="top"><a href="javascript: popup('project-quote.asp?pid=<%=pid%>');"><%=dictLanguage("Create_Project_Quote")%></a></td>
<%			sql = sql_GetProjectQuotesByProjectID(project_id)
			Call runSQL(sql, rsQuotes)
			if not rsQuotes.eof then %>
							<td width="50%" valign="top"><%=dictLanguage("Existing_Project_Quotes")%><br><br>
<%			while not rsQuotes.eof %>
								<a href="javascript: popup('project-quote.asp?Submit=view&pid=<%=project_id%>&pqid=<%=rsQuotes("projectquote_id")%>');"><%=dictLanguage("Date")%>: <%=rsQuotes("dateentered")%>, <%=dictLanguage("Total")%>: <%=formatnumber(rsQuotes("total_price"),2,-1,0,0)%></a></br>	
<%				rsQuotes.movenext 
			wend
			Response.Write "</td>"
			else
				Response.Write "<td>&nbsp;</td>"
			end if
			rsQuotes.close
			set rsQuotes = nothing%>					
						</tr>					
					
					</table>
				</td>
			</tr>	
<%	end if %>

<%  if gsFileRepository then %>
			<tr><td colspan="2">&nbsp;</td></tr>
			<tr bgcolor="<%=gsColorHighlight%>">
				<td class="columnheader"><%=dictLanguage("Project_Documents")%></td>
				<td align="right"><a href="repository/default.asp?folder=<%=pfolderid%>"><small><%=dictLanguage("Click_Doc_Rep")%></small></a></td>
			</tr>
			<tr>
				<td colspan="2">
		<% if pfolderID <> "" then %>				
									<table width="100%" border="0" cellpadding="2" cellspacing="0" style="border=2px; border-color=<%=gsColorHighlight%>; border-style: solid;">
			<%
			sql = sql_GetSubFoldersByFolderID(folder)
			Call RunSQL(sql, rsSubs)

			sql = sql_GetDocumentsByFolderID(folder)
			Call RunSQL(sql, rsSearch)
				
			if (not rsSubs.eof) or (not rsSearch.eof) then%>
										<tr bgcolor="<%=gsColorHighlight%>" >
											<td nowrap><b class="bolddark"><%=dictLanguage("Document")%></b></td>
											<td nowrap><b class="bolddark"><%=dictLanguage("Date_Submitted")%></b></td>
											<td nowrap><b class="bolddark"><%=dictLanguage("Contact")%></b></td>
											<td><img src="<%=gsSiteRoot%>images/document.gif" WIDTH="11" HEIGHT="14"></td>
										</tr>
			<%	do while not rsSubs.eof
					intRowcounter = intRowcounter + 1
					If intRowcounter MOD 2 = 1 then 
						colour = "bgcolor=#F0F8FF"
					Else
						colour = "bgcolor=#ffffff"
					End If 
					Response.Write "<tr " & bgcolor & ">"
					strFolderID		= rsSubs("folder_id")
					strFolderName	= trim(rsSubs("folder"))
					strFolderImage	= "<img src='" & gsSiteRoot & "images/close_folder.gif' WIDTH='19' HEIGHT='19'>" %>
											<td nowrap class="small">
												<%=strFolderImage%>
												<a href="dashboard.asp?folder=<%=strFolderID%>&pid=<%=pid%>" class="small" title="<%=strFolderName%>"><%=strFolderName%></a>
											</td>
											<td nowrap class="small">&nbsp;</td>
											<td nowrap class="small">&nbsp;</td>
											<td nowrap class="small">&nbsp;</td>
										</tr>		
			<%		rsSubs.movenext
				loop
				do while not rsSearch.eof
					sql = sql_GetEmployeesByID(rsSearch("submitBy"))
					Call RunSQL(sql, rsContact)

					intRowcounter = intRowcounter + 1

					If intRowcounter MOD 2 = 1 then 
						colour = "bgcolor=#F0F8FF"
					Else
						colour = "bgcolor=#ffffff"
					End If 
					Response.Write "<tr " & bgcolor & ">"
					strFileName = rsSearch("fileName")
					if strFileName<>"" then
						boolURL = FALSE
						if right(lcase(strFileName),3) = "doc" or right(lcase(strFileName),3) = "txt" or right(lcase(strFileName),3) = "htm" or right(lcase(strFileName),3) = "rtf" or right(lcase(strFileName),3) = "xls" or right(lcase(strFileName),3) = "pdf" or right(lcase(strFileName),3) = "ppt" then
							file_image = right(lcase(strFileName),3) & ".gif"
						elseif right(lcase(strFileName),4) = "html" then
							file_image = "htm.gif"
						else
							file_image = "document.gif"
						end if 
					else 
						boolURL=TRUE
						strFileName = rsSearch("url")
						file_image = "htm.gif"
					end if %>
											<td nowrap class="small">
												<img src="<%=gsSiteRoot%>images/<%=file_image%>" align="baseline" WIDTH="11" HEIGHT="14">
												<%if boolURL then%>
												<a href="<%=SpaceEncode(strFileName)%>" class="small" title="<%=rsSearch("description")%>" target="_blank"><%=SpaceDecode(strFileName)%></a> (URL)
												<%else%>
												<a href="<%=gsSiteRoot%>repository/library/<%=SpaceEncode(strFileName)%>" title="<%=rsSearch("description")%>" class="small" target="_blank"><%=SpaceDecode(strFileName)%></a>
												<%end if%>
											</td>
											<td nowrap align="center" class="small"><%=rsSearch("submitDate")%></td>
											<td nowrap class="small"><a href="<%=gsSiteRoot%>employees/default.asp?employeeid=<%=rsSearch("submitBy")%>" class="small"><%=rsContact("employeeName")%></a></td>
											<td align="left" class="small">
			<%		if trim(rsSearch("submitBy")) = trim(session("employee_id")) then%>
												<a href="repository/delete.asp?file=<%=rsSearch("ID")%>" class="small" onClick="return confirm('<%=dictLanguage("Confirm_Delete_Document")%>');"><img src="<%=gsSiteRoot%>images/delete.gif" WIDTH="20" HEIGHT="19" border="0"></a>
			<%		end if%>
											</td>
										</tr>
			<%		rsSearch.MoveNext
				loop 
			else %>						<tr>
											<td align="center"><b><%=dictLanguage("Folder_Empty")%></b></td>
										</tr>
			<%end if
			rsSubs.close
			set rsSubs = nothing
			rsSearch.close
			set rsSearch = nothing%>	
									</table>			
			<%else%>
				<%=dictLanguage("No_Project_Documents")%>
			<%end if%>
			
				</td>
			</tr>	
<%	end if %>

<%  if gsDiscussion then %>			
			<tr><td colspan="2">&nbsp;</td></tr>
			<tr bgcolor="<%=gsColorHighlight%>">
				<td class="columnheader"><%=dictLanguage("Project_Discussion")%></td>
				<td align="right"><a href="forum/default.asp?fid=<%=pforumID%>"><small><%=dictLanguage("Click_Discussion_Forum")%></small></a></td>
			</tr>
			<tr>
				<td colspan="2">
				<% if pforumid <> "" then %>
					<!--#include file="../includes/forum_common.asp"-->

					<!-- forum code based on Ti Portal's forum sample.  http://www.transworldportal.com -->

					<table border="0" cellpadding="2" cellspacing="0" style="border=2px; border-color=<%=gsColorHighlight%>; border-style: solid;">
						<tr>
							<td>  								
							<%
							boolAdmin = false

							iActiveForumId = pforumid	'Set Project Folder here.
							If IsNumeric(iActiveForumId) Then
								iActiveForumId = CInt(iActiveForumId)
							Else
								iActiveForumId = 0
							End If

							iPeriodToShow = Request.QueryString("pts")
							If IsNumeric(iPeriodToShow) Then
								iPeriodToShow = CInt(iPeriodToShow)
							Else
								iPeriodToShow = 0
							End If

							' Get Forum Info and count of messages in the forum
									
							if session("su_section")<>"" then
							    objForumRS= GetForumsBySect(session("su_section"))
							else
							    objForumRS= GetAllForums()
							end if
							objForumCountRS = GetAllForumsMessageCounts()

							If Isarray(objForumRS) Then	
								For intcounter = 0 To UBound(objforumrs)
									strForumBreakdownType = Trim(LCase(objForumRS(intcounter)(forums_forum_grouping)))
											
									'Get Message Count for Forum
									iForumMessageCount=0
									if isarray(objForumCountRS) then
										for intInnerCounter = 0 to ubound(objForumCountRS)
											If objForumCountRS(intInnerCounter)(0) = objForumRS(intcounter)(forums_forum_id) then
												iForumMessageCount = objForumCountRS(intInnerCounter)(1)
											end if
										next
									end if

									' If active forum -> show messages Else just show forum
									If objForumRS(intcounter)(forums_forum_id) = iActiveForumId Then
										ShowForumLine objForumRS(intcounter)(forums_forum_id), "open", _
													  objForumRS(intcounter)(forums_forum_name), _
													  objForumRS(intcounter)(forums_forum_description), _
													  iForumMessageCount,0,false

										' Show links to previous months
										iPeriodsToGoBack = DateDiff("m", objForumRS(intcounter)(forums_forum_start_date), Now())
												
										' Make adjustments to periods to go back and show for non-monthly breakdown
										Select Case strForumBreakdownType
											Case "7days"
												iPeriodsToGoBack = iPeriodsToGoBack + 3
											Case "monthly"
												' Nothing to do!
											Case Else
												iPeriodsToGoBack = 0
												iPeriodToShow = 0
										End Select
														
										For iPeriodLooper = 0 To iPeriodsToGoBack
											If strForumBreakdownType = "7days" Or strForumBreakdownType = "monthly" Then
												'Do period message count here.
												ShowPeriodLine objForumRS(intcounter)(forums_forum_id), _
															   strForumBreakdownType, iPeriodLooper, 0, 0, false
											End If

											If iPeriodLooper = iPeriodToShow Then
												'Show Root Level Posts for the selected period and their reply count
												Select Case strForumBreakdownType
													Case "7days"
														If iPeriodToShow <= 2 Then
															dStartDate = Date() - (7 * (iPeriodToShow + 1)) + 1
															dEndDate = Date() - (7 * iPeriodToShow) + 1
														Else
															dStartDate = GetNMonthsAgo(iPeriodToShow - 3)
															dEndDate = GetNMonthsAgo(iPeriodToShow - 4)
														End If
													Case "monthly"
														dStartDate = GetNMonthsAgo(iPeriodToShow)
														dEndDate = GetNMonthsAgo(iPeriodToShow - 1)
													Case Else
														dStartDate = objForumRS(intcounter)(forums_forum_start_date)
														dEndDate = Date() + 1
												End Select
														
												objMessageRS = GetForumMessages(iActiveForumId)
							   					'Build the list of root posts we need counts for
												If isarray(objMessageRS) Then
													for int2InnerCount = 0 to  ubound(objMessageRS)
														strThreadList = strThreadList & objMessageRS(int2InnerCount)(Message_thread_id) & ","
													next
													strThreadList = Left(strThreadList, Len(strThreadList) - 1)
												Else
													strThreadList = (0)
												End If
											
												objMessageCountRS = GetMessageReplies(strThreadList)
														
												If isarray(objMessageRS) Then
													for int2InnerCount = 0 to  ubound(objMessageRS)
														'Get Message Replies
														intMessageReplies = 0	
														if isarray(objMessageCountRS) then
															for int3InnerCount = 0 to ubound(objMessageCountRS)
																if objMessageCountRS(int3InnerCount)(0) = objMessageRS(int2innercount)(messages_message_id)then
																	intMessageReplies = objMessageCountRS(int3InnerCount)(1) - 1
																end if
															next
														end if

														ShowMessageLine 1, objMessageRS(int2innercount)(messages_message_id), _
																		   objMessageRS(int2innercount)(messages_message_subject), _
																		   objMessageRS(int2innercount)(messages_message_author), _
																		   objMessageRS(int2innercount)(messages_message_author_email), _
																		   objMessageRS(int2innercount)(messages_message_timestamp), _
																		   intMessageReplies, "forum", 0,false
													next												   
												End If

											End If
										Next 'iPeriodLooper
										'Set active Forum Name for later use in post line
										iActiveForumName = objForumRS(intcounter)(forums_forum_name)
									End If
								next
							Else
								WriteLine dictLanguage("No_Forums_Open") & "<BR>"
							End If


							If iActiveForumId <> 0 Then
							%>
							<BR>
							<A HREF="forum/post_message.asp?fid=<%= iActiveForumId %>"><I><%=dictLanguage("Post_Message_To")%> <B><%= iActiveForumName %></B></I></A><BR>
							<%
							End If
							%>
						</TD>		
						</TR>	
						<tr><td align="right"><small><%=dictLanguage("Forum_Code_Based")%> <a href="http://www.asp101.com/" target="_blank"><small>ASP 101</small></a>'s free forum.</small></td></tr>		
					</table>
				<%else%>
					<%=dictLanguage("No_Project_Discussion")%>
				<%end if%>		
				</td>
			</tr>
<%	end if %>			
			
		</table>
	</td>
</tr>






<tr><td colspan="2">&nbsp;</td></tr>
<tr bgcolor="<%=gsColorHighlight%>"><td class="columnheader" colspan="2"><%=dictLanguage("Project_Tasks")%></td></tr>	
<tr>
	<td colspan="2">
		<table border="0" cellpadding="2" cellspacing="2" width="100%" style="border=2px; border-color=<%=gsColorHighlight%>; border-style: solid;">
			<tr bgcolor="<%=gsColorHighlight%>">
				<td valign=top class="columnheader">&nbsp;</td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Done")%>?</td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Date_Due")%></td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Priority")%></td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Created_By")%></td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Assigned_To")%></td>
				<td valign=top class="columnheader" width="100%"><%=dictLanguage("Description")%></td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Est_Total_Hours")%></td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Est_Hours_Left")%></td>
				<td valign=top class="columnheader" nowrap><%=dictLanguage("Date_Created")%></td>
				<td valign=top class="columnheader">&nbsp;</td>
			</tr>
			<tr><td valign="top" colspan="11"><b><%=dictLanguage("Active_Tasks")%></b></td></tr>
		<%
			sql = sql_GetTaskViewByProjectID(1,pID) '(active, projectID)
			Call RunSQL(sql, rs)
			while not rs.eof
				boolDone = rs("Done") %>
			<tr<%if boolDone = 0 or not boolDone then%> bgcolor="#FFFF77" <%else%> bgcolor="#EEEEEE" <%end if%>>
				<td valign="top" class="small"><a href="tasks/edit.asp?id=<%=rs("Task_ID")%>"><%=dictLanguage("Edit")%></a></td>
				<td valign="top" align="center"><a href="tasks/toggle_done.asp?id=<%=rs("Task_ID")%>&value=
		<%		if rs("Done") = 0 then
					response.write "no"
				else
					response.write "yes"
				end if %>&referring_page=view2.asp&sql=<%=Server.URLEncode(sql)%>">
		<%		if boolDone = True then
					response.write "<img src='" & gsSiteRoot & "images/done.gif' border=0>"
				else
					response.write "<img src='" & gsSiteRoot & "images/notDone.gif' border=0>"
				end if %>
				</td>
				<td valign=top class="small"><%=rs("DateDue")%></td>
				<td valign="top" align=center class="small"> 
		<%		select case rs("Priority")
					case 1
						response.write dictLanguage("Low")
					case 2
						response.write "<font color=""blue""><b>" & dictLanguage("Medium") & "</b></font>"
					case 3
						response.write "<font color=""red""><b>" & dictLanguage("High") & "</b></font>"
					case 0 
						response.write "<font color=""green""><b>" & dictLanguage("On_Hold") & "</b></font>"
					case else
						response.write "<font color=""blue""><b>" & dictLanguage("Medium") & "</b></font>"
				end select%>
				</td>
				<td valign=top class="small"><%=rs("EmployeeName")%></td>
				<td valign=top class="small"><%=rs("EmployeeName_1")%></td>
				<td valign=top class="small"><%=rs("Desc1")%></td>
				<td valign=top class="small"><%=rs("EstimatedHours")%></td>
		<%
				sqlHours = sql_GetTimecardHoursByTaskID(rs("Task_ID"))
				Call RunSQL(sqlHours, rsHours)
				if rsHours("TimeCharged")<>"" then
					timeCharged = rsHours("TimeCharged")
				else 
					timeCharged = 0
				end if
				estHoursLeft = rs("EstimatedHours") - timeCharged
				rsHours.close
				set rsHours = nothing %>
				<td valign=top class="small"><font <%if estHoursLeft<0 then%>color="#FF0033" <%end if%>class="small"><%=estHoursLeft%></font></td>
				<td valign=top class="small"><%=rs("DateCreated")%></td>
				<td valign="top" class="small">
		<%	if rs("show") = "1" or rs("show") = "-1" or rs("show") then %>			
					<a href="tasks/delete.asp?id=<%=rs("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Delete")%>');"><%=dictLanguage("Delete")%></a>
		<%	else %>
					<a href="tasks/undelete.asp?id=<%=rs("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Restore")%>');"><%=dictLanguage("Undelete")%></a>
		<%	end if %>
				</td>
			</tr>
		<%		rs.movenext
			wend
			rs.close
			set rs = nothing %>
			<tr><td valign="top" colspan="11"><b><%=dictLanguage("Inactive_Tasks")%></b></td></tr>
		<%
			sql = sql_GetTaskViewByProjectID(0,pID) '(active, projectID)
			Call RunSQL(sql, rs)
			while not rs.eof
				boolDone = rs("Done") %>
			<tr<%if boolDone = 0 or not boolDone then%> bgcolor="#FFFF77" <%else%> bgcolor="#EEEEEE" <%end if%>>
				<td valign="top" class="small"><a href="tasks/edit.asp?id=<%=rs("Task_ID")%>"><%=dictLanguage("Edit")%></a></td>
				<td valign="top" align="center"><a href="tasks/toggle_done.asp?id=<%=rs("Task_ID")%>&value=
		<%		if rs("Done") = 0 then
					response.write "no"
				else
					response.write "yes"
				end if %>&referring_page=view2.asp&sql=<%=Server.URLEncode(sql)%>">
		<%		if boolDone = True then
					response.write "<img src='" & gsSiteRoot & "images/done.gif' border=0>"
				else
					response.write "<img src='" & gsSiteRoot & "images/notDone.gif' border=0>"
				end if %>
				</td>
				<td valign=top class="small"><%=rs("DateDue")%></td>
				<td valign="top" align=center class="small"> 
		<%		select case rs("Priority")
					case 1
						response.write dictLanguage("Low")
					case 2
						response.write "<font color=""blue""><b>" & dictLanguage("Medium") & "</b></font>"
					case 3
						response.write "<font color=""red""><b>" & dictLanguage("High") & "</b></font>"
					case 0 
						response.write "<font color=""green""><b>" & dictLanguage("On_Hold") & "</b></font>"
					case else
						response.write "<font color=""blue""><b>" & dictLanguage("Medium") & "</b></font>"
				end select%>
				</td>
				<td valign=top class="small"><%=rs("EmployeeName")%></td>
				<td valign=top class="small"><%=rs("EmployeeName_1")%></td>
				<td valign=top class="small"><%=rs("Desc1")%></td>
				<td valign=top class="small"><%=rs("EstimatedHours")%></td>
		<%
				sqlHours = sql_GetTimecardHoursByTaskID(rs("Task_ID"))
				Call RunSQL(sqlHours, rsHours)
				if rsHours("TimeCharged")<>"" then
					timeCharged = rsHours("TimeCharged")
				else 
					timeCharged = 0
				end if
				estHoursLeft = rs("EstimatedHours") - timeCharged
				rsHours.close
				set rsHours = nothing %>
				<td valign=top class="small"><font <%if estHoursLeft<0 then%>color="#FF0033" <%end if%>class="small"><%=estHoursLeft%></font></td>
				<td valign=top class="small"><%=rs("DateCreated")%></td>
				<td valign="top" class="small">
		<%	if rs("show") = "1" or rs("show") = "-1" or rs("show") then %>			
					<a href="tasks/delete.asp?id=<%=rs("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Delete")%>');"><%=dictLanguage("Delete")%></a>
		<%	else %>
					<a href="tasks/undelete.asp?id=<%=rs("Task_ID")%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Task_Restore")%>');"><%=dictLanguage("Undelete")%></a>
		<%	end if %>
				</td>
			</tr>
		<%		rs.movenext
			wend
			rs.close
			set rs = nothing %>			
		</table>	
	
	</td>
</tr>


<tr><td colspan="2">&nbsp;</td></tr>
<tr bgcolor="<%=gsColorHighlight%>"><td class="columnheader" colspan="2"><%=dictLanguage("Project_Timecards")%></td></tr>
<tr>
	<td colspan="2">
	<%
		'''''''''''''''''''''''''''
		'load in employee id lookup table
		'''''''''''''''''''''''''''
		dim employee(150)

		sql = sql_GetAllEmployees()
		Call RunSQL(sql, rs)
		do while not rs.eof
			for x=0 to UBound(employee)
				if x=cint(rs("Employee_ID")) then
					employee(x)=rs("EmployeeLogin")
					exit for
				end if
			next
			rs.movenext
		loop
		rs.close
		set rs = nothing

		'''''''''''''''''''''''''''
		'build query from form info or querystring info
		'''''''''''''''''''''''''''
		Client_ID = cID  'set client id here
		Project_ID = pID	'set project id here
		strCompanyClause = "tbl_Clients.Client_ID = " & Client_ID & " AND " & "tbl_Projects.Project_ID = " & Project_ID

		'put the clauses together
		strWhereClause = " WHERE tbl_Projects.ProjectType_ID = tbl_ProjectTypes.ProjectType_ID and tbl_Clients.Client_ID = tbl_Projects.Client_ID and tbl_Employees.Employee_ID = tbl_TimeCards.Employee_ID and tbl_TimeCardTypes.TimeCardType_ID = tbl_TimeCards.TimeCardType_ID and tbl_Projects.Project_ID = tbl_TimeCards.Project_ID "
		strWhereClause = strWhereClause & " AND " & strCompanyClause

		strSQL = "SELECT tbl_Clients.Client_Name, tbl_Clients.Rep_ID, " & _
			"tbl_Employees.EmployeeLogin, tbl_Employees.Employeetype_id, tbl_TimeCards.TimeAmount, " & _
			"tbl_TimeCards.DateWorked, tbl_TimeCards.Cost, tbl_TimeCards.Rate, tbl_TimeCards.Employee_ID, " & _
			"tbl_Timecards.timecardtype_ID, tbl_TimeCardTypes.TimeCardTypeDescription, " & _
			"tbl_TimeCards.[Non-Billable], tbl_TimeCards.Reconciled, tbl_Timecards.TimeCard_ID, " & _
			"tbl_TimeCards.WorkDescription, tbl_Projects.Description, tbl_Projects.ProjectType_ID, " & _
			"tbl_timecards.projectphaseID, tbl_ProjectTypes.ProjectTypeDescription as ProjType, " & _
			"tbl_TimeCards.Project_ID " & _
			"FROM tbl_ProjectTypes, tbl_Clients, tbl_projects, tbl_TimeCardTypes, tbl_timeCards, tbl_Employees " & _
			strWhereClause & _
			" ORDER BY tbl_Projects.ProjectType_ID, tbl_Clients.Client_Name, tbl_Projects.Description, " & _
			"tbl_timecards.[non-billable] desc, tbl_Employees.EmployeeLogin, tbl_TimeCards.DateWorked"
		'response.write strSQL & "<BR>"
		Call runSQL(strSQL, adoRS)
		%>

		<table cellpadding="2" cellspacing="2" border="0" width="100%">
		<%

		intTotal = 0
		intGrandTotal = 0
		intBillable = 0
		intBillableTotal = 0
		intNonBillable = 0
		intNonBillableTotal = 0

		strProjectName = ""
		strClientName = ""
		boolFirst = True

		strBGColor = "#FFFFFF"
		Do While Not adoRS.EOF
			boolDataExists = True
			If adoRS("Client_Name") <> strClientName or adoRS("Description") <> strProjectName then
				'does not write unless the table header has been written
				If Not boolFirst Then
		%>
			<tr bgcolor="<%=gsColorHighlight%>"><td colspan="8"></td></tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="bolddark"><%=dictLanguage("Billable")%>:</td>
				<td><b><%=intBillable%></b></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="bolddark"><%=dictLanguage("Non-Billable")%>:</td>
				<td><b><%=intNonBillable%></b></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="bolddark"><%=dictLanguage("Total")%>:</td>
				<td><b><%=intTotal%></b></td>
			</tr>
			<tr bgcolor="<%=gsColorHighlight%>"><td colspan="8"></td></tr>
		</Table>
		<%
					intTotal = 0
					intBillable = 0
					intNonBillable = 0
				End If %>

		<table border="0" width=100% cellspacing="2" cellpadding="2" style="border=2px; border-color=<%=gsColorHighlight%>; border-style: solid;">
			<tr bgcolor="<%=gsColorHighlight%>">
				<td class="columnheader"><!--this cell will contain the edit link--></td>
				<td class="columnheader" align="center" width=50><%=dictLanguage("Date")%></td>
				<td class="columnheader" align="center" width=50><%=dictLanguage("Employee")%></td>
				<td class="columnheader" align="center" width=50><%=dictLanguage("Time")%></td>
				<td class="columnheader" align="center" width=50><%=dictLanguage("Type")%></td>
				<td class="columnheader" align="center" width=50><%=dictLanguage("Billable")%></td>
				<td class="columnheader" align="center" width=50><%=dictLanguage("Phase")%></td>
				<td class="columnheader" align="center" width="100%"><%=dictLanguage("Comments")%></td>
			</tr>
		<%
				strClientName = adoRS("client_name")
				strProjectName = adoRS("Description")
				boolFirst = False
			End If

			if strBGColor = "#FFFFFF" then
				strBGColor = "#EEEEEE"
			else
				strBGColor = "#FFFFFF"
			end if %>
			<tr bgcolor="<%=strBGColor%>">
				<td align="center" class="small" valign="top"><a href="<%=gsSiteRoot%>timecards/timecard-edit.asp?timecard_id=<%=Server.URLEncode(adoRS("timeCard_ID"))%>">edit</a></td>
				<td align="center" class="small" valign="top"><%=adoRS("dateworked")%></td>
				<td align="center" class="small" valign="top"><%=adoRS("employeelogin")%></td>
				<td align="center" class="small" valign="top"><%=adoRS("timeamount")%></td>
				<td align="center" class="small" valign="top"><%=adoRS("timecardtypedescription")%></td>
				<td align="center" class="small" valign="top">
		<%	if adoRS("Non-Billable") = "True" then
				response.write dictLanguage("No")
			else
				response.write dictLanguage("Yes")
			end if%></td>
				<td align="center" class="small" valign="top">
		<%	if adoRS("ProjectPhaseID")="0" then 
				response.write dictLanguage("None")
			elseif adoRS("ProjectPhaseID")<>"" then
				response.write GetProjectPhaseName(adoRS("ProjectPhaseID"))
			else
				response.write "&nbsp;"
			end if %></td>
				<td class="small" valign="top"><%=adoRS("workdescription")%></td>
			</tr>
		<%
		'Update Counters for bottom of the page

		if adoRS("Non-Billable") = "True" then
		    intNonBillable = intNonBillable + adoRS("timeamount")
		    intNonBillableTotal = intNonBillableTotal + adoRS("timeamount")
		else
			intBillable = intBillable + adoRS("timeamount")
		    intBillableTotal = intBillableTotal + adoRS("timeamount")
			intRateTotal = intRateTotal + adoRS("rate")
		End If
			intTotal = intTotal + adoRS("timeamount")
		    intGrandTotal = intGrandTotal + adoRS("timeamount")
			intCostTotal = intCostTotal + adoRS("cost")
		adoRS.MoveNext
		Loop

		'Write out end data for last table
		If boolDataExists Then
		%>
			<tr bgcolor="<%=gsColorHighlight%>"><td colspan="8"></td></tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="bolddark"><%=dictLanguage("Billable")%>:</td>
				<td><b><%=intBillable%></b></td>
				<td colspan=5></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="bolddark"><%=dictLanguage("Non-Billable")%>:</td>
				<td><b><%=intNonBillable%></b></td>
				<td colspan=5></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td class="bolddark"><%=dictLanguage("Total")%>:</td>
				<td><b><%=intTotal%></b></td>
				<td colspan=5>&nbsp;</td>	
			</tr>
			<tr bgcolor="<%=gsColorHighlight%>"><td colspan="8"></td></tr>		
		</Table>

		<%	
		End If
		adoRS.Close
		set adoRS = nothing
		%>

	</td>
</tr>
</table>

<% end if %>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<br>

<%
function GetProjectPhaseName(id)
	sql = sql_GetProjectPhasesByID(id)
	set rsPhase = Conn.execute(sql)
	if not rsPhase.eof then
		rsPhase.movefirst
		phaseName = rsPhase("projectphaseName")
	else
		phaseName = "None"
	end if
	rsPhase.close
	set rsPhase = nothing
	GetProjectPhaseName = phaseName
end function

function SpaceEncode(strText)
	if strText <> "" and isNull(strText) = False then
		strText = replace(strText," ","%20")
	end if
	SpaceEncode = strText
end function

function SpaceDecode(strText)
	if strText <> "" and isNull(strText) = False then
		strText = replace(strText,"%20"," ")
	end if
	SpaceDecode = strText
end function

Function GetEmpName(ID)
	if ID <> "" then
		sql = sql_GetEmployeesByID(ID)
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
	Call RunSQL(sql, rs)
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

%>
<!--#include file="../includes/main_page_close.asp"-->