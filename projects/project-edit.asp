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
project_id = request("project_id")
if project_id = "" then
	response.redirect ("main.asp")
end if

Sub EmployeeDropDown(Name,ID,boolProject)
	response.write("<select name='" & ID & "' size='1' class='formstyleLong'>")
	response.write("<option value='0' selected>" & dictLanguage("Select") & "&nbsp;" & Name & "</option>")
	rsEmployees.MoveFirst
	while not rsEmployees.EOF
		intEmployeeId = CINT(rsEmployees("Employee_ID"))
		if boolProject then
			if isNumeric(rsProject(ID)) then
				intEmployee_ID = CINT(rsProject(ID))
			end if		
		else
			if boolThisBudget then
				if isNumeric(rsBudget("employee_id")) then
					intEmployee_ID = CINT(rsBudget("employee_id"))
				else
					intEmployee_ID = 0
				end if
			else
				intEmployee_ID = 0
			end if
		end if
		response.write("<option  value='" & intEmployeeId & "'")
		if intEmployeeId = intEmployee_ID then
			response.write(" selected")
		end if
		response.write(">" & rsEmployees("EmployeeName") & "</option>")
		rsEmployees.MoveNext
  	wend
    response.write("</select>")
End Sub

sql = sql_GetProjectsByID(project_id)
Call RunSQL(sql, rsProject)
	start_date = rsProject("Start_Date")
	if start_date = "" then
		start_date = date()
	else
		if not isdate(start_date) then
			start_date = date()
		end if
	end if
	launch_date = rsProject("Launch_Date")
	if launch_date = "" then
		launch_date = date()
	else
		if not isdate(launch_date) then
			launch_date = date()
		end if
	end if
	
sql = sql_GetAllClients()
Call RunSQL(sql, rsClientList)

sql = sql_GetProjectTypes()
Call RunSQL(sql, rsProjectType)

sql = sql_GetEmployeeTypes()
Call RunSQL(sql, rsRates)

sql = sql_GetActiveEmployees()
Call RunSQL(sql, rsEmployees)

sql = sql_GetProjectBudgetsByID(project_id)
Call RunSQL(sql, rsBudget)
%>

<!--#include file="../includes/main_page_open.asp"-->

<form method="post" action="project-edit-processor.asp" name="strForm" id="strForm">
<input type="hidden" name="project_id" value="<%=project_id%>">
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr><td colspan="4" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("Project_Details")%></b></td></tr>
		<tr><td colspan="4" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
<%	if rsProject("forum_id")<>"" and gsDiscussion then %>
		<tr><td colspan="4"><a href="<%=gsSiteRoot%>forum/default.asp?fid=<%=rsProject("forum_id")%>"><%=dictLanguage("Project_Discussion")%></a></td></tr> 
<%	end if %>	
<%	if rsProject("folder_id")<>"" and gsDiscussion then %>
		<tr><td colspan="4"><a href="<%=gsSiteRoot%>repository/default.asp?folder=<%=rsProject("folder_id")%>"><%=dictLanguage("Project_Documents")%></a></td></tr> 
<%	end if %>	
<%  if gsProjectQuotes then %>
		<tr><td colspan="2" valign="top"><a href="javascript: popup('project-quote.asp?pid=<%=project_id%>');"><%=dictLanguage("Create_Project_Quote")%></a></td>
<%		sql = sql_GetProjectQuotesByProjectID(project_id)
		Call runSQL(sql, rsQuotes)
		if not rsQuotes.eof then %>
			<td colspan="2" align="right" valign="top"><%=dictLanguage("Existing_Project_Quotes")%><br>
<%			while not rsQuotes.eof %>
				<a href="javascript: popup('project-quote.asp?Submit=view&pid=<%=project_id%>&pqid=<%=rsQuotes("projectquote_id")%>');"><%=dictLanguage("Date")%>: <%=rsQuotes("dateentered")%>, <%=dictLanguage("Total")%>: <%=formatnumber(rsQuotes("total_price"),2,-1,0,0)%></a></br>	
<%				rsQuotes.movenext 
			wend
			Response.Write "</td></tr>"
		else
			Response.Write "<td colspan=""2"">&nbsp;</td></tr>"
		end if
		rsQuotes.close
		set rsQuotes = nothing
		Response.Write "<tr><td colspan=""4"">&nbsp;</td></tr>"
    end if %>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Name")%>:</b><font class="alert">*</font></td>
			<td colspan="3"><input name="Description" size="50" value="<%=rsProject("Description")%>" class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client")%>:</b><font class="alert">*</font></td>
			<td colspan="3">
				<select name="Client_ID" size="1" class="formstyleLong">
					<option value="" selected><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Client")%></option>
<%
while not rsClientList.EOF
	intClientId = CINT(rsClientList("Client_ID"))
	if IsNumeric(rsProject("Client_ID")) = True then
		intClient_ID = CINT(rsProject("Client_ID"))
	end if
%>
					<option value="<%=intClientId%>" <%if intClientId = intClient_ID then%> selected <%end if%>><%=rsClientList("Client_Name")%></option>
<%	rsClientList.MoveNext
wend%>
				</select>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Account_Rep")%>:</b><font class="alert">*</font></td>
			<td colspan="3"><%Call EmployeeDropDown(dictLanguage("Account_Rep"),"AccountExec_ID",TRUE)%></td>
	    </tr>	
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Work_Order")%>:</b><font class="alert">*</font></td>
			<td colspan="3"><input name="WorkOrder_Number" class="formstyleShort" size="20" value="<%=rsProject("WorkOrder_Number")%>" maxlength="50"> </td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Date")%>:</b><font class="alert">*</font></td>
			<td colspan="3">
				<input name="Start_Date" class="formstyleShort" size="20" value="<%=start_date%>" onkeypress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Start_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Date")%>:</b><font class="alert">*</font></td>
			<td colspan="3">
				<input name="Launch_Date" class="formstyleShort" size="20" value="<%=launch_date%>" onkeypress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Launch_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Project_Type")%>:</b><font class="alert">*</font></td>
			<td colspan="3">
				<select name="ProjectType_ID" size="1" class="formstyleLong">
					<option value="" selected><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Project_Type")%></option>
<%
while not rsProjectType.EOF
	intProjectTypeId = CINT(rsProjectType("ProjectType_ID"))
	if isNumeric(rsProject("ProjectType_ID")) = True then
		intProjectType_ID = CINT(rsProject("ProjectType_ID"))
	end if
%>			
				 	<option value="<%=intProjectTypeId%>" <%if intProjectTypeId = intProjectType_ID then%> selected <%end if%>><%=rsProjectType("ProjectTypeDescription")%></option>		
<%	rsProjectType.Movenext
wend%>
				</select>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Status")%>:</b></td>
			<td colspan="3">
				<input name="color" type="radio" value="Green"<%if rsProject("color")="Green" then%> Checked<%end if%>> <%=dictLanguage("Green")%><br>
				<input name="color" type="radio" value="Blue"<%if rsProject("color")="Blue" then%> Checked<%end if%>> <%=dictLanguage("Blue")%><br>
				<input name="color" type="radio" value="Red"<%if rsProject("color")="Red" then%> Checked<%end if%>> <%=dictLanguage("Red")%>
			</td>
		</tr> 
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Active")%>:</b></td>
			<td colspan="3">
				<Select name="active" size="1" class="formstyleLong">
	  				<option value="1"<%if rsProject("active")=TRUE then%> Selected<%end if%>><%=dictLanguage("True")%></option>
	      			<option value="0"<%if rsProject("active")=FALSE then%> Selected<%end if%>><%=dictLanguage("False")%></option>
				</select>
			</td>
		</tr>
		<tr>
			<td valign="top"><b class="bolddark"><%=dictLanguage("Comments")%>:</b></td>
			<td colspan="3"><textarea name="comments" rows="4" class="formstyleLong"><%=rsProject("comments")%></textarea></td>
		</tr> 	 	

		<tr>
			<td>&nbsp;</td>
			<td align="center"><%=dictLanguage("Rate")%></td>
			<td align="center"><%=dictLanguage("Hours")%></td>
			<td><%=dictLanguage("Core_Team")%></td>
		</tr>
	
<%while not rsRates.eof 
	boolThisBudget = FALSE
	if not rsBudget.eof then
		if rsBudget("employeetype_id") = rsRates("employeetype_id") then 
			boolThisBudget = TRUE %>
		<tr>
			<td><b class="bolddark"><%=rsRates("EmployeeType")%></b></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xRate" class="formstyleTiny" size="4" value="<%=rsBudget("rate")%>"></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xHours" class="formstyleTiny" size="4" value="<%=rsBudget("Hours")%>"></td>
			<td><%Call EmployeeDropDown(dictLanguage("Team_Member"),rsRates("EmployeeType_ID") & "xEmployee_ID", FALSE)%></td>
		</tr>
<%			rsBudget.movenext
		else %>
		<tr>
			<td><b class="bolddark"><%=rsRates("EmployeeType")%></b></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xRate" class="formstyleTiny" size="4" value="<%=rsRates("Rate")%>"></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xHours" class="formstyleTiny" size="4" value="0"></td>
			<td><%Call EmployeeDropDown(dictLanguage("Team_Member"),rsRates("EmployeeType_ID") & "xEmployee_ID", FALSE)%></td>
		</tr>		
<%		end if
	else %>
		<tr>
			<td><b class="bolddark"><%=rsRates("EmployeeType")%></b></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xRate" class="formstyleTiny" size="4" value="<%=rsRates("Rate")%>"></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xHours" class="formstyleTiny" size="4" value="0"></td>
			<td><%Call EmployeeDropDown(dictLanguage("Team_Member"),rsRates("EmployeeType_ID") & "xEmployee_ID", FALSE)%></td>
		</tr>	
<%	end if
	rsRates.movenext 
wend %>
	</table>
	<%if session("permProjectsEdit") then%>
	<p align="center"><input type="Submit" name="Submit" value="Submit" class="formButton"> <input type="Reset" name="Reset" value="Reset" class="formButton"> </p>
	<%end if%>
</form>

<%
rsEmployees.Close
set rsEmployees=nothing
rsClientList.Close
set rsClientList=nothing
rsProjectType.Close
set rsProjectType=nothing
rsProject.close
set rsProject = nothing
rsRates.close
set rsRates = nothing
rsBudget.close
set rsBudget = nothing %>

<!--#include file="../includes/main_page_close.asp"-->