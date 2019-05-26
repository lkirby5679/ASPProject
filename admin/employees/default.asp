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

<%
strAct		= trim(request("act"))
strEmpID	= trim(request("id"))
%>


<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<script language = "Javascript">

function checkAll(dis) {
	if (dis.checked) {
		document.strForm.permAdmin.checked = true;
		document.strForm.permAdminCalendar.checked = true;
		document.strForm.permAdminDatabaseSetup.checked = true;
		document.strForm.permAdminEmployees.checked = true;
		document.strForm.permAdminEmployeesPerms.checked = true;
		document.strForm.permAdminFileRepository.checked = true;
		document.strForm.permAdminForum.checked = true;
		document.strForm.permAdminNews.checked = true;	
		document.strForm.permAdminResources.checked = true;	
		document.strForm.permAdminSurveys.checked = true;
		document.strForm.permAdminThoughts.checked = true;
		document.strForm.permClientsAdd.checked = true;
		document.strForm.permClientsEdit.checked = true;
		document.strForm.permClientsDelete.checked = true;	
		document.strForm.permProjectsAdd.checked = true;
		document.strForm.permProjectsEdit.checked = true;
		document.strForm.permProjectsDelete.checked = true;
		document.strForm.permTasksAdd.checked = true;
		document.strForm.permTasksEdit.checked = true;
		document.strForm.permTasksDelete.checked = true;
		document.strForm.permTimecardsAdd.checked = true;
		document.strForm.permTimecardsEdit.checked = true;
		document.strForm.permTimecardsDelete.checked = true;
		document.strForm.permTimesheetsEdit.checked = true;
		document.strForm.permForumAdd.checked = true;
		document.strForm.permRepositoryAdd.checked = true;
		document.strForm.permPTOAdmin.checked = true;
	}
	else {
		document.strForm.permAdmin.checked = false;
		document.strForm.permAdminCalendar.checked = false;
		document.strForm.permAdminDatabaseSetup.checked = false;
		document.strForm.permAdminEmployees.checked = false;
		document.strForm.permAdminEmployeesPerms.checked = false;
		document.strForm.permAdminFileRepository.checked = false;
		document.strForm.permAdminForum.checked = false;
		document.strForm.permAdminNews.checked = false;	
		document.strForm.permAdminResources.checked = false;
		document.strForm.permAdminSurveys.checked = false;
		document.strForm.permAdminThoughts.checked = false;
		document.strForm.permClientsAdd.checked = false;
		document.strForm.permClientsEdit.checked = false;
		document.strForm.permClientsDelete.checked = false;	
		document.strForm.permProjectsAdd.checked = false;
		document.strForm.permProjectsEdit.checked = false;
		document.strForm.permProjectsDelete.checked = false;
		document.strForm.permTasksAdd.checked = false;
		document.strForm.permTasksEdit.checked = false;
		document.strForm.permTasksDelete.checked = false;
		document.strForm.permTimecardsAdd.checked = false;
		document.strForm.permTimecardsEdit.checked = false;
		document.strForm.permTimecardsDelete.checked = false;
		document.strForm.permTimesheetsEdit.checked = false;
		document.strForm.permForumAdd.checked = false;
		document.strForm.permRepositoryAdd.checked = false;	
		document.strForm.permPTOAdmin.checked = false;
	}
}
</script>

<table border="0" cellspacing="2" cellpadding="2">
<tr><td valign="top">

<table border="0" cellspacing="2" cellpadding="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td colspan="4" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Employees")%></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=add"><IMG SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_New_Employee")%>" border="0"></a></td>
	</tr>
<% strSQL = sql_GetAllEmployees()
   Call RunSQL(strSQL, rs)
   do while not rs.eof
		empID = trim(rs("employee_id"))
		empName = trim(rs("employeename"))
		blnActive = trim(rs("active")) %>
	<tr>
		<td valign="top"><%=empID%></td>
		<%if blnActive = "True" then%>
		<td valign="top" nowrap><a href="javascript: popupwide('<%=gsSiteRoot%>employees/default.asp?EmployeeID=<%=empID%>');"><%=empName%></a></td>
		<%else%>
		<td valign="top" nowrap><a href="javascript: popupwide('<%=gsSiteRoot%>employees/default.asp?EmployeeID=<%=empID%>');"><font color=red><B><%=empName%> (<%=dictLanguage("Inactive")%>)</b></font></a></td>
		<%end if%>
		<td valign="top"><a href="default.asp?act=edit&id=<%=empID%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Employee")%>"></a></td>
		<%if Session("permAdminEmployeesPerms") then%><td valign="top"><a href="default.asp?act=perms&id=<%=empID%>"><IMG SRC="../images/permissions.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Permissions")%>"></a></td><%end if%>
	</tr>
<%		rs.movenext
	loop 
	rs.close
	set rs = nothing %>
</table>

</td><td align="center" width="100%" valign="top">

<table border="0" cellspacing="2" cellpadding="2" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td align="Center" class="homeheader"><%=dictLanguage("Workspace")%></td>
	</tr>
</table>

<%if strAct = "add" then %>

<form name="strForm" id="strForm" action="hnd_Edit.asp" method="POST" ENCTYPE="multipart/form-data">
<input type="hidden" name="act" value="add">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Add_New_Employee")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_ID")%>:</b></td>
		<td valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Name")%>:</b></td>
		<td valign="top"><input type="text" name="empName" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Title")%>:</b></td>
		<td valign="top"><input type="text" name="empTitle" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Login")%>:</b>&nbsp;(<%=dictLanguage("Initials")%>)</td>
		<td valign="top"><input type="text" name="empLogin" value="" class="formstyleShort" maxLength="4"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Password")%>:</b></td>
		<td valign="top"><input type="text" name="empPassword" value="" class="formstyleShort" maxLength="20"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Image")%>:</b></td>
		<td valign="top"><input type="file" name="empImage" value="" class="formstyleLong" maxLength="255"></td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Email")%>:</b></td>
		<td valign="top"><input type="text" name="empEmail" value="" class="formstyleLong" maxLength="100"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Alt_Email")%>:</b></td>
		<td valign="top"><input type="text" name="empEmailAlt" value="" class="formstyleLong" maxLength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("IM_Name")%>:</b></td>
		<td valign="top"><input type="text" name="empIMName" value="" class="formstyleLong" maxLength="50"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Commission_Percentage")%>:</b></td>
		<td valign="top"><input type="text" name="empCommissionRate" value="0" class="formstyleShort" maxLength="3">%</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Monthly_Sales_Goal")%>:</b></td>
		<td valign="top">$<input type="text" name="empSalesGoal" value="0" class="formstyleShort" maxLength="7"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Monthly_Production_Goal")%>:</b></td>
		<td valign="top"><input type="text" name="empProductionGoal" value="0" class="formstyleShort" maxLength="7">&nbsp;Hrs</td>
	</tr>								
	<tr>
		<td valign="top"><b><%=dictLanguage("Start_Date")%>:</b></td>
		<td valign="top"><input type="text" name="empStartDate" value="<%=date%>" class="formstyleShort" maxLength="10" onKeyPress="txtDate_onKeypress();">
			<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','empStartDate',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Hourly")%>:</b></td>
		<td valign="top"><input type="checkbox" name="empHourly" value="1"></td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Active")%>:</b></td>
		<td valign="top"><input type="checkbox" name="empActive" value="1"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_Type")%>:</b></td>
		<td valign="top"><select name="empType" class="formstylelong">
							<option value="0"><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee_Type")%></option>
<%	sql = sql_GetEmployeeTypes()
	Call RunSQL(sql, rs)
	while not rs.eof
		Response.Write "<option value=""" & trim(rs("employeetype_id")) & """>" & trim(rs("employeetype")) & "</option>"
		rs.movenext
	wend
	rs.close
	set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Department")%>:</b></td>
		<td valign="top"><input type="text" name="empDepartment" value="" class="formstyleLong" maxLength="50"></td>
	</tr>					
	<tr>
		<td valign="top"><b><%=dictLanguage("Reports_To")%>:</b></td>
		<td valign="top"><select name="empReportsTo" class="formstylelong">
							<option value="0"><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee")%></option>
<%	sql = sql_GetAllEmployees()
	Call RunSQL(sql, rs)
	while not rs.eof
		Response.Write "<option value=""" & trim(rs("employee_id")) & """>" & trim(rs("employeename")) & "</option>"
		rs.movenext
	wend
	rs.close
	set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Birth_Date")%>:</b></td>
		<td valign="top"><input type="text" name="empBirthDate" value="<%=date%>" class="formstyleShort" maxLength="10">
			<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','empBirthDate',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Home_Phone")%>:</b></td>
		<td valign="top"><input type="text" name="empHomePhone" value="" class="formstyleMed" maxLength="50"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Mobile_Phone")%>:</b></td>
		<td valign="top"><input type="text" name="empMobilePhone" value="" class="formstyleMed" maxLength="50"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Work_Phone")%>:</b></td>
		<td valign="top">
			<input type="text" name="empWorkPhone" value="" class="formstyleMed" maxLength="50">
			<%=dictLanguage("ext.")%><input type="text" name="empWorkPhoneExt" value="" class="formstyleTiny" maxLength="50">
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Voice_Mail")%>:</b></td>
		<td valign="top"><input type="text" name="empVoiceMail" value="" class="formstyleMed" maxLength="50"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Home_Address")%>:</b></td>
		<td valign="top"><input type="text" name="empHomeStreet1" value="" class="formstyleLong" maxLength="100"></td>
	</tr>	
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top"><input type="text" name="empHomeStreet2" value="" class="formstyleLong" maxLength="100"></td>
	</tr>			
	<tr>
		<td valign="top"><b><%=dictLanguage("City")%> / <%=dictLanguage("State")%> / <%=dictLanguage("Zip")%>:</b></td>
		<td valign="top">
			<input type="text" name="empHomeCity" value="" class="formstyleMed" maxLength="100">
			<input type="text" name="empHomeState" value="" class="formstyleTiny" maxLength="50">
			<input type="text" name="empHomeZip" value="" class="formstyleShort" maxLength="50">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>			
</table>
</form>

<%elseif strAct = "edit" and strEmpID <> "" then 
	sql = sql_GetEmployeesByID(strEmpID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		empID			= rs("employee_id")
		empName			= trim(rs("employeename"))
		empStartDate	= trim(rs("startdate"))
		empLogin		= trim(rs("employeelogin"))
		empTitle		= trim(rs("employeeTitle"))
		empEmail		= trim(rs("emailaddress"))
		empEmailAlt		= trim(rs("emailaddressalt"))
		empCommissionRate = rs("commissionRate")
		empSalesGoal	= rs("salesgoal")
		empProductionGoal = rs("productiongoal")
		empPassword		= trim(rs("password"))
		empActive		= rs("active")
		if empActive = 1 or empActive then
			empActive = TRUE
		else
			empActive = FALSE
		end if
		empHourly		= rs("Hourly") 
		if empHourly = 1 or empHourly then
			empHourly = TRUE
		else
			empHourly = FALSE
		end if		
		empType			= rs("employeetype_id")
		empReportsTo	= rs("reportsto")
		empDepartment	= rs("department")
		empBirthDate	= rs("bdate")
		empHomePhone	= trim(rs("HomePhone"))
		empMobilePhone	= trim(rs("MobilePhone"))
		empWorkPhone	= trim(rs("WorkPhone"))
		empWorkPhoneExt = trim(rs("WorkPhoneExt"))
		empVoiceMail	= trim(rs("VoiceMail"))
		empHomeStreet1  = trim(rs("HomeStreet1"))
		empHomeStreet2  = trim(rs("HomeStreet2"))
		empHomeCity		= trim(rs("HomeCity"))
		empHomeState	= trim(rs("HomeState"))
		empHomeZip		= trim(rs("HomeZip"))
		empIMName		= trim(rs("IMNAME"))
		empImage		= trim(rs("image"))		
	else
		Response.Redirect "default.asp"
	end if
	rs.close
	set rs = nothing


%>

<form name="strForm" id="strForm" action="hnd_Edit.asp" method="POST" ENCTYPE="multipart/form-data" >
<input type="hidden" name="act" value="edit">
<input type="hidden" name="empID" value="<%=empID%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Edit_Employee")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_ID")%>:</b></td>
		<td valign="top"><%=empID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Name")%>:</b></td>
		<td valign="top"><input type="text" name="empName" value="<%=empName%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Title")%>:</b></td>
		<td valign="top"><input type="text" name="empTitle" value="<%=empTitle%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Login")%>:</b>&nbsp;(initials)</td>
		<td valign="top"><input type="text" name="empLogin" value="<%=empLogin%>" class="formstyleShort" maxLength="4"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Password")%>:</b></td>
		<td valign="top"><input type="password" name="empPassword" value="<%=empPassword%>" class="formstyleShort" maxLength="20"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Current_Image")%>:</b></td>
		<td valign="top"><a href="<%=gsSiteRoot%>employees/images/<%=empImage%>" target="_blank"><%=empImage%></a></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("New_Image")%>:</b></td>
		<td valign="top">
			<input type="hidden" name="empImageOld" value="<%=empImage%>">
			<input type="file" name="empImage" value="" class="formstyleLong" maxLength="255">
		</td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Email")%>:</b></td>
		<td valign="top"><input type="text" name="empEmail" value="<%=empEmail%>" class="formstyleLong" maxLength="100"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Alt_Email")%>:</b></td>
		<td valign="top"><input type="text" name="empEmailAlt" value="<%=empEmailAlt%>" class="formstyleLong" maxLength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("IM_Name")%>:</b></td>
		<td valign="top"><input type="text" name="empIMName" value="<%=empIMName%>" class="formstyleLong" maxLength="50"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Commission_Percentage")%>:</b></td>
		<td valign="top"><input type="text" name="empCommissionRate" value="<%=empCommissionRate%>" class="formstyleShort" maxLength="3">%</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Monthly_Sales_Goal")%>:</b></td>
		<td valign="top">$<input type="text" name="empSalesGoal" value="<%=empSalesGoal%>" class="formstyleShort" maxLength="7"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Monthly_Production_Goal")%>:</b></td>
		<td valign="top"><input type="text" name="empProductionGoal" value="<%=empProductionGoal%>" class="formstyleShort" maxLength="7">&nbsp;Hrs</td>
	</tr>								
	<tr>
		<td valign="top"><b><%=dictLanguage("Start_Date")%>:</b></td>
		<td valign="top"><input type="text" name="empStartDate" value="<%=empStartDate%>" class="formstyleShort" maxLength="10" onKeyPress="txtDate_onKeypress();">
			<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','empStartDate',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Hourly")%>:</b></td>
		<td valign="top"><input type="checkbox" name="empHourly" value="1" <%if empHourly then Response.Write "Checked"%>></td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Active")%>:</b></td>
		<td valign="top"><input type="checkbox" name="empActive" value="1" <%if empActive then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_Type")%>:</b></td>
		<td valign="top"><select name="empType" class="formstylelong">
							<option value="0"><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee_Type")%></option>
<%	sql = sql_GetEmployeeTypes()
	Call RunSQL(sql, rs)
	while not rs.eof
		Response.Write "<option value=""" & trim(rs("employeetype_id")) & """"
		if trim(rs("employeetype_id")) = trim(empType) then 
			Response.Write " Selected"
		end if 
		Response.Write ">" & trim(rs("employeetype")) & "</option>"
		rs.movenext
	wend
	rs.close
	set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Department")%>:</b></td>
		<td valign="top"><input type="text" name="empDepartment" value="<%=empDepartment%>" class="formstyleLong" maxLength="50"></td>
	</tr>					
	<tr>
		<td valign="top"><b><%=dictLanguage("Reports_To")%>:</b></td>
		<td valign="top"><select name="empReportsTo" class="formstylelong">
							<option value="0"><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee")%></option>
<%	sql = sql_GetAllEmployees()
	Call RunSQL(sql, rs)
	while not rs.eof
		Response.Write "<option value=""" & trim(rs("employee_id")) & """"
		if trim(rs("employee_id")) = trim(empReportsTo) then
			Response.Write " Selected"
		end if
		Response.Write ">" & trim(rs("employeename")) & "</option>"
		rs.movenext
	wend
	rs.close
	set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Birth_Date")%>:</b></td>
		<td valign="top"><input type="text" name="empBirthDate" value="<%=empBirthDate%>" class="formstyleShort" maxLength="10">
			<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','empBirthDate',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Home_Phone")%>:</b></td>
		<td valign="top"><input type="text" name="empHomePhone" value="<%=empHomePhone%>" class="formstyleMed" maxLength="50"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Mobile_Phone")%>:</b></td>
		<td valign="top"><input type="text" name="empMobilePhone" value="<%=empMobilePhone%>" class="formstyleMed" maxLength="50"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Work_Phone")%>:</b></td>
		<td valign="top">
			<input type="text" name="empWorkPhone" value="<%=empWorkPhone%>" class="formstyleMed" maxLength="50">
			<%=dictLanguage("ext.")%><input type="text" name="empWorkPhoneExt" value="<%=empWorkPhoneExt%>" class="formstyleTiny" maxLength="50">
		</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Voice_Mail")%>:</b></td>
		<td valign="top"><input type="text" name="empVoiceMail" value="<%=empVoiceMail%>" class="formstyleMed" maxLength="50"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Home_Address")%>:</b></td>
		<td valign="top"><input type="text" name="empHomeStreet1" value="<%=empHomeStreet1%>" class="formstyleLong" maxLength="100"></td>
	</tr>	
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top"><input type="text" name="empHomeStreet2" value="<%=empHomeStreet2%>" class="formstyleLong" maxLength="100"></td>
	</tr>			
	<tr>
		<td valign="top"><b><%=dictLanguage("City")%> / <%=dictLanguage("State")%> / <%=dictLanguage("Zip")%>:</b></td>
		<td valign="top">
			<input type="text" name="empHomeCity" value="<%=empHomeCity%>" class="formstyleMed" maxLength="100">
			<input type="text" name="empHomeState" value="<%=empHomeState%>" class="formstyleTiny" maxLength="50">
			<input type="text" name="empHomeZip" value="<%=empHomeZip%>" class="formstyleShort" maxLength="50">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%		if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%		end if %>			
</table>
</form>

<% elseif strAct = "perms" and strEmpID <> "" then 
	empID = strEmpID
	sql = sql_GetEmployeesByID(strEmpID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		empName = trim(rs("employeename"))
	end if
	rs.close
	set rs = nothing		
	
	sql = sql_GetEmployeesPermissionsByID(strEmpID)
	Call RunSQL(sql, rsPerms)
	if not rsPerms.eof then
		permAll					= rsPerms("permAll")
		permAdmin				= rsPerms("permAdmin")
		permAdminCalendar		= rsPerms("permAdminCalendar")
		permAdminDatabaseSetup	= rsPerms("permAdminDatabaseSetup")
		permAdminEmployees		= rsPerms("permAdminEmployees")
		permAdminEmployeesPerms = rsPerms("permAdminEmployeesPerms")
		permAdminFileRepository	= rsPerms("permAdminFileRepository")
		permAdminForum			= rsPerms("permAdminForum")
		permAdminNews			= rsPerms("permAdminNews")
		permAdminResources		= rsPerms("permAdminResources")
		permAdminSurveys		= rsPerms("permAdminSurveys")
		permAdminThoughts		= rsPerms("permAdminThoughts")
		
		permClientsAdd			= rsPerms("permClientsAdd")
		permClientsEdit			= rsPerms("permClientsEdit")
		permClientsDelete		= rsPerms("permClientsDelete")
		permProjectsAdd			= rsPerms("permProjectsAdd")
		permProjectsEdit		= rsPerms("permProjectsEdit")
		permProjectsDelete		= rsPerms("permProjectsDelete")
		permTasksAdd			= rsPerms("permTasksAdd")
		permTasksEdit			= rsPerms("permTasksEdit")
		permTasksDelete			= rsPerms("permTasksDelete")
		permTimecardsAdd		= rsPerms("permTimecardsAdd")
		permTimecardsEdit		= rsPerms("permTimecardsEdit")
		permTimecardsDelete		= rsPerms("permTimecardsDelete")
		permTimesheetsEdit		= rsPerms("permTimesheetsEdit")
		permForumAdd			= rsPerms("permForumAdd")
		permRepositoryAdd		= rsPerms("permRepositoryAdd")
		permPTOAdmin			= rsPerms("permPTOAdmin")
	end if
	rsPerms.close
	set rsPerms = nothing %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndPerms">
<input type="hidden" name="id" value="<%=empID%>">
<table cellpadding="4" cellspacing="2" align="center">
	<tr><td colspan="4" align="center"><b><%=dictLanguage("Edit_Permissions")%></b></td></tr>
	<tr>
		<td valign="top" colspan="2"><b><%=dictLanguage("Employee_ID")%>:</b></td>
		<td valign="top" colspan="2"><%=empID%></td>
	</tr>
	<tr>
		<td valign="top" colspan="2"><b><%=dictLanguage("Name")%>:</b></td>
		<td valign="top" colspan="2"><%=empName%></td>
	</tr>
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr>
		<td valign="top" colspan="4" align="center">
			<b><%=dictLanguage("All_Permissions")%>:</b>&nbsp;
			<input type="checkbox" name="permAll" value="1" <%if permAll then Response.write "checked"%> onClick="checkAll(this);">
		</td>
	</tr>
	<tr><td colspan="4">&nbsp;</td></tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin_Permissions")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdmin" value="1" <%if permAdmin then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Clients")%> - <%=dictLanguage("Add")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permClientsAdd" value="1" <%if permClientsAdd then Response.write "checked"%>></td>		
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Calendar")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminCalendar" value="1" <%if permAdminCalendar then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Clients")%> - <%=dictLanguage("Edit")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permClientsEdit" value="1" <%if permClientsEdit then Response.write "checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Database_Setup")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminDatabaseSetup" value="1" <%if permAdminDatabaseSetup then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Clients")%> - <%=dictLanguage("Delete")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permClientsDelete" value="1" <%if permClientsDelete then Response.write "checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Employees")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminEmployees" value="1" <%if permAdminEmployees then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Projects")%> - <%=dictLanguage("Add")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permProjectsAdd" value="1" <%if permProjectsAdd then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("EmployeesPerms")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminEmployeesPerms" value="1" <%if permAdminEmployeesPerms then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Projects")%> - <%=dictLanguage("Edit")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permProjectsEdit" value="1" <%if permProjectsEdit then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Document_Repository")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminFileRepository" value="1" <%if permAdminFileRepository then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Projects")%> - <%=dictLanguage("Delete")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permProjectsDelete" value="1" <%if permProjectsDelete then Response.write "checked"%>></td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Discussion_Forum")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminForum" value="1" <%if permAdminForum then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Tasks")%> - <%=dictLanguage("Add")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permTasksAdd" value="1" <%if permTasksAdd then Response.write "checked"%>></td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminNews" value="1" <%if permAdminNews then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Tasks")%> - <%=dictLanguage("Edit")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permTasksEdit" value="1" <%if permTasksEdit then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Resources")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminResources" value="1" <%if permAdminResources then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Tasks")%> - <%=dictLanguage("Delete")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permTasksDelete" value="1" <%if permTasksDelete then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Surveys")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminSurveys" value="1" <%if permAdminSurveys then Response.write "checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Timecards")%> - <%=dictLanguage("Add")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permTimecardsAdd" value="1" <%if permTimecardsAdd then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin")%> - <%=dictLanguage("Thoughts")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permAdminThoughts" value="1" <%if permAdminThoughts then Response.write "checked"%>></td>	
		<td valign="top"><b><%=dictLanguage("Timecards")%> - <%=dictLanguage("Edit")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permTimecardsEdit" value="1" <%if permTimecardsEdit then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td colspan="2">&nbsp;</td>	
		<td valign="top"><b><%=dictLanguage("Timecards")%> - <%=dictLanguage("Delete")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permTimecardsDelete" value="1" <%if permTimecardsDelete then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td colspan="2">&nbsp;</td>	
		<td valign="top"><b><%=dictLanguage("Timesheets")%> - <%=dictLanguage("Edit")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permTimesheetsEdit" value="1" <%if permTimesheetsEdit then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td colspan="2">&nbsp;</td>	
		<td valign="top"><b><%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Add_Messages")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permForumAdd" value="1" <%if permForumAdd then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td colspan="2">&nbsp;</td>	
		<td valign="top"><b><%=dictLanguage("Document_Repository")%> - <%=dictLanguage("Upload_Documents")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permRepositoryAdd" value="1" <%if permRepositoryAdd then Response.write "checked"%>></td>
	</tr>	
	<tr>
		<td colspan="2">&nbsp;</td>	
		<td valign="top"><b><%=dictLanguage("PTO_Admin")%>:</b></td>
		<td valign="top"><input type="checkbox" name="permPTOAdmin" value="1" <%if permPTOAdmin then Response.write "checked"%>></td>
	</tr>								
	<tr><td colspan="4">&nbsp;</td></tr>
<%		if session("permAdminEmployees") and session("permAdminEmployeesPerms") then %>
	<tr><td colspan="4" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%		end if %>		
</table>
</form>

<% elseif strAct = "hndPerms" and strEmpID <> "" then 
		permAll					= Permissions(request("permAll"))
		permAdmin				= Permissions(request("permAdmin"))
		permAdminCalendar		= Permissions(request("permAdminCalendar"))
		permAdminDatabaseSetup	= Permissions(request("permAdminDatabaseSetup"))
		permAdminEmployees		= Permissions(request("permAdminEmployees"))
		permAdminEmployeesPerms = Permissions(request("permAdminEmployeesPerms"))
		permAdminFileRepository	= Permissions(request("permAdminFileRepository"))
		permAdminForum			= Permissions(request("permAdminForum"))
		permAdminNews			= Permissions(request("permAdminNews"))
		permAdminResources		= Permissions(request("permAdminResources"))
		permAdminSurveys		= Permissions(request("permAdminSurveys"))
		permAdminThoughts		= Permissions(request("permAdminThoughts"))
		
		permClientsAdd			= Permissions(request("permClientsAdd"))
		permClientsEdit			= Permissions(request("permClientsEdit"))
		permClientsDelete		= Permissions(request("permClientsDelete"))
		permProjectsAdd			= Permissions(request("permProjectsAdd"))
		permProjectsEdit		= Permissions(request("permProjectsEdit"))
		permProjectsDelete		= Permissions(request("permProjectsDelete"))
		permTasksAdd			= Permissions(request("permTasksAdd"))
		permTasksEdit			= Permissions(request("permTasksEdit"))
		permTasksDelete			= Permissions(request("permTasksDelete"))
		permTimecardsAdd		= Permissions(request("permTimecardsAdd"))
		permTimecardsEdit		= Permissions(request("permTimecardsEdit"))
		permTimecardsDelete		= Permissions(request("permTimecardsDelete"))
		permTimesheetsEdit		= Permissions(request("permTimesheetsEdit"))
		permForumAdd			= Permissions(request("permForumAdd"))
		permRepositoryAdd		= Permissions(request("permRepositoryAdd"))
		permPTOAdmin		    = Permissions(request("permPTOAdmin"))
		
		sql = sql_UpdatePermissions(strEmpID, permAll, permAdmin, permAdminCalendar, _
					permAdminDatabaseSetup, permAdminEmployees, permAdminEmployeesPerms, _
					permAdminFileRepository, permAdminForum, permAdminNews, permAdminResources, _
					permAdminSurveys, permAdminThoughts, permClientsAdd, permClientsEdit, _
					permClientsDelete, permProjectsAdd, permProjectsEdit, permProjectsDelete, _
					permTasksAdd, permTasksEdit, permTasksDelete, permTimecardsAdd, _
					permTimecardsEdit, permTimecardsDelete, permTimesheetsEdit, permForumAdd, _
					permRepositoryAdd, permPTOAdmin)
		'Response.write sql & "<BR>"
		Call DoSQL(sql)
		
		Session("strMessage") = dictLanguage("Permissions_Updated")		
		Response.redirect "default.asp?act=hnd_thnk" 

   elseif strAct = "delete" and strEmpID <> "" then  
		'we are not going to delete employees for now, instead - just make them inactive

   elseif strAct = "hnd_thnk" then %>
<p align="Center"><%=session("strMessage")%></p>
<%	session("strMessage") = ""%>

<%end if%>

</td>
</tr>
</table>

<p align="center"><a href=".."><%=dictLanguage("Return_Admin_Home")%></a></p>

<%
Function Permissions(strPerm)
	if strPerm <> 1 then
		strPerm = 0
	end if
	Permissions = strPerm
End Function %>



<!--#include file="../../includes/main_page_close.asp"-->
