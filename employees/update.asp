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


<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/popup_page_open.asp"-->


<table border="0" cellspacing="2" cellpadding="2" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td align="Center" class="homeheader"><%=dictLanguage("Edit_My_Profile")%></td>
	</tr>
</table>

<%if strAct = "edit" and strEmpID <> "" then 
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
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_ID")%>:</b></td>
		<td valign="top"><%=empID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Name")%>:</b></td>
		<td valign="top"><%=empName%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Login")%>:</b>&nbsp;(<%=dictLanguage("Initials")%>)</td>
		<td valign="top"><%=empLogin%></td>
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
			ext.<input type="text" name="empWorkPhoneExt" value="<%=empWorkPhoneExt%>" class="formstyleTiny" maxLength="50">
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
<%		if trim(session("employee_id")) = trim(strEmpID) then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%		end if %>			
</table>
</form>

<%  elseif strAct = "hnd_thnk" then %>
	<script language="JavaScript">
		opener.location.reload();
		window.close();
	</script>
<p align="Center"><%=session("strMessage")%></p>
<%	session("strMessage") = ""%>

<%end if%>


<!--#include file="../includes/popup_page_close.asp"-->

