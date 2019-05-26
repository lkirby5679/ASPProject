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
employeeID = trim(request("employeeID"))

if session("errorMsg") <> "" then
	response.write "<p class='alert'>" & session("errorMsg") & "</p>"
	session("errorMsg") = ""
end if

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

if employeeID<>"" then 
		sql = sql_GetEmployeesByID(employeeID)
		Call runSQL(sql, rs)
		if not rs.eof then
			strEmpName			= trim(rs("EmployeeName"))
			strEmpTitle			= trim(rs("EmployeeTitle"))
			strEmpImage			= trim(rs("image"))
			strStartDate		= rs("StartDate")
			strBDate			= rs("BDate")
			strEmailAddress		= trim(rs("EmailAddress"))
			strEmailAddressAlt	= trim(rs("EmailAddressAlt"))
			strWorkPhone		= trim(rs("WorkPhone"))
			strWorkPhoneExt		= trim(rs("WorkPhoneExt"))
			strMobilePhone		= trim(rs("MobilePhone"))
			strVoiceMail		= trim(rs("VoiceMail"))
			strHomePhone		= trim(rs("HomePhone"))
			strHomeStreet1		= trim(rs("HomeStreet1"))
			strHomeStreet2		= trim(rs("HomeStreet2"))
			strHomeCity			= trim(rs("HomeCity"))
			strHomeState		= trim(rs("HomeState"))
			strHomeZip			= trim(rs("HomeZip"))
			strIMName			= trim(rs("IMName"))
			strReportsTo		= rs("reportsTo")
			strDepartment		= rs("Department")
		end if
		rs.close
		set rs = nothing 
		
		if strEmpImage <> "" then
			strEmpImage = "<img src=""" & gsSiteRoot & "employees/images/" & strempImage & """>"
		else
			strEmpImage = "<img src=""" & gsSiteRoot & "employees/images/imageNA.gif"">"
		end if
		if strEmailAddress <> "" then
			strEmailAddress = "<a href=""mailto:" & strEmailAddress & """>" & strEmailAddress & "</a>"
		else
			strEmailAddress = "&nbsp;"
		end if
		if strEmailAddressAlt <> "" then
			strEmailAddressAlt = "<a href=""mailto:" & strEmailAddressAlt & """>" & strEmailAddressAlt & "</a>"
		else
			strEmailAddressAlt = "&nbsp;"
		end if
		if strIMName <> "" then
			strIMName = "<a href=""aim:goim?screenname=" & strIMName & """>" & strIMName & "</a>(<a href=""aim:addbuddy?screenname=" & strIMName & """>+</a>)"
		else
			strIMName = "&nbsp;"
		end if
		if strReportsTo <> "" then
			strReportsTo = "<a href=""default.asp?employeeID=" & strReportsTo & """>" & GetEmpName(strReportsTo) & "</a>"
		else
			strReportsTo = "&nbsp;"
		end if
		
		%>

<table cellpadding="2" cellspacing="2" border=0 align="center" width="400">
	<tr><td colspan="2" bgcolor="<%=gsColorHighlight%>" class="homeheader" align="center"><%=dictLanguage("Employee_Directory")%></td></tr>
	<tr>
		<td valign="top">
			<table cellpadding="2" cellspacing="2" border="0">
				<tr>
					<td class="bolddark" nowrap valign="top"><%=dictLanguage("Name")%>: </td>
					<td valign="top"><font size="+1"><%=strEmpName%></font></td>
				</tr>
				<tr>
					<td class="bolddark" nowrap valign="top"><%=dictLanguage("Title")%>: </td>
					<td valign="top"><%=strEmpTitle%></td>
				</tr>
				<tr>
					<td class="bolddark" nowrap valign="top"><%=dictLanguage("Phone")%>: </td>
					<td valign="top"><%=strWorkPhone%></td>
				</tr>
				<tr>
					<td class="bolddark" nowrap valign="top"><%=dictLanguage("Email")%>: </td>
					<td valign="top"><%=strEmailAddress%></a></td>
				</tr>												
			</table>
		</td>
		<td valign="top" align="right"><%=strEmpImage%></td>
	</tr>
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" class="homeheader"><%=dictLanguage("Business_Details")%></td></tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Department")%>: </td>
		<td valign="top"><%=strDepartment%></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Reports_To")%>: </td>
		<td valign="top"><%=strReportsTo%></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Start_Date")%>: </td>
		<td valign="top"><%=strStartDate%></td>
	</tr>	
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" class="homeheader"><%=dictLanguage("Personal_Details")%></td></tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Birth_Date")%>: </td>
		<td valign="top"><%=strBDate%></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Home_Address")%>: </td>
		<td valign="top"><%=strHomeStreet1%><br><%=strHomeCity%>,&nbsp;<%=strHomeState%>&nbsp;<%=strHomeZip%></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Home_Phone")%>: </td>
		<td valign="top"><%=strHomePhone%></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Mobile_Phone")%>: </td>
		<td valign="top"><%=strMobilePhone%></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Voice_Mail")%>: </td>
		<td valign="top"><%=strVoiceMail%></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("Alt_Email")%>: </td>
		<td valign="top"><%=strEmailAddressAlt%></a></td>
	</tr>
	<tr>
		<td class="bolddark" nowrap valign="top"><%=dictLanguage("IM_Name")%>: </td>
		<td valign="top"><%=strIMName%></td>
	</tr>
</table>
<br>
<p align="center">
<%	if trim(session("employee_id")) = trim(employeeID) then %>
<a href="javascript: popup('update.asp?act=edit&id=<%=employeeID%>');"><%=dictLanguage("Edit_My_Profile")%></a><br>
<%	end if %>
<a href="default.asp"><%=dictLanguage("Return_Employee_Directory")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>



<% else %>

<form method="POST" action="default.asp" id=strForm name=strForm>
<input type="hidden" name="mode" value="view">
<table cellpadding="2" cellspacing="2" border=0 align="center">
	<tr><td colspan="2" bgcolor="<%=gsColorHighlight%>" class="homeheader" align="center"><%=dictLanguage("Employee_Directory")%></td></tr>
	<tr>
		<td class="bolddark" nowrap><%=dictLanguage("Employee")%>: </td>
		<td>
			<select name="employeeID" class="formstyleXL">
				<option value="">--<%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee")%>--</option>
<%	sql = sql_GetActiveEmployees()
	Call RunSQL(sql, rs)
	if not rs.eof then 
		while not rs.eof %>
				<option value="<%=rs("employee_id")%>"><%=rs("employeeName")%></option>
<%			rs.movenext
		wend
	end if
	rs.close
	Set rs = nothing %>
			</select>
			<input type="submit" value="View" class="formButton" id=submit1 name=submit1>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td colspan="2"><a href="directory.asp"><b><%=dictLanguage("Click_Printable_Emp_Directory")%></b></a>&nbsp(HTML)</td></tr>
	<tr><td colspan="2"><a href="directory_download.asp?mode=X"><b><%=dictLanguage("Click_Printable_Emp_Directory")%></b></a>&nbsp(excel)</td></tr>
	<tr><td colspan="2"><a href="directory_download.asp"><b><%=dictLanguage("Click_Printable_Emp_Directory")%></b></a>&nbsp(word)</td></tr>
</table>
</form>

<br><br>
<p align="center">
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>


<% end if %>

<!--#include file="../includes/main_page_close.asp"-->