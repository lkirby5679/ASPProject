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

<html>
<head>
	<title><%=dictLanguage("Employee_Directory")%></title>
</head>

<!-- #include file="../includes/style.asp"-->

<%
sql = sql_GetActiveEmployees()
Call runSQL(sql, rs)
%>

<body>
<p class="homeheader"><%=gsSiteName%>&nbsp;<%=dictLanguage("Employee_Contact_Directory")%></p>

<table cellspacing="0" cellpadding="2" width="100%" border="0" bordercolor="<%=gsColorBackground%>">
	<tr>
		<td valign="top" class="small"><b><%=dictLanguage("Name")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("B-Day")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("Work")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("Mobile")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("Home")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("Address")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("City")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("State")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("Zip")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("Email")%></b></td>
		<td valign="top" class="small"><b><%=dictLanguage("IM_Name")%></b></td>
	</tr>
	
<%	while not rs.eof 
		strEmpID			= rs("employee_id")
		strEmpName			= trim(rs("EmployeeName"))
		strBDate			= rs("BDate")
		strEmailAddress		= trim(rs("EmailAddress"))
		strWorkPhone		= trim(rs("WorkPhone"))
		strWorkPhoneExt		= trim(rs("WorkPhoneExt"))
		strMobilePhone		= trim(rs("MobilePhone"))
		strHomePhone		= trim(rs("HomePhone"))
		strHomeStreet1		= trim(rs("HomeStreet1"))
		strHomeStreet2		= trim(rs("HomeStreet2"))
		strHomeCity			= trim(rs("HomeCity"))
		strHomeState		= trim(rs("HomeState"))
		strHomeZip			= trim(rs("HomeZip"))
		strIMName			= trim(rs("IMName"))
		
		if isNull(strEmpName) then
			strEmpName = "&nbsp;"
		else
			strEmpName = "<a href=""default.asp?employeeID=" & strEmpID & """ class=""small"">" & strEmpName & "</a>"
		end if
		
		if isNull(strBDate) then
			strBDate = "&nbsp;"
		else
			if isDate(strBDate) then
				'do nothing
			else
				strBDate = "&nbsp;"
			end if
		end if
				
		if isNull(strEmailAddress) then
			strEmailAddress = "&nbsp;"
		else
			strEmailAddress = "<a href=""mailto:" & strEmailAddress & """ class=""small"">" & strEmailAddress & "</a>"
		end if
		
		if isNull(strWorkPhone) then
			strWorkPhone = "&nbsp;"
		end if

		if isNull(strWorkPhoneExt) then
			'do nothing
		else
			strWorkPhoneExt = "x" & strWorkPhoneExt
			strWorkPhone = strWorkPhone & " " & strWorkPhoneExt
		end if
		
		if isNull(strMobilePhone) then
			strMobilePhone = "&nbsp;"
		end if
		
		if isNull(strHomePhone) then
			strHomePhone = "&nbsp;"
		end if
		
		if isNull(strHomeStreet1) then
			strHomeStreet1 = "&nbsp;"
		else
			if strHomeStreet2<>"" then
				strHomeStreet1 = strHomeStreet1 & "<BR>" & strHomeStreet2
			end if
		end if

		if isNull(strHomeState) then
			strHomeState = "&nbsp;"
		end if

		if isNull(strHomeZip) then
			strHomeZip = "&nbsp;"
		end if
		
		if isNull(strIMName) then
			strIMName = "&nbsp;"
		else
			strIMName = "<a href=""aim:goim?screenname=" & strIMName & """ class=""small"">" & strIMName & "</a>(<a href=""aim:addbuddy?screenname=" & strIMName & """ class=""small"">+</a>)"
		end if
		
		if bgcolor = "#EEEEEE" then 
			bgcolor="#FFFFFF"
		else 
			bgcolor="#EEEEEE"
		end if
		
		%>
		
	<tr bgcolor="<%=bgcolor%>">
		<td valign="top" class="small"><%=strEmpName%></td>
		<td valign="top" class="small"><%=strBDate%></td>
		<td valign="top" class="small"><%=strWorkPhone%></td>
		<td valign="top" class="small"><%=strMobilePhone%></td>
		<td valign="top" class="small"><%=strHomePhone%></td>
		<td valign="top" class="small"><%=strHomeStreet1%></td>
		<td valign="top" class="small"><%=strHomeCity%></td>
		<td valign="top" class="small"><%=strHomeState%></td>
		<td valign="top" class="small"><%=strHomeZip%></td>
		<td valign="top" class="small"><%=strEmailAddress%></td>
		<td valign="top" class="small"><%=strIMName%></td>
	</tr>		
<%		rs.movenext
	wend
	rs.close
	set rs = nothing %>
</table>

</body>
</html>