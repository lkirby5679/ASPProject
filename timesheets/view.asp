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

<form method="POST" action="view2.asp">
<table cellpadding="2" cellspacing="2" border="0" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" align="center" class="homeheader"><%=dictLanguage("Timesheet_Details")%></td></tr>
<%	if session("errMsg") <> "" then %>
	<tr><td colspan="2">&nbsp;</td></tr>	
	<tr><td colspan="2"><font color="red"><%=session("errMsg")%></font></td></tr>	
	<tr><td colspan="2">&nbsp;</td></tr>
<%		session("errMsg") = ""
	end if %>
	<tr>
		<td class="bolddark"><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee")%>:</td>
		<td>
			<select name="employee" class="formstyleLong">
				<option value="">--<%=dictLanguage("Select")%>--</option>
<%
sql = sql_GetHourlyEmployees()
Call RunSQL(sql, rs)
while not rs.EOF
	strEmployeeName = rs("EmployeeName")
	strEmpID = rs("Employee_id") %>
	<option value="<%=strEmpID%>"><%=strEmployeeName%></option>
<%	rs.MoveNext
wend
rs.close
set rs = nothing %>
			</select>
		</td>
	</tr>
	<tr><td colspan="2" align="center"><input type="Submit" value="Submit" class="formButton"></td></tr>
</table>
</form>

<!--#include file="../includes/main_page_close.asp"-->