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


<table border="0" cellspacing="2" cellpadding="2" align="center" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" align="Center" class="homeheader"><%=dictLanguage("Admin")%></td></tr>
	<tr>
		<td valign="top" width="200">
	
	<table border="0" cellpadding="2" cellpadding="2">	
<%if session("permAdminDatabaseSetup") then %>				
	<tr><td><a href="config/default.asp"><%=dictLanguage("Site_Configuration")%></a></td></tr>
<%end if%>				
<%if session("permAdminCalendar") then %>
	<tr><td><a href="calendar/default.asp"><%=dictLanguage("Calendar")%></a></td></tr>
<%end if
  if session("permAdminDatabaseSetup") then %>
	<!--tr><td><a href="dbsetup.asp" target="_blank"><%=dictLanguage("Database_Setup")%></a></td></tr-->
<%end if
  if session("permAdminForum") then %>
	<tr><td><a href="forum/admin.asp"><%=dictLanguage("Discussion_Forum")%></a></td></tr>
<%end if
  if session("permAdminFileRepository") then %>
	<tr><td><a href="repository/default.asp"><%=dictLanguage("Document_Repository")%></a></td></tr>
<%end if
  if session("permAdminEmployees") then %>
	<tr><td><a href="employees/default.asp"><%=dictLanguage("Employees")%></a></td></tr>
	<tr><td><a href="oddsnends/employeetypes.asp"><%=dictLanguage("Employee_Types")%></a></td></tr>
	<tr><td><a href="oddsnends/payperiods.asp"><%=dictLanguage("Employee_Pay_Periods")%></a></td></tr>
	<tr><td><a href="oddsnends/timecardtypes.asp"><%=dictLanguage("Timecard_Types")%></a></td></tr>
<%end if
  if session("permAdminNews") then %>
	<tr><td><a href="news/default.asp"><%=dictLanguage("News")%></a></td></tr>
<%end if%>
<%if session("permAdminEmployees") then %>
	<tr><td><a href="oddsnends/projectphases.asp"><%=dictLanguage("Project_Phases")%></a></td></tr>
<%end if%>
<%if session("permAdminResources") then%>
	<tr><td><a href="resources/default.asp"><%=dictLanguage("Resources")%></a></td></tr>
<%end if%>
<%if session("permAdminSurveys") then%>
	<tr><td><a href="survey/default.asp"><%=dictLanguage("Surveys")%></a></td></tr>
<%end if 
  if session("permAdminThoughts") then%>
	<tr><td><a href="thoughts/default.asp"><%=dictLanguage("Thoughts")%></a></td></tr>
<%end if %>
	<tr><td>&nbsp;</td></tr>
	</table>

		</td>
		<td valign="top">
			
<%if session("permAdminDatabaseSetup") then%>
	<table cellpadding="2" cellspacing="2" border="0" align="center">
		<tr><td colspan="4"><%=dictLanguage("Database_Export_Utils")%></td></tr>
		<tr>
			<td><%=dictLanguage("Employees")%>:</td>
			<td><a href="exports/employees.asp">[<%=dictLanguage("Table_View")%>]</a></td>
			<td><a href="exports/employees.asp?strType=excel">[<%=dictLanguage("Excel")%>]</a></td>
			<td><a href="exports/employees.asp?strType=cdtf">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>
		<tr>
			<td><%=dictLanguage("Clients")%>:</td>
			<td><a href="exports/clients.asp">[<%=dictLanguage("Table_View")%>]</a></td>
			<td><a href="exports/clients.asp?strType=excel">[<%=dictLanguage("Excel")%>]</a></td>
			<td><a href="exports/clients.asp?strType=cdtf">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>
		<tr>
			<td><%=dictLanguage("Projects")%>:</td>
			<td><a href="exports/projects.asp">[<%=dictLanguage("Table_View")%>]</a></td>
			<td><a href="exports/projects.asp?strType=excel">[<%=dictLanguage("Excel")%>]</a></td>
			<td><a href="exports/projects.asp?strType=cdtf">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>	
		<tr>
			<td><%=dictLanguage("Tasks")%>:</td>
			<td><a href="exports/tasks.asp">[<%=dictLanguage("Table_View")%>]</a></td>
			<td><a href="exports/tasks.asp?strType=excel">[<%=dictLanguage("Excel")%>]</a></td>
			<td><a href="exports/tasks.asp?strType=cdtf">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>	
		<tr>
			<td><%=dictLanguage("Timecards")%>:</td>
			<td><a href="exports/timecards.asp">[<%=dictLanguage("Table_View")%>]</a></td>
			<td><a href="exports/timecards.asp?strType=excel">[<%=dictLanguage("Excel")%>]</a></td>
			<td><a href="exports/timecards.asp?strType=cdtf">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>
		<tr>
			<td><%=dictLanguage("Timesheets")%>:</td>
			<td><a href="exports/timesheets.asp">[<%=dictLanguage("Table_View")%>]</a></td>
			<td><a href="exports/timesheets.asp?strType=excel">[<%=dictLanguage("Excel")%>]</a></td>
			<td><a href="exports/timesheets.asp?strType=cdtf">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>														
		<tr><td colspan="4">&nbsp;</td></tr>
		<form name="strForm" id="strForm" method="POST" action="exports/variable.asp">
		<input type="hidden" name="strType" value="">
		<tr>
			<td valign="top"><select name="strTable" size="1" class="formstyleMed">
<%	Set rs = conn.OpenSchema(20, Array(Empty, Empty, Empty, "TABLE"))
	Do While Not rs.EOF
		strNewTable = rs("TABLE_NAME").Value
		If InStr(strNewTable, "MSys") = 0 Then
			Response.Write("<option value='" & strNewTable & "' selected>" & strNewTable & "</option>")
			rs.MoveNext
		End If
	Loop
	rs.Close
	Set rs = Nothing %>
				</select>
			</td>
			<td valign="top"><a href="javascript: document.all.strForm.strType.value=''; document.all.strForm.submit();">[<%=dictLanguage("Table_View")%>]</a></td>
			<td valign="top"><a href="javascript: document.all.strForm.strType.value='excel'; document.all.strForm.submit();">[<%=dictLanguage("Excel")%>]</a></td>
			<td valign="top"><a href="javascript: document.all.strForm.strType.value='cdtf'; document.all.strForm.submit();">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>
		</form>
		<tr><td colspan="4">&nbsp;</td></tr>
		<form name="strForm2" id="strForm2" method="POST" action="exports/variable.asp">
		<input type="hidden" name="strType" value="">
		<tr>
			<td valign="top"><%=dictLanguage("SQL_Query")%>:<br><textarea name="sqlquery" rows="5" cols="20" class="formstylemed">SELECT <%=VBCRLF%>FROM <%=VBCRLF%>WHERE <%=VBCRLF%>ORDER BY </textarea><br>
			<td><a href="javascript: document.all.strForm2.strType.value=''; document.all.strForm2.submit();">[<%=dictLanguage("Table_View")%>]</a></td>
			<td><a href="javascript: document.all.strForm2.strType.value='excel'; document.all.strForm2.submit();">[<%=dictLanguage("Excel")%>]</a></td>
			<td><a href="javascript: document.all.strForm2.strType.value='cdtf'; document.all.strForm2.submit();">[<%=dictLanguage("CDTF")%>]</a></td>
		</tr>
		</form>
	</table>
<%end if%>

		</td>
	</tr>
</table>

<!--#include file="../includes/main_page_close.asp"-->
