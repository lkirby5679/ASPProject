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
Sub EmployeeDropDown(Name,ID)
	response.write("<select name='" & ID & "' size='1' class='formstyleLong'>")
	response.write("<option value='' selected>" & dictLanguage("Select") & "&nbsp;" & Name & "</option>")
	if not (rsEmployees.bof and rsEmployees.eof) then
		rsEmployees.MoveFirst
		while not rsEmployees.EOF
			intEmployeeId = CINT(rsEmployees("Employee_ID"))
			if isNumeric(session(ID)) = True then
				intEmployee_ID = CINT(session(ID))
			end if
			response.write("<option  value='" & intEmployeeId & "'")
			if intEmployeeId = intEmployee_ID then
				response.write(" selected")
			end if
			response.write(">" & rsEmployees("EmployeeName") & "</option>")
			rsEmployees.MoveNext
  		wend
	end if
	response.write("</select>")
End Sub

if Session("Start_Date")="" then
	Session("Start_Date") = date()
elseif not isDate(Session("Start_Date")) then
	Session("Start_Date") = date()	
end if
if Session("Launch_Date")="" then
	Session("Launch_Date") = date()
elseif not isDate(Session("Launch_Date")) then
	Session("Launch_Date") = date()	
end if

sql = sql_GetAllClients()
Call RunSQL(sql, rsClientList)

sql = sql_GetProjectTypes()
Call RunSQL(sql, rsProjectType)

sql = sql_GetEmployeeTypes()
Call RunSQL(sql, rsRates)

sql = sql_GetActiveEmployees()
Call RunSQL(sql, rsEmployees)

if gsDynWorkOrderNumber = TRUE then 
	sql = sql_SelectLargestWorkOrderNum()
	Call RunSQL(sql, rsLargestWorkOrderNum)
	if rsLargestWorkOrderNum("maxnumber") <> NULL or rsLargestWorkOrderNum("maxnumber") <> "" then
		if session("WorkOrder_Number") = "" then
			if isNumeric(rsLargestWorkOrderNum("maxnumber")) then
				session("WorkOrder_Number") = (rsLargestWorkOrderNum("maxnumber"))+1
			end if
		end if
	end if
end if

%>

<!--#include file="../includes/main_page_open.asp"-->

<form method="post" action="project-add-processor.asp" name="strForm" id="strForm">
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr><td colspan="4" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("Project_Details")%></b></td></tr>
		<tr><td colspan="4" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Name")%>:</b><font class="alert">*</font></td>
			<td colspan="3"><input name="ProjectName" value="<%=Session("ProjectName")%>" class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client")%>:</b><font class="alert">*</font></td>
			<td colspan="3">
				<select name="Client_ID" size="1" class="formstyleLong">
					<option value="" selected><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Client")%></option>
<%
while not rsClientList.EOF
	intClientId = CINT(rsClientList("Client_ID"))
	if IsNumeric(session("Client_ID")) = True then
		intClient_ID = CINT(session("Client_ID"))
	end if %>
					<option value="<%=intClientId%>" <%if intClientId = intClient_ID then%> selected <%end if%>><%=rsClientList("Client_Name")%></option>
<%	rsClientList.MoveNext
wend%>
				</select>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Account_Rep")%>:</b><font class="alert">*</font></td>
			<td colspan="3"><%Call EmployeeDropDown(dictLanguage("Account_Rep"),"AccountExec_ID")%></td>
		</tr>	
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Work_Order")%>:</b><font class="alert">*</font></td>
			<td colspan="3"><input name="WorkOrder_Number" size="20" value="<%=Session("WorkOrder_Number")%>" class="formstyleShort" maxlength="50"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Date")%>:</b><font class="alert">*</font></td>
			<td colspan="3">
				<input name="Start_Date" size="20" value="<%=Session("Start_Date")%>" class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Start_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Date")%>:</b><font class="alert">*</font></td>
			<td colspan="3"><input name="Launch_Date" size="20" value='<%=Session("Launch_Date")%>' class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Launch_Date',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Project_Type")%>:</b><font class="alert">*</font></td>
			<td colspan="3">
				<select name="ProjectType_ID" size="1" class="formstyleMed">
					<option value="" selected><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Project_Type")%></option>
<%
while not rsProjectType.EOF
	intProjectTypeId = CINT(rsProjectType("ProjectType_ID"))
	if isNumeric(session("ProjectType_ID")) = True then
		intProjectType_ID = CINT(session("ProjectType_ID"))
	end if %>			
		 	<option value="<%=intProjectTypeId%>" <%if intProjectTypeId = intProjectType_ID then%> selected <%end if%>><%=rsProjectType("ProjectTypeDescription")%></option>		
<%	rsProjectType.Movenext
wend%>
				</select>
			</td>
		</tr>
		<tr>
			<td valign="top"><b class="bolddark"><%=dictLanguage("Status")%>:</b></td>
			<td colspan="3">
				<input name="color" type="radio" value="Green"<%if session("color")="Green" then%> Checked<%end if%>> <%=dictLanguage("Green")%><br>
				<input name="color" type="radio" value="Blue"<%if session("color")="Blue" then%> Checked<%end if%>> <%=dictLanguage("Blue")%><br>
				<input name="color" type="radio" value="Red"<%if session("color")="Red" then%> Checked<%end if%>> <%=dictLanguage("Red")%>
			</td>
	    </tr>
	    <tr>
			<td valign="top"><b class="bolddark"><%=dictLanguage("Comments")%>:</b></td>
			<td colspan="3"><textarea name="comments" rows="4" class="formstyleLong"><%=session("comments")%></textarea></td>
		</tr> 	 
		<tr>
			<td>&nbsp;</td>
			<td align="center"><%=dictLanguage("Rate")%></td>
			<td align="center"><%=dictLanguage("Hours")%></td>
			<td><%=dictLanguage("Core_Team")%></td>
		</tr>
<% while not rsRates.eof %>
		<tr>
			<td><b class="bolddark"><%=rsRates("EmployeeType")%></b></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xRate" class="formstyleTiny" size="4" value="<%if session(rsRates("EmployeeType_ID") & "xRate")<>"" then%><%=session(rsRates("EmployeeType_ID") & "xRate")%><%else%><%=rsRates("Rate")%><%end if%>"></td>
			<td><input type="text" name="<%=rsRates("EmployeeType_ID")%>xHours" class="formstyleTiny" size="4" value="<%if session(rsRates("EmployeeType_ID") & "xHours")<>"" then%><%=session(rsRates("EmployeeType_ID") & "xHours")%><%else response.write "0"%><%end if%>"></td>
			<td><%Call EmployeeDropDown(dictLanguage("Team_Member"),rsRates("EmployeeType_ID") & "xEmployee_ID")%></td>
		</tr>
<% 		rsRates.movenext 
	wend
%>
	</table>

	<% if session("permProjectsAdd") then%>
	<p align="center">
		<input type="Submit" name="Submit" value="Submit" class="formButton">
		<input type="Reset" name="Reset" value="Reset" class="formButton">
	</p>
	<% end if %>
</form>

<%
rsEmployees.Close
set rsEmployees=nothing
rsClientList.Close
set rsClientList=nothing
rsProjectType.Close
set rsProjectType=nothing
rsRates.close
set rsRates = nothing
%>

<br>

<!--#include file="../includes/main_page_close.asp"-->