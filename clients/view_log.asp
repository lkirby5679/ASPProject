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
thisPage = "view_log.asp"
mode = Request("mode")
client_id = request("clientName")
%>

<!--#include file="../includes/main_page_open.asp"-->

<%if mode = "" then 
	sqlClients = "SELECT * FROM tbl_Clients ORDER BY Client_Name"
	Call runSQL(sqlClients, rsClients)%> 

<form name="logForm" method="post" action="<%=thisPage%>">
	<input type="hidden" name="mode" value="viewLog">
	<table border="0" cellpadding="1" cellspacing="1" align="center">
		<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%" class="header"><%=dictLanguage("Client_Contact_Log")%></td></tr>
	    <tr>
			<td><b class="bolddark"><%=dictLanguage("Select_Client_Log")%>:</b></td>
			<td>
				<select name="clientName" class="formstyleLong">
<%	firstTime = true
	do while not rsClients.eof
		vOptions = vOptions & "<option value=" & rsClients("Client_ID") 
		if firstTime=true then
			vOptions = vOptions & " selected"
			firstTime = false
		end if
		vOptions = vOptions & ">" & rsClients("Client_Name") & "</option>"
		rsClients.movenext
	loop
	Response.Write vOptions 
	rsClients.close
	set rsClients = nothing %>
				</select>
			</td>
		</tr>
		<tr><td colspan="2" align="center"><input type="submit" value="Submit" class="formButton"></td>
	</table>
</form>
	
	
<%elseif mode = "viewLog" then 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Display Log
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sqlClients =	"SELECT * FROM tbl_Clients right join tbl_employees on " & _
					"tbl_clients.rep_id = tbl_employees.employee_id " & _
					"WHERE Client_ID = " & client_id  & ""
	Call RunSQL(sqlClients, rsClients)
	%>
<table border="0" cellpadding="1" cellspacing="1" align="center">
	<tr><td colspan="5" align="center" bgcolor="<%=gsColorHighlight%>" width="100%" class="header"><%=dictLanguage("Client_Contact_Log")%></td></tr>
    <tr>
		<td colspan="5"><a href="addMessage.asp?client=<%=client_id%>"><%=dictLanguage("Add_Message")%></a></td>
	</tr>
	<tr>
		<td><b class="bolddark" nowrap><%=dictLanguage("Client_Name")%>:</b></td>
		<td colspan="4"><%=rsClients("Client_Name")%></td>
	</tr>
<%	if rsClients("Contact_Name") <> "" then %>
	<tr>
		<td><b class="bolddark" nowrap><%=dictLanguage("Contact")%>:</b></td>
		<td colspan="4"><%=rsClients("Contact_Name")%></td>
	</tr>
<%	end if%>
<%	if rsClients("Contact_Phone") <> "" then %>
	<tr>
		<td><b class="bolddark" nowrap><%=dictLanguage("Phone")%>:</b></td>
		<td colspan="4"><%=rsClients("Contact_Phone")%></td>
	</tr>
<%	end if%>
<%	if rsClients("LiveSite_URL") <> "" then %>
	<tr>
		<td><b class="bolddark" nowrap><%=dictLanguage("URL")%>:</b></td>
		<td colspan="4"><a href="http://<%=rsClients("LiveSite_URL")%>" target="_blank"><%=rsClients("LiveSite_URL")%></a></td>
	</tr>
<%	end if%>
<%	if rsClients("EmployeeName")<>"" then %>
	<tr>
		<td><b class="bolddark" nowrap>Account Rep:</b></td>
		<td colspan="4"><%=rsClients("EmployeeName")%></td>
	</tr>
<%	end if %>	
	<tr>
	    <td bgcolor="<%=gsColorHighlight%>" align=center class="columnheader" nowrap><%=dictLanguage("Date")%></td>
	    <td bgcolor="<%=gsColorHighlight%>" align=center class="columnheader" nowrap><%=dictLanguage("Time")%></td>
	    <td bgcolor="<%=gsColorHighlight%>" align=center class="columnheader" nowrap><%=dictLanguage("Rep")%></td>
	    <td bgcolor="<%=gsColorHighlight%>" align=center class="columnheader" nowrap><%=dictLanguage("Spoke_To")%></td>
	    <td bgcolor="<%=gsColorHighlight%>" align=center class="columnheader" width="300"><%=dictLanguage("Message")%></td>
	</tr>
<% 

	sql = "select * FROM tbl_PhoneLog WHERE Client_ID = " & client_id & " ORDER BY callDate desc, callTime DESC"
	Call RunSQL(sql, rsLog)
	if not rsLog.eof then
		intRowcounter = 0
		do while not rsLog.eof
			intRowcounter = intRowcounter + 1 
			If intRowcounter MOD 2 = 1 then 
				strRowColor = "#ffffff"
			Else
				strRowColor = gsColorWhite
			End If
%>
	<tr bgcolor="<%=strRowColor%>">
		<td valign="top" nowrap align=center class="small"><%=datevalue(rsLog("callDate"))%></td>
		<td valign="top" nowrap align=center class="small"><%=timevalue(rsLog("callTime"))%></td>
		<td valign="top" nowrap align=center class="small"><%=rsLog("sbEmp")%></td>
		<td valign="top" nowrap align=center class="small"><%=rsLog("talkedTo")%></td>
		<td valign="top" class="small" width="300"><%=rsLog("message")%></td>
	</tr>
<%
			rsLog.movenext
		loop
	end if
	rsLog.close
	set rsLog = nothing
	rsClients.close
	set rsClients = nothing
%>
</table>
<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>
	
<%end if%>

<!--#include file="../includes/main_page_close.asp"-->
