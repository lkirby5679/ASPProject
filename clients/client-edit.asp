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
client_id = request.querystring("client_id")
sql = sql_GetClientsByID(client_id)
Call RunSQL(sql, rsClient)
if not rsClient.eof then
	boolFound = TRUE
	cName			= trim(rsClient("client_name"))
	cRepID			= rsClient("rep_ID")
	cClientSince	= rsClient("Client_Since")
	if not isDate(cClientSince) then
		cClientSince = date()
	end if
	cStandardRate	= rsClient("Standard_Rate")
	cAddress1		= trim(rsClient("Address1"))
	cAddress2		= trim(rsClient("Address2"))
	cCSZ			= trim(rsClient("City_State_Zip"))
	cLiveSiteURL	= trim(rsClient("LiveSite_URL"))
	cContactName	= trim(rsClient("Contact_Name"))
	cContactPhone	= trim(rsClient("Contact_Phone"))
	cContactEmail	= trim(rsClient("Contact_Email"))
	cContact2Name	= trim(rsClient("Contact2_Name"))
	cContact2Phone	= trim(rsClient("Contact2_Phone"))
	cContact2Email	= trim(rsClient("Contact2_Email"))
	cActive			= rsClient("Active")
	if cActive = 1 or cActive=-1 or cActive then
		cActive = TRUE
	else 
		cActive = false
	end if
else
	boolFound = FALSE	
end if
rsClient.close
set rsClient = nothing
%>

<!-- #include file="../includes/main_page_open.asp"-->

<% if boolFound then %>
<form method="post" action="client-edit-processor.asp" name="strForm" value="strForm">
	<input type="hidden" name="client_id" value="<%=client_id%>">
	<table border="0" cellpadding="1" cellspacing="1" align="center">
		<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%" class="header"><%=dictLanguage("Client_Details")%></td></tr>
		<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
	    <tr>
			<td><b class="bolddark"><%=dictLanguage("Name")%>:</b><font class="alert">*</font></td>
			<td><input name="Name" size="50" value="<%=cName%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client_Rep")%>:</b><font class="alert">*</font></td>
		    <td>
				<select name="Rep" size="1" class="formStyleLong">
					<option value=""><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Client_Rep")%></option>
<%
sql = sql_GetActiveEmployees()
Call RunSQL(sql, rsEmployees)
while not rsEmployees.EOF%>
 					<option value="<%=rsEmployees("Employee_ID")%>" <%if cRepID = rsEmployees("Employee_ID") then Response.Write "Selected"%>><%=rsEmployees("EmployeeName")%></option>
<%	rsEmployees.MoveNext
wend
rsEmployees.close
set rsEmployees = nothing %>
				</select>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client_Since")%>:</b><font class="alert">*</font></td>
			<td>
				<input name="Client_Since" size="20" value="<%=cClientSince%>" class="formStyleShort" onKeyPress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(cClientSince)%>','Date_Change','Client_Since',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Standard_Rate")%>:</b><font class="alert">*</font></td>
			<td><input name="Standard_Rate" size="20" value="<%=cStandardRate%>" class="formStyleShort" maxlength="8"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Address_1")%>:</b></td>
			<td><input name="Address1" size="20" value="<%=cAddress1%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Address_2")%>:</b></td>
			<td><input name="Address2" size="20" value="<%=cAddress2%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("City")%>, <%=dictLanguage("State")%>, <%=dictLanguage("Zip")%>:</b></td>
			<td><input name="City_State_Zip" size="20" value="<%=cCSZ%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Web_Site_URL")%>:</b></td>
			<td><input name="LiveSite_URL" size="20" value="<%=cLiveSiteURL%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_Name")%>:</b></td>
			<td><input name="Contact_Name" size="20" value="<%=cContactName%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_Phone")%>:</b></td>
			<td><input name="Contact_Phone" size="20" value="<%=cContactPhone%>" class="formStyleMed" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_Email")%>:</b></td>
			<td><input name="Contact_Email" size="20" value="<%=cContactEmail%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_2_Name")%>:</b></td>
			<td><input name="Contact2_Name" size="20" value="<%=cContact2Name%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_2_Phone")%>:</b></td>
			<td><input name="Contact2_Phone" size="20" value="<%=cContact2Phone%>" class="formStyleMed" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_2_Email")%>:</b></td>
			<td><input name="Contact2_Email" size="20" value="<%=cContact2Email%>" class="formStyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Active")%>:</b></td>
			<td>
				<select name="Active" size="1" class="formStyleShort">
					<option value=False>False</option>
					<option value=True <%if cActive then Response.Write "Selected"%>>True</option>
				</select>
			</td>
		</tr>
	</table>
	<%if session("permClientsEdit") then %>
   <p align="center"><input type="Submit" name="Submit" value="Submit" class="formButton"></p>
	<%end if%>
</form>

<% else %>
	<div align="center"><%=dictLanguage("No_Client_ID")%>.</div>
<% end if %>

<p align="center">
<a href="client-view.asp"><%=dictLanguage("View_Clients")%></a><br>
<a href="client-add.asp"><%=dictLanguage("Add_Client")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<!-- #include file="../includes/main_page_close.asp"-->