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
if Session("Client_Since")="" then
	Session("Client_Since") = date()
elseif not isDate(Session("Client_Since")) then
	Session("Client_Since") = date()	
end if
%>


<!-- #include file="../includes/main_page_open.asp"-->

<font color="Red"><%=Session("strErrorMessage")%></font><% Session("strErrorMessage") = ""%>

<form method="post" action="client-add-processor.asp" name="strForm" id="strForm">
	<table border="0" cellpadding="1" cellspacing="1" align="center">
		<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%" class="homeHeader"><%=dictLanguage("Client_Details")%></td></tr>
		<tr><td colspan="2" align="right" class="alert">* <%=dictLanguage("Required_Items")%></td></tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Name")%>:</b><font class="alert">*</font></td>
			<td><input name="Name" value='<%=Session("Name")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client_Rep")%>:</b><font class="alert">*</font></td>
			<td>
				<select name="Rep" size="1" class="formstyleLong">
					<option value="" selected><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Rep")%></option>
<%
sql = sql_GetActiveEmployees()
Call RunSQL(sql, rsEmployees)
while not rsEmployees.EOF%>
 					<option value='<%=rsEmployees("Employee_ID")%>' <%if trim(session("Rep"))=trim(rsEmployees("Employee_ID")) then Response.Write "Selected"%>><%=rsEmployees("EmployeeName")%></option>
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
				<input name="Client_Since" value='<%=Session("Client_Since")%>' class="formstyleShort" onKeyPress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onClick="openCalendar('<%=server.urlencode(date())%>','Date_Change','Client_Since',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>		
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Standard_Rate")%>:</b><font class="alert">*</font></td>
			<td><input name="Standard_Rate" size="20" value='<%=Session("Standard_Rate")%>' class="formstyleShort" maxlength="8"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Address_1")%>:</b></td>
			<td><input name="Address1" size="20" value='<%=Session("Address1")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Address_2")%>:</b></td>
			<td><input name="Address2" size="20" value='<%=Session("Address2")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("City")%>, <%=dictLanguage("State")%>, <%=dictLanguage("Zip")%>:</b></td>
			<td><input name="City_State_Zip" size="20" value='<%=Session("City_State_Zip")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Web_Site_URL")%>:</b></td>
			<td><input name="LiveSite_URL" size="20" value='<%=Session("LiveSite_URL")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_Name")%>:</b></td>
			<td><input name="Contact_Name" size="20" value='<%=Session("Contact_Name")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_Phone")%>:</b></td>
			<td><input name="Contact_Phone" size="20" value='<%=Session("Contact_Phone")%>' class="formstyleMed" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_Email")%>:</b></td>
			<td><input name="Contact_Email" size="20" value='<%=Session("Contact_Email")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_2_Name")%>:</b></td>
			<td><input name="Contact2_Name" size="20" value='<%=Session("Contact2_Name")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_2_Phone")%>:</b></td>
			<td><input name="Contact2_Phone" size="20" value='<%=Session("Contact2_Phone")%>' class="formstyleMed" maxlength="100"></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Contact_2_Email")%>:</b></td>
			<td><input name="Contact2_Email" size="20" value='<%=Session("Contact2_Email")%>' class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
			<td valign="top"><b class="bolddark"><%=dictLanguage("Notes")%>:</b></td>
			<td><textarea name="Notes" class="formstyleLong" rows="4"><%=session("Notes")%></textarea></td>
		</tr>	
	</table>
	<% if session("permClientsAdd") then %>
	<p align="center"><input type="Submit" name="Submit" value="Submit" class="formbutton"></p>
	<% end if %>
</form>

<!--#include file="../includes/main_page_close.asp"-->
