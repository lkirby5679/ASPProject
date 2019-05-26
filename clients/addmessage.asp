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
mode = Request("mode")
thisPage = "addMessage.asp"
client_id = request("client")
%>

<!-- #include file="../includes/main_page_open.asp"-->

<%if mode = "" then 
	sqlClients = sql_GetClientsByID(client_id)
	Call RunSQL(sqlClients, rsClients)
	if not rsClients.eof then
		strClientName = rsClients("Client_Name")
	end if
	rsClients.close
	set rsClients = nothing %>

<form name="strForm" id="strForm" method="post" action="<%=thisPage & "?client=" & client_id & ""%>">
	<input type="hidden" name="mode" value="submitMessage">
	<table border="0" cellpadding="1" cellspacing="1" align="center">
		<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" width="100%" class="header"><%=dictLanguage("Client_Contact_Log")%></td></tr>
	    <tr> 
		    <td><b class="bolddark"><%=dictLanguage("Rep")%>:</b></td>
			<td><b><%=session("userid")%></b></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client")%>:</b></td>
			<td><b><%=strClientName%></b></td>
		</tr>
		<tr>
		    <td><b class="bolddark"><%=dictLanguage("Contact")%>: (<%=dictLanguage("spoke_to")%>)</b></td>
		    <td><input type="Text" name="talkedTo" class="formstyleLong" maxlength="100"></td>
		</tr>
		<tr>
		    <td><b class="bolddark"><%=dictLanguage("Date_Of_Call")%>:</b></td>
		    <td>
				<input type="Text" name="callDate" value="<%=date()%>" class="formstyleShort" onkeypress="txtDate_onKeypress();" maxlength="10">
				<a href="javascript:doNothing()" onclick="openCalendar('<%=server.urlencode(cClientSince)%>','Date_Change','callDate',150,300)"><img border="0" src="<%=gsSiteRoot%>images/calendaricon.jpg" onmouseover="this.style.cursor='hand'" WIDTH="16" HEIGHT="15"></a>
			</td>
		</tr>
		<tr>
		    <td><b class="bolddark"><%=dictLanguage("Time_Of_Call")%>:</b></td>
		    <td><input type="Text" name="callTime" size="49" value="<%=time()%>" class="formstyleShort"></td>
		</tr>
		<tr>
		    <td valign="top"><b class="bolddark"><%=dictLanguage("Message")%>:</b></td>
		    <td><textarea name="message" rows="15" class="formstyleLong"></textarea></td>
		</tr>
	   <tr><td colspan="2" align="center"><input type="submit" value="Submit" class="formButton"></td></tr>
	</table>
</form>
<%elseif mode = "submitMessage" then 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Submit Message into Client Phone Log
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql = sql_InsertPhoneLog( _
		session("userid"), _
		SQLEncode(Request.Form("CallDate")), _
		SQLEncode(Request.Form("calltime")), _
		client_id, _
		SQLEncode(Request.Form("talkedTo")), _
		SQLEncode(Request.Form("message"))) 
	'response.write(sql)
	Call DoSQL(sql)

	Response.Redirect "view_log.asp?mode=viewLog&clientName=" & client_id & "" %>
<%end if%>

<!-- #include file="../includes/main_page_close.asp"-->
