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
thisPage = "quickAdd.asp"
mode = Request.Form("mode")
client_id = request("clientName")
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>
<head>
	<title><%=gsSiteName%></title>
	<meta name="description" content="<%=gsMetaDescription%>">
	<meta name="keywords" content="<%=gsMetaKeywords%>">
</head>

<!-- #include file="../includes/style.asp"-->
<!-- #include file="../includes/js.asp"-->

<body>

<%if mode = "" then
	sql = sql_GetAllClients() 
	Call RunSQL(sql, rsClients)%> 

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
		    <td><input type="Text" name="callTime" value="<%=time()%>" class="formstyleShort"></td>
		</tr>
		<tr>
		    <td valign="top"><b class="bolddark"><%=dictLanguage("Message")%>:</b></td>
		    <td><textarea name="message" rows="15" class="formstyleLong"></textarea></td>
		</tr>
	   <tr><td colspan="2" align="center"><input type="submit" value="Submit" class="formButton" id=submit1 name=submit1></td></tr>
	</table>
</form>

<%elseif mode = "submitMessage" then 
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Submit Message into Client Phone Log
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql = sql_InsertPhoneLog( _
		session("userid"), _
		SQLEncode(request.form("callDate")), _
		SQLEncode(request.form("callTime")), _
		client_id, _
		SQLEncode(request.form("talkedTo")), _
		SQLEncode(request.form("message")))
	'response.write(sql)
	Call DoSQL(sql)
	Response.Redirect "quickAdd.asp?mode="
end if%>

</body>
</html>

<!-- #include file="../includes/connection_close.asp" -->