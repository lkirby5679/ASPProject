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

<%'load in record
cal_id = request("id")

if cal_id = "" then
	Response.Redirect "main.asp"
else
	sql = sql_GetEventsByID(cal_id)
	Call runSQL(sql, rs)
	if rs.eof then
		Response.Redirect "main.asp"	
	else
		nHeading	= rs("calendar_heading")
		nDate		= datevalue(rs("calendar_date"))
		nContent	= replace(rs("calendar_content"),vbcrlf,"<BR>")
		if rs("calendar_startTime") <> "" then
			nStartTime  = timevalue(rs("calendar_startTime"))
		end if
		if rs("calendar_endTime") <> "" then
			nEndTime	= timevalue(rs("calendar_endTime"))
		end if
	end if
	rs.close
	set rs = nothing
end if

%>

<!--#include file="../includes/popup_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border="0" width="100%" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td class="homeheader" align="Center"><%=dictLanguage("Calendar")%></td></tr> 
	<tr><td>&nbsp;</td></tr>
	<tr><td class="bolddark"><%=nHeading%></td></tr>
	<tr><td><%=nDate%>&nbsp&nbsp<%=nStartTime%> - <%=nEndTime%></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td><%=nContent%></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
</table>
<p>&nbsp;</p>

<!--#include file="../includes/popup_page_close.asp"-->
