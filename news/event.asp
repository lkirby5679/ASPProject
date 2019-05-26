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
news_id = request("id")
if news_id = "" then
	Response.Redirect gsSiteRoot & "news/"
else
	sql = sql_GetNewsByID(news_id)
	Call runSQL(sql, rs)	
	if rs.eof then
		Response.Redirect gsSiteRoot & "news/"	
	else
		nHeading	= rs("news_heading")
		nDate		= rs("news_date")
		nContent	= replace(rs("news_content"),vbcrlf,"<BR>")
	end if
	rs.close
	set rs = nothing
end if

%>

<!--#include file="../includes/popup_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border="0" width="100%" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td class="homeheader" align="center"><%=dictLanguage("News")%></td></tr> 
	<tr><td>&nbsp;</td></tr>
	<tr><td class="bolddark"><%=nHeading%></td></tr>
	<tr><td><%=nDate%></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr><td><%=nContent%></td></tr>
	<tr><td colspan="2">&nbsp;</td></tr>
</table>
<p>&nbsp;</p>

<!--#include file="../includes/popup_page_close.asp"-->
