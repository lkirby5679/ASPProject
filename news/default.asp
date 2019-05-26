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
<!--#include file="../includes/popup_page_open.asp"-->

<table cellpadding="2" cellspacing="2" border="0" width="100%" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td class="homeheader" colspan="2" align="center"><%=dictLanguage("News")%></td></tr> 
	<tr><td colspan="2">&nbsp;</td></tr>

<%	sql = sql_GetActiveNews()
	Call runSQL(sql, rs)
	if not rs.eof then 
		while not rs.eof %>

	<tr>
		<td nowrap>
<%			if trim(rs("news_url")) <> "" then %>
				<a href="<%=trim(rs("news_url"))%>" target="_blank" class="bolddark"><%=rs("news_date")%></a>&nbsp;&nbsp;
<%			else %>
				<a href="javascript: popup2('event.asp?id=<%=rs("id")%>');" class="bolddark"><%=rs("news_date")%></a>&nbsp;&nbsp;
<%			end if %>	
		</td>
		<td width="100%">
<%			if trim(rs("news_url"))<>"" then %>
			<a href="<%=trim(rs("news_url"))%>" target="_blank" class="bolddark"><%=rs("news_heading")%></a> (<%=rs("news_source")%>)
<%			else %>
			<a href="javascript: popup2('event.asp?id=<%=rs("id")%>');" class="bolddark"><%=rs("news_heading")%></a>
<%			end if %>
		</td>
	</tr>
<%			if trim(rs("news_url"))<>"" then
				'do nothing
			else %>
	<tr>
		<td>&nbsp;</td>
		<td><%=rs("news_abstract")%>&nbsp;<a href="javascript: popup2('event.asp?id=<%=rs("id")%>');" class="small"><%=dictLanguage("more")%> >></a></td>
	</tr>
<%			end if %>
	<tr><td colspan="2">&nbsp;</td></tr>
<%			rs.movenext 
		wend
	end if
	rs.close
	set rs = nothing %>

</table>

<!--#include file="../includes/popup_page_close.asp"-->
