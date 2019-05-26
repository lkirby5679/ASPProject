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
	<tr bgcolor="<%=gsColorHighlight%>"><td class="homeheader" colspan="2" align="center"><%=dictLanguage("Resources")%></td></tr> 
	<tr><td colspan="2">&nbsp;</td></tr>

<%	sql = sql_GetResources()
	Call runSQL(sql, rs)
	if not rs.eof then 			
		while not rs.eof %>
	<tr>
		<td	colspan="2"><a href="<%=trim(rs("resources_url"))%>" target="_blank" class="bolddark"><%=rs("resources_title")%></a></td>
	</tr>
<%			if rs("resources_caption")<>"" then %>							
	<tr>
		<td>&nbsp;</td>
		<td><font class="small"><%=rs("resources_caption")%></small></td>
	</tr>
<%			end if 
			rs.movenext 
		wend 
	rs.close
	set rs = nothing 
	end if %>

</table>
<br>

<!--#include file="../includes/popup_page_close.asp"-->

