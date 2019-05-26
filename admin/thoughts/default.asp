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
<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<%
'*bring back records from database
strSQL = "select * from tbl_Thoughts order by id"
Call RunSQL(strSQL, rs)
%>

<table border="0" cellspacing="2" cellpadding="2" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="7" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Thoughts")%></td></tr>
	<tr>
		<td colspan="6">&nbsp;</td>
		<td align="right"><a href="processor.asp?action=add"><%=dictLanguage("Add_A_Thought")%>..</a></td>
	</tr>
<%
do while not rs.eof
	intID = rs("id")
	strThought = rs("description") %>
	<tr>
		<td valign="top" colspan="5"><%=Cut(strThought)%></td>
		<td align="center" valign="top"><a href="processor.asp?action=edit&id=<%=intID%>"><%=dictLanguage("Edit")%></a> / <a href="thought_processor.asp?action=delete&id=<%=intID%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Delete_Thought")%>')"><%=dictLanguage("Delete")%></a></td>
	</tr>
<%	rs.movenext
loop %>
</table>

<p align="center">
<a href=".."><%=dictLanguage("Return_Admin_Home")%></a>
</p>

<%
'*free up resources and close record set connections
rs.close
set rs = nothing

'*this function cuts the text string at 60 characters and adds ... to the end
function Cut(strThoughts)
	Cut = Mid(strThoughts,1,60)
	Cut = Cut & "...."
end function
%>

<!--#include file="../../includes/main_page_close.asp"-->
