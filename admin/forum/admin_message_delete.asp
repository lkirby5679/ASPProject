<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->
<!--#include file="../../includes/forum_common.asp"-->
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
<table border="0" width="500" cellpadding="3" cellspacing="3" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader">Discussion Forum Admin - Message Deleted</td></tr>
	<tr>
		<td colspan="2" >  								

<%
Dim strSQL, iActiveMessageId
iActiveMessageId = cInt(Request.Querystring("mid"))
DeleteMessage iActiveMessageId
%>
The message has been deleted.<br>
Back to <a href='admin_forum_display.asp'>Forum</a>



							
							
	</TD>		
	<TR>		
</table>
<!--#include file="../../includes/main_page_close.asp"-->