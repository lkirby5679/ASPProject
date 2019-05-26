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
id = Request("id")

'*determine whether we are adding to ,updating, or deleting a record from the database table.
if request.querystring("action") = "add" then
	'create database connection
	sql = "insert into tbl_Thoughts (Description, shows) values  ('" & SQLencode(request.form("thoughts")) & "',0);"
elseif request.querystring("action") = "edit" then
	sql = "update tbl_Thoughts set description = '" & SQLencode(request.form("description")) & "' where id = " & id & ""
	'response.write(sql)
elseif request.querystring("action") = "delete" then
	sql = "delete from tbl_Thoughts where id = " & id & ";"
end if

Call DoSQL(sql)

session("id") = ""
%>
	
<table border="0" cellspacing="2" cellpadding="2" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Thoughts")%></td></tr>
	<tr><td align="Center"><%=dictLanguage("Thought_Thanks")%></td></tr>
</table>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Thoughts_Admin_Home")%></a><br>
<a href=""><%=dictLanguage("Return_Admin_Home")%></a>
</p>

<!--#include file="../../includes/main_page_close.asp"-->
