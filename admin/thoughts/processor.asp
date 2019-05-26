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


<% if request.querystring("action") = "add" then %>

<form action="thought_processor.asp?action=add" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Add_A_Thought")%></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td valign="top"><b class="bolddark"><%=dictLanguage("Thought")%>:</b></td>
		<td valign="top">
			<textarea name="thoughts" rows="12" class="formstyleXL"><%=session("thoughts")%></textarea>
		</td>
	</tr>
	<tr><td colspan="2" align="center">	<input type="submit" value="Submit" id=submit2 name=submit2 class="formButton"></td></tr>
</table>
</form>
		
<% elseif request.querystring("action") = "edit" then
	'*open database connection
	sql = "select * from tbl_thoughts where id =" &  request.querystring("id") & ";"
	Call RunSQL(sql, rs)

	'create session variable for the id 
	id = request.querystring("id") %>

<form action="thought_processor.asp?action=edit" method="post">
<input type="hidden" name="id" value="<%=id%>">
<table border="0" cellspacing="2" cellpadding="2" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Edit_Thought")%></td></tr>
	<tr><td>&nbsp;</td></tr>
	<tr>
		<td valign="top"><b class="bolddark"><%=dictLanguage("Thought")%>:</b></td>
		<td valign="top">
			<textarea name="description" rows="12" class="formstyleXL"><%=rs("description")%></textarea>
		</td>
	</tr>
	<tr><td colspan="2" align="center"><input type="submit" value="Submit" id=submit1 name=submit1 class="formButton"></td></tr>
</table>
</form>
	
<% end if%>

<!--#include file="../../includes/main_page_close.asp"-->
