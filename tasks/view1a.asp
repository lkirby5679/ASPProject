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

<% active = request("active")
   if active <> 0 then
		active = 1
   end if %>

<!--#include file="../includes/main_page_open.asp"-->


<table cellpadding="2" cellspacing="2" border=0 align="center">
	<tr>
		<td>
			<form method="POST" action="view2.asp">
			<input type="hidden" name="active" value="<%=active%>"> 
				<table border="0" cellpadding="2" cellspacing="2" align="Center">
					<tr bgcolor="<%=gsColorHighlight%>"><td align="center" colspan="2" class="homeheader"><%=dictLanguage("Tasks_For_Client")%></td></tr>
					<tr><td colspan="2">&nbsp;</td></tr>
					<tr>
						<td><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Client")%>:</td>
						<td>
							<select name="Client_ID" class="formstyleXL">
<%
sql = sql_GetTaskListAndCount(active)
Call RunSQL(sql, rs)
while not rs.eof
	if rs("Client_ID")<>curid then
		response.write "<option value='" & rs("Client_ID") & "~" & rs("Client_Name") & "'>" & rs("Client_Name") & " (" & rs("CountOfid") & " " & dictLanguage("tasks") & ")"
	end if
	curid=rs("Client_ID")
	rs.movenext
wend
rs.close
set rs = nothing
%>
							</select>
							<input type="submit" value="Submit" class="formButton">
						</td>
					</tr>
				</table>
			</form>
			<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>
		</td>
	</tr>
</table>

<!--#include file="../includes/main_page_close.asp"-->
