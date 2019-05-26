<%@language="VBSCript"%>
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
<!--#include file="../includes/main_page_open.asp"-->

<% 
SQL = sql_GetActiveEmployees()
Call RunSQL(sql, rs)

sqlFolders = sql_GetAllFolders()
Call RunSQL(sqlFolders, rsFolders)

 %>

<table border="0" cellpadding="2" cellspacing="2" width="100%" align="center">
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" class="homeheader" align="center"><%=dictLanguage("Document_Repository")%> - <%=dictLanguage("Utilities")%></tr></tr>
	<tr>
		<td valign="top">
			<!-- #include file="../includes/fr_menu.asp" -->
			<br>
			<!-- #include file="../includes/fr_folders.asp"-->
		</td>
		<td valign="top">		
			<table border="0" cellpadding="2" cellspacing="2" width="600">	
				<tr><td colspan=2><b class="bolddark"><%=dictLanguage("Document_Search")%>:</b></td></tr>
				<tr>
					<td colspan=2><%=dictLanguage("Rep_Inst_2")%></td>
				</tr>
<%	if session("msg2") <> "" then%>
				<tr><td valign=top colspan="2" class="alert"><%=session("msg2")%></td></tr>
<%		session("msg2") = ""
	end if %>	
				<form action="search.asp" method="post" id=form1 name=form1>
				<tr>
					<td><%=dictLanguage("Contact")%>:</td>
					<td>
						<Select Name="byEmployee" size="1" class="formstyleLong">
							<option value="">--<%=dictLanguage("All")%>&nbsp;<%=dictLanguage("Contacts")%>--</option>
<%	While Not rs.EOF %>
							<option value="<%=rs("employee_ID")%>"><%=rs("employeeName")%></option>
<%		rs.MoveNext
	Wend
	rs.Close
	set rs = nothing %>		
						</Select>
					</td>
				</tr>
				<tr>
					<td nowrap><%=dictLanguage("File_Name")%>: </td>
					<td><input type="text" Name="byFile" size="20" class="formstyleLong"></td>
				</tr>
				<tr><td align=center colspan=2><input type="Submit" value="Search" class="formButton"></td></tr>
				</form>
<%if session("permRepositoryAdd") then%>
				<tr><td colspan=2><hr color="<%=gsColorBackground%>"></td></tr>
				<tr><td colspan="2"><b class="bolddark"><%=dictLanguage("Upload_Document")%>:</b></td></tr>
<%	if session("msg") <> "" then%>
				<tr><td valign=top colspan="2" class="alert"><%=session("msg")%></td></tr>
<%		session("msg") = ""
	end if %>
				<tr>
					<td colspan=2><%=dictLanguage("DocRepInst_1")%></td>
				</tr>
				<form method="post" action="upload.asp" enctype="multipart/form-data" id=form2 name=form2>
				<tr>
					<td><%=dictLanguage("Folder")%>: </td>
					<td>
						<Select Name="folder" size="1" class="formstyleLong">
							<option selected value="">--<%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Folder")%>--</option>
<%	While Not rsFolders.EOF %>
							<option value="<%=rsFolders("folder_id")%>"><%=rsFolders("folder")%></option>
<%		rsFolders.MoveNext
	Wend
	rsFolders.Close
	set rsFolders = nothing	%>		
						</Select>
					</td>
				</tr>
				<tr>
					<td><%=dictLanguage("File_Path")%>: </td>
					<td><input type="File" name="blob" class="formstyleLong"></td>
				</tr>
				<TR>
					<td>&nbsp;&nbsp;&nbsp;&nbsp;<b><%=dictLanguage("OR")%></b></td>
					<td>&nbsp;</td>
				</tr>
				<tr>
					<td><%=dictLanguage("URL")%>: </td>
					<td><input type="text" name="url" value="http://" class="formstyleLong"></td>
				</tr>				
				<tr>
					<td valign="top"><%=dictLanguage("Description")%>: </td>
					<td><textarea name="description" rows="4" class="formstyleLong"></textarea></td>
				</tr>
				<tr>
					<td align="center" colspan=2><input type="Submit" value="Submit" class="formButton"></td>
				</tr>
				<input TYPE="HIDDEN" NAME="where" value="<%=gsSiteRoot%>repository/library">
			</form>
<%end if%>			
			</table>
		</td>
	</tr>
</table>

<!--#include file="../includes/main_page_close.asp"-->
 