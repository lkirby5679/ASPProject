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
		
<%	
sqlSearch = "SELECT * FROM tbl_documents"
if request.form("byFile") <> "" then
	fileNamesql = " (fileName LIKE '%" & SQLEncode(request.form("byFile")) & "%' " & _
		"or description LIKE '%" & SQLEncode(Request.Form("byFile")) & "%') "
	fileNameFlag = 1
end if
if request.form("byEmployee") <> "" then
	contactsql = "submitBy = " & request.form("byEmployee") & " "
	contactFlag = 1
end if
if fileNameFlag = 1 or contactFlag = 1 then
	sqlSearch = sqlSearch & " WHERE "
	if fileNameFlag = 1 then
		sqlSearch = sqlSearch & fileNamesql
		if contactFlag = 1 then
			sqlSearch = sqlSearch & " AND "
		end if
	end if
	if contactFlag = 1 then
		sqlSearch = sqlSearch & contactsql
	end if
end if
sqlSearch = sqlSearch & " ORDER BY submitDate DESC, fileName"
'response.write(sqlSearch)
Call RunSQL(sqlSearch, rsSearchDocs)
if rsSearchDocs.eof then
	session("msg2")="There were no records matching that search criteria."
	Response.Redirect "utilities.asp"
else
%>

<!--#include file="../includes/main_page_open.asp"-->

<table border="0" cellpadding="2" cellspacing="2" width="100%" align="center">
<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" align="center" class="homeheader"><%=dictLanguage("Document_Repository")%>: <%=dictLanguage("Search_Results")%></td></tr> 
	<tr>
		<td valign="top">
			<!-- #include file="../includes/fr_menu.asp" -->
			<br>
			<!-- #include file="../includes/fr_folders.asp"-->
		</td>
		<td valign="top">
			<table border="0" cellpadding="2" cellspacing="2" width="600">		
				<tr>
					<td><%=dictLanguage("Rep_Inst_1")%></td>
				</tr>
				<tr>
					<td>
						<table width="100%" border="0" cellpadding="3" cellspacing="0">
							<tr bgcolor="<%=gsColorHighlight%>">
								<td nowrap><b class="bolddark"><%=dictLanguage("Document")%></b></td>
								<td nowrap><b class="bolddark"><%=dictLanguage("Date_Submitted")%></b></td>
								<td nowrap><b class="bolddark"><%=dictLanguage("Contact")%></b></td>
								<td><img src="<%=gsSiteRoot%>images/document.gif" WIDTH="11" HEIGHT="14"></td>
							</tr>
<%	do while not rsSearchDocs.eof

		strFileName = rsSearchDocs("filename")
		boolURL = FALSE
		if strFileName = "" then
			boolURL = TRUE
			strfileName = rsSearchDocs("url")
		end if

		Set rsContact=Server.CreateObject("ADODB.Recordset")
		sql = sql_GetEmployeesByID(rsSearchDocs("submitBy"))
		Call RunSQL(sql, rsContact)

		intRowcounter = intRowcounter + 1
	
		If intRowcounter MOD 2 = 1 then 
			colour = "bgcolor=#F0F8FF"
		Else
		 	colour = "bgcolor=#ffffff"
		End If %>
							<tr <%=colour%>>
								<td nowrap class="small">
								<%if boolURL then%>
								<a href="<%=SpaceEncode(strFileName)%>" class="small" title="<%=rsSearchDocs("description")%>" target="_blank"><%=SpaceDecode(strFileName)%></a> (URL)
								<%else%>
								<a href="library/<%=SpaceEncode(strFileName)%>" title="<%=rsSearchDocs("description")%>" class="small" target="_blank"><%=SpaceDecode(strFileName)%></a>
								<%end if%>
								</td>
								<td nowrap class="small"><%=rsSearchDocs("submitDate")%></td>
							    <td nowrap class="small"><a href="<%=gsSiteRoot%>employees/default.asp?employeeid=<%=rsSearchDocs("submitBy")%>" class="small"><%=rsContact("employeeName")%></a></td>
								<td>
								<%if trim(rsSearchDocs("submitBy")) = trim(session("employee_id")) then%>
									<a href="delete.asp?file=<%=rsSearchDocs("ID")%>" class="small" onClick="return confirm('<%=dictLanguage("Confirm_Delete_Document")%>');"><img src="<%=gsSiteRoot%>images/delete.gif" WIDTH="20" HEIGHT="19" border="0" alt="<%=dictLanguage("Delete")%>"></a>
								<%end if%>
								&nbsp;</td>
							</tr>
<%		rsContact.close
		set rsContact = nothing
		rsSearchDocs.MoveNext
	loop %>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<br>
<p align="center">
<a href="utilities.asp"><%=dictLanguage("Return_Search")%></a><br>
<a href="default.asp"><%=dictLanguage("Return_Document_Repository")%></a><br>
<a href="main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<%
rsSearchDocs.close
set rsSearchDocs = nothing

function SpaceEncode(strText)
	if strText <> "" and isNull(strText) = False then
		strText = replace(strText," ","%20")
	end if
	SpaceEncode = strText
end function

function SpaceDecode(strText)
	if strText <> "" and isNull(strText) = False then
		strText = replace(strText,"%20"," ")
	end if
	SpaceDecode = strText
end function
%>

<!--#include file="../includes/main_page_close.asp"-->

<% end if %>