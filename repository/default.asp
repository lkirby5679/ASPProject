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
<!--#include file="../includes/main_page_open.asp"-->
	 
<table border"0" cellpadding="2" cellspacing="2" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>"><td colspan="2" align="center" class="homeheader"><%=dictLanguage("Document_Repository")%></td></tr> 
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
						<table width="100%" border="0" cellpadding="2" cellspacing="0">
<%
sql = sql_GetSubFoldersByFolderID(folder)
Call RunSQL(sql, rsSubs)

sql = sql_GetDocumentsByFolderID(folder)
Call RunSQL(sql, rsSearch)
	
if (not rsSubs.eof) or (not rsSearch.eof) then%>
							<tr bgcolor="<%=gsColorHighlight%>" >
								<td nowrap><b class="bolddark"><%=dictLanguage("Document")%></b></td>
								<td nowrap><b class="bolddark"><%=dictLanguage("Date_Submitted")%></b></td>
								<td nowrap><b class="bolddark"><%=dictLanguage("Contact")%></b></td>
								<td><img src="<%=gsSiteRoot%>images/document.gif" WIDTH="11" HEIGHT="14"></td>
							</tr>
<%	do while not rsSubs.eof
		intRowcounter = intRowcounter + 1
		If intRowcounter MOD 2 = 1 then 
			colour = "bgcolor=#F0F8FF"
		Else
			colour = "bgcolor=#ffffff"
		End If 
		Response.Write "<tr " & bgcolor & ">"
		strFolderID		= rsSubs("folder_id")
		strFolderName	= trim(rsSubs("folder"))
		strFolderImage	= "<img src='" & gsSiteRoot & "images/close_folder.gif' WIDTH='19' HEIGHT='19'>" %>
								<td nowrap class="small">
									<%=strFolderImage%>
									<a href="default.asp?folder=<%=strFolderID%>" class="small" title="<%=strFolderName%>"><%=strFolderName%></a>
								</td>
								<td nowrap class="small">&nbsp;</td>
								<td nowrap class="small">&nbsp;</td>
								<td nowrap class="small">&nbsp;</td>
							</tr>		
<%		rsSubs.movenext
	loop
	do while not rsSearch.eof
		sql = sql_GetEmployeesByID(rsSearch("submitBy"))
		Call RunSQL(sql, rsContact)

		intRowcounter = intRowcounter + 1

		If intRowcounter MOD 2 = 1 then 
			colour = "bgcolor=#F0F8FF"
		Else
			colour = "bgcolor=#ffffff"
		End If 
		Response.Write "<tr " & bgcolor & ">"
		strFileName = rsSearch("fileName")
		if strFileName<>"" then
			boolURL = FALSE
			if right(lcase(strFileName),3) = "doc" or right(lcase(strFileName),3) = "txt" or right(lcase(strFileName),3) = "htm" or right(lcase(strFileName),3) = "rtf" or right(lcase(strFileName),3) = "xls" or right(lcase(strFileName),3) = "pdf" or right(lcase(strFileName),3) = "ppt" then
				file_image = right(lcase(strFileName),3) & ".gif"
			elseif right(lcase(strFileName),4) = "html" then
				file_image = "htm.gif"
			else
				file_image = "document.gif"
			end if 
		else 
			boolURL=TRUE
			strFileName = rsSearch("url")
			file_image = "htm.gif"
		end if %>
								<td nowrap class="small">
									<img src="<%=gsSiteRoot%>images/<%=file_image%>" align="baseline" WIDTH="11" HEIGHT="14">
									<%if boolURL then%>
									<a href="<%=SpaceEncode(strFileName)%>" class="small" title="<%=rsSearch("description")%>" target="_blank"><%=SpaceDecode(strFileName)%></a> (URL)
									<%else%>
									<a href="library/<%=SpaceEncode(strFileName)%>" title="<%=rsSearch("description")%>" class="small" target="_blank"><%=SpaceDecode(strFileName)%></a>
									<%end if%>
								</td>
								<td nowrap align="center" class="small"><%=rsSearch("submitDate")%></td>
								<td nowrap class="small"><a href="<%=gsSiteRoot%>employees/default.asp?employeeid=<%=rsSearch("submitBy")%>" class="small"><%=rsContact("employeeName")%></a></td>
								<td align="left" class="small">
<%		if trim(rsSearch("submitBy")) = trim(session("employee_id")) then%>
									<a href="delete.asp?file=<%=rsSearch("ID")%>" class="small" onClick="return confirm('<%=dictLanguage("Confirm_Delete_Document")%>');"><img src="<%=gsSiteRoot%>images/delete.gif" WIDTH="20" HEIGHT="19" border="0" alt="<%=dictLanguage("Delete")%>"></a>
<%		end if%>
								</td>
							</tr>
<%		rsSearch.MoveNext
	loop 
else %>						<tr>
								<td align="center"><b><%=dictLanguage("Folder_Empty")%></b></td>
							</tr>
<%end if%>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>


<%
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
 
 
 