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
<%
strAct		= trim(request("act"))
strEmpID	= trim(request("id"))
strDocID    = trim(request("doc_id"))
strFoldID   = trim(request("folder_id"))
'used for keeping folder "session"
folder		= request("folder")
if folder = "" then
	folder = 5
end if
%>


<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->




<table border="0" cellspacing="2" cellpadding="2">
<tr><td width="" valign="top">

	<table width="300px" border="0" cellspacing="0" cellpadding="2">
		<tr>
		    <td colspan="4" valign="top">
				<!-- #include file="../../includes/admin_folders.asp"-->
		    </td>
		</tr>		
<%
sql = sql_GetSubFoldersByFolderID(folder)
Call RunSQL(sql, rsSubs)

sql = sql_GetDocumentsByFolderID(folder)
Call RunSQL(sql, rsSearch)
%>
		<tr bgcolor="<%=gsColorHighlight%>" >
			<td colspan="3" nowrap><b class="bolddark"><%=dictLanguage("Document")%></b></td>
			<td align="right"><a href="default.asp?act=doc_add&folder=<%=folder%>"><IMG align="right" SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_Document")%>" border="0"></a></td>
		</tr>
		
<%
if (not rsSubs.eof) or (not rsSearch.eof) then
	do while not rsSubs.eof
		intRowcounter = intRowcounter + 1
		If intRowcounter MOD 2 = 1 then 
			colour = "bgcolor=#F0F8FF"
		Else
			colour = "bgcolor=#ffffff"
		End If 
		
		Response.Write "<tr " & bgcolor & ">"
		strFolderID = rsSubs("folder_id")
		strFolderName	= trim(rsSubs("folder"))
		strFolderImage	= "<img src='" & gsSiteRoot & "../../images/close_folder.gif' WIDTH='19' HEIGHT='19'>" %>
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
		strFileID = rsSearch("id")
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
		
			<td colspan="3" nowrap class="small"><img src="<%=gsSiteRoot%>images/<%=file_image%>" align="baseline" WIDTH="11" HEIGHT="14">
				<%if boolURL then%>
					<%=SpaceDecode(strFileName)%> (URL)
				<%else%>
					<%=SpaceDecode(strFileName)%>
				<%end if%>
			</td>
			<td align="right" class="small">
				<a href="default.asp?act=doc_edit&doc_id=<%=strFileID%>&folder=<%=folder%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Document")%>"></a>&nbsp;
				<a href="hnd_repository.asp?act=doc_delete&doc_id=<%=strFileID%>&cur_folder=<%=folder%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Delete_Document")%>');"><IMG SRC="../images/delete.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Delete_Document")%>"></a>
			</td>
		</tr>
		<%
		rsSearch.MoveNext
	loop 
else %>	
		<tr>
			<td align="center"><b><%=dictLanguage("Folder_Empty")%></b></td>
		</tr>
<%
end if%>
	</table>

</td><td align="center" width="100%" valign="top">

		<table border="0" cellspacing="2" cellpadding="2" width="100%">
			<tr bgcolor="<%=gsColorHighlight%>">
				<td align="Center" class="homeheader"><%=dictLanguage("Workspace")%></td>
			</tr>
<% 
SQL = sql_GetActiveEmployees()
Call RunSQL(sql, rs)
             
sqlFolders = sql_GetAllFolders()
Call RunSQL(sqlFolders, rsFolders)   
		
select case strAct
	case "doc_add" %>
			<tr>
				<td>
					<table border="0" cellpadding="2" cellspacing="2" width="400">	
						<tr><td colspan=2><hr color="<%=gsColorBackground%>"></td></tr>
						<tr><td colspan="2"><b class="bolddark">Upload a Document:</b></td></tr>
						<tr>
							<td colspan=2><%=dictLanguage("DocRepInst_1")%></td>
						</tr>
						<form method="post" action="upload.asp" enctype="multipart/form-data" id=form2 name=form2>
						<tr>
							<td><%=dictLanguage("Contact")%>:</td>
							<td>
								<Select Name="byEmployee" size="1" class="formstyleLong">
							<%While Not rs.EOF %>
									<option value="<%=rs("employee_ID")%>"><%=rs("employeeName")%></option>
							<%	rs.MoveNext
							  Wend
							  rs.Close
							  set rs = nothing %>		
								</Select>
							</td>
						</tr>				
							<td><%=dictLanguage("Folder")%>: </td>
							<td>
								<Select Name="folder" size="1" class="formstyleLong">
							<%While Not rsFolders.EOF %>
									<option value="<%=rsFolders("folder_id")%>" <%if trim(rsFolders("folder_id")) = trim(folder) then Response.write "Selected"%>><%=rsFolders("folder")%></option>
							<%	rsFolders.MoveNext
							  Wend
							  rsFolders.Close
							  set rsFolders = nothing%>		
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
						<%if session("permAdminFileRepository") then %>
						<tr>
							<td align="center" colspan=2><input type="Submit" value="Submit" class="formButton"></td>
						</tr>
						<%end if%>
						<input TYPE="HIDDEN" NAME="where" value="<%=gsSiteRoot%>repository/library">
						</form>
					</table>				
				</td>
			</tr>			
			
<%	case "doc_edit"
		SQL = sql_GetDocumentByDocumentID(strDocID)
		Call RunSQL(sql, rsDoc)	
		if not rsDoc.eof then
			loc_ID = rsDoc("ID")
			loc_FileName = rsDoc("FileName")	
			loc_submitdate = rsDoc("submitdate")	
			loc_submitby = rsDoc("submitby")	
			loc_description = rsDoc("description")	
			loc_folder = rsDoc("folder")	
			loc_url = rsDoc("url")	%>
			<tr>
				<td>
					<table border="0" cellpadding="2" cellspacing="2" width="400">	
						<tr><td colspan="2"><b class="bolddark"><%=dictLanguage("Edit_Document")%>:</b></td></tr>
						<form method="post" action="hnd_repository.asp"id=form2 name=form2>
						<tr>
						<tr>
							<td><%=dictLanguage("Contact")%>:</td>
							<td>
								<Select Name="Employee_id" size="1" class="formstyleLong">
							<%While Not rs.EOF %>
									<option <%if loc_submitby = rs("employee_ID")then Response.Write"SELECTED" end if %> value="<%=rs("employee_ID")%>"><%=rs("employeeName")%></option>
							<%	rs.MoveNext
							  Wend
							  rs.Close
							  set rs = nothing %>		
							</Select>
							</td>
						</tr>				
							<td><%=dictLanguage("Folder")%>: </td>
							<td>
								<Select Name="folder_id" size="1" class="formstyleLong">
							<%While Not rsFolders.EOF %>
									<option <%if loc_folder = rsFolders("folder_id")then Response.Write"SELECTED" end if %> value="<%=rsFolders("folder_id")%>"><%=rsFolders("folder")%></option>
							<%	rsFolders.MoveNext
							  Wend
							  rsFolders.Close
							  set rsFolders = nothing%>		
								</Select>
							</td>
						</tr>
						<tr>
							<td valign="top"><%=dictLanguage("Description")%>: </td>
							<td><textarea name="description" rows="4" class="formstyleLong"><%=loc_description%></textarea></td>
						</tr>
						<%if session("permAdminFileRepository") then %>
						<tr>
							<td align="center" colspan=2><input type="Submit" value="Submit" class="formButton" id=Submit1 name=Submit1></td>
						</tr>
						<%end if%>
						<input TYPE="HIDDEN" NAME="cur_folder" value="<%=folder%>">
						<input TYPE="HIDDEN" NAME="act" value="<%=strAct%>">
						<input TYPE="HIDDEN" NAME="id" value="<%=loc_id%>">
						</form>
					</table>
				</td>
			</tr>
		<%else%>	
			<tr><td><%=dictLanguage("No_Document_Found")%></td></tr>	
		<%end if%>	

<%	case "folder_add"
		sub_to = request("sub_to")
		if sub_to <> "" then
			disp_head = "Sub-"
		else
			sub_to = 0	
		end if%>
			<tr>
				<td>
					<table border="0" cellpadding="2" cellspacing="2" width="400">	
						<tr><td colspan="2"><b class="bolddark"><%=dictLanguage("Add")%>&nbsp;<%=disp_head%><%=dictLanguage("Folder")%>:</b></td></tr>
						<form method="post" action="hnd_repository.asp"id=form2 name=form2>
						<tr>
							<td valign="top"><%=dictLanguage("Folder_Name")%>: </td>
							<td><input type="text" name="folder_name" maxlength="30" size="30" class="formstyleLong"></td>
						</tr>
						<%if session("permAdminFileRepository") then %>
						<tr>
							<td align="center" colspan=2><input type="Submit" value="Submit" class="formButton" id=Submit1 name=Submit1></td>
						</tr>
						<%end if%>
						<input TYPE="HIDDEN" NAME="cur_folder" value="<%=folder%>">
						<input TYPE="HIDDEN" NAME="act" value="<%=strAct%>">
						<input TYPE="HIDDEN" NAME="sub_to" value="<%=sub_to%>">
						</form>
					</table>
				</td>
			</tr>
<%	case "folder_edit"
		SQL = sql_GetFolderByID(strFoldID)
		Call RunSQL(sql, rsCurFolder)	
		if not rsCurFolder.eof then
			loc_ID     = rsCurFolder("folder_id")
			loc_folder = rsCurFolder("folder")
			loc_sub_to = rsCurFolder("sub_to") %>
			<tr>
				<td>
					<table border="0" cellpadding="2" cellspacing="2" width="400">	
						<tr><td colspan="2"><b class="bolddark"><%=dictLanguage("Edit_Folder")%>:</b></td></tr>
						<form method="post" action="hnd_repository.asp"id=form2 name=form2>
						<tr>
							<td valign="top"><%=dictLanguage("Folder_Name")%>: </td>
							<td><input type="text" name="folder_name" value="<%=loc_folder%>" maxlength="30" size="30" class="formstyleLong"></td>
						</tr>
						<tr>
							<td><%=dictLanguage("SubFolder_Of")%>: </td>
							<td>
								<Select Name="sub_to" size="1" class="formstyleLong">
									<option <%if loc_sub_to = 0 then Response.Write"SELECTED" end if %> value="0"><%=dictLanguage("None_TopLevelFolder")%></option>	
							 <%While Not rsFolders.EOF %>
									<option <%if loc_sub_to = rsFolders("folder_id")then Response.Write"SELECTED" end if %> value="<%=rsFolders("folder_id")%>"><%=rsFolders("folder")%></option>
							 <%	rsFolders.MoveNext
							   Wend
							   rsFolders.Close
							   set rsFolders = nothing	%>		
								</Select>
							</td>
						</tr>
						<%if session("permAdminFileRepository") then %>					
						<tr>
							<td align="center" colspan=2><input type="Submit" value="Submit" class="formButton" id=Submit1 name=Submit1></td>
						</tr>
						<%end if%>
						<input TYPE="HIDDEN" NAME="cur_folder" value="<%=folder%>">
						<input TYPE="HIDDEN" NAME="act" value="<%=strAct%>">
						<input TYPE="HIDDEN" NAME="folder_id" value="<%=strFoldID%>">
						</form>
					</table>
				</td>
			</tr>
		<%else%>	
			<tr><td><%=dictLanguage("No_Folder_Found")%></td></tr>	
	   <%end if%>								   								
<%	case else %>
			<tr>
				<td>
					<p class="alert"><%=session("msg")%></p>
					<%session("msg")=""%>
				
					<%=dictLanguage("DocRepInst_2")%> 
					(<img src="../images/edit.gif" WIDTH="10" HEIGHT="10" border="0" alt="<%=dictLanguage("Edit")%>"> = <%=dictLanguage("Edit")%>, <img src="../images/delete.gif" WIDTH="10" HEIGHT="10" border="0" alt="<%=dictLanguage("Delete")%>"> = <%=dictLanguage("Delete")%>, <img src="../images/new.gif" WIDTH="10" HEIGHT="10" border="0" alt="<%=dictLanguage("Add")%>"> = <%=dictLanguage("Add")%>)
					<%=dictLanguage("DocRepInst_3")%>
				</td>
			</tr>			
			
<%end select %>	
		</table>
</td>
</tr>
</table>

<p align="center"><a href=".."><%=dictLanguage("Return_Admin_Home")%></a></p>

<%
Function Permissions(strPerm)
	if strPerm <> 1 then
		strPerm = 0
	end if
	Permissions = strPerm
End Function %>

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

<!--#include file="../../includes/main_page_close.asp"-->
