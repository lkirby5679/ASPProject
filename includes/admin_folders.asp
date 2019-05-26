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

			<table border="0" width="100%" cellpadding="4" cellspacing="0" bgcolor="<%=gsColorBackground%>">
				<tr>
					<td bgcolor="<%=gsColorHighlight%>"><b class="homeheader"><%=dictLanguage("Folders")%></b></TD>
					<td valign="top"  bgcolor="<%=gsColorHighlight%>" align="right"><a href="default.asp?act=folder_add"><IMG  SRC="images/new.gif"   height="10" width="10" alt="<%=dictLanguage("Add_Folder")%>" border="0"></a></td>
				</tr>
				<tr>
					<td colspan="2" class="homecontent" nowrap>
<%	
'*****Record count per folder
if request("folder") <> "" then
	folder = request("folder")
else
	folder = 5
end if
thispage="default.asp"

sql = sql_GetFoldersCount()
'Response.Write sql & "<BR>"
Call RunSQL(sql, rsSearch)

' Create a dictionary object to hold counts 
set dCount = Server.CreateObject("Scripting.Dictionary")
do while not rsSearch.EOF
    if rsSearch(0) <> "" then
		dCount.add cstr(rsSearch(0)), trim(rsSearch(1))
    end if
    rsSearch.movenext
loop
rsSearch.close
set rsSearch=nothing

sql = sql_GetMainFolders()
Call RunSQL(sql, rsFolders2)
do while not rsFolders2.eof 
    folder_id = trim(rsFolders2("folder_id"))
    fcount = dCount.Item(trim(folder_id))
    if fcount = "" then
		fcount = 0
    end if
	strFolders = ""
    if folder_id = trim(folder) then
		boolCheck = TRUE
	else
		boolCheck = AreMyDescendantsKing(folder_id, folder, FALSE)
	end if
	if boolCheck then
		Call DisplayFolderInfo(folder_id, 0, 0, 0, 0)
	else
		Call DisplayFolderLine(folder_id, trim(rsFolders2("folder")), 0, fcount, FALSE)
	end if
	Response.Write strFolders 
	rsFolders2.movenext
loop
rsFolders2.Close
set rsFolders2 = nothing
%>	
					</td>
				</tr>
			</table>


<%
Private sub DisplayFolderInfo(fid, strLevel, boolFatherIsKing, boolBrotherIsKing, boolNewphewIsKing)
    folder_count = dCount.Item(trim(fid))
    if folder_count ="" then
		folder_count = 0
    end if

	sql = sql_GetFolderByID(fid)
	Call RunSQL(sql, rsF)
	if not rsF.eof then
		folder_name = rsF("folder")
		folder_subto = rsF("sub_to")
	end if
	rsF.close
	set rsF = nothing		

	if trim(fid) = trim(folder) then
		'Display my two kids but no grandkids
		sql = sql_GetSubFoldersByFolderID(fid)
		Call RunSQL(sql, rsSubFolders)
		while not rsSubFolders.eof
			Call DisplayFolderInfo(rsSubFolders("folder_id"), strLevel + 1, 1, 0, 0)
			rsSubFolders.movenext
		wend
		rsSubFolders.close
		set rsSubFolders = nothing		

		'I am the King, print me open 
		Call DisplayFolderLine(fid, folder_name, strLevel, folder_count, TRUE)
	
	elseif boolFatherIsKing then
		'print me closed
		Call DisplayFolderLine(fid, folder_name, strLevel, folder_count, FALSE)

	elseif boolBrotherIsKing then
		'display me closed, no children
		Call DisplayFolderLine(fid, folder_name, strLevel, folder_count, FALSE)
	
	else
		'if I have children, check to see if one of them is the king, if they are then print me open
		boolLocalOpen = AreMyDescendantsKing(fid, folder, FALSE)
		sql = sql_GetSubFoldersByFolderID(fid)
		Call RunSQL(sql, rsSubFolders)
		boolLocalBrother = FALSE
		while not rsSubFolders.eof
			if trim(rsSubFolders("folder_id")) = trim(folder) then
				boolLocalBrother = TRUE
			end if
			rsSubFolders.movenext
		wend
		if not rsSubFolders.bof then
			rsSubFolders.movefirst
		end if
		while not rsSubFolders.eof
			if boolLocalOpen then
				Call DisplayFolderInfo(rsSubFolders("folder_id"), strLevel + 1, 0, boolLocalBrother, 1)
			else
				Call DisplayFolderInfo(rsSubFolders("folder_id"), strLevel + 1, 0, boolLocalBrother, 0)
			end if
			rsSubFolders.movenext
		wend
		rsSubFolders.close
		set rsSubFolders = nothing			
		'if one of them is open then show me open
		if boolLocalOpen then
			Call DisplayFolderLine(fid, folder_name, strLevel, folder_count, TRUE)
		elseif boolNewphewIsKing then
			Call DisplayFolderLine(fid, folder_name, strLevel, folder_count, FALSE)
		elseif folder_subto = 0 then
			Call DisplayFolderLine(fid, folder_name, strLevel, folder_count, FALSE)			
		end if		
	end if
end sub

Sub DisplayFolderLine(fid, fName, intLevel, intCount, boolOpen)
	strSubFolders = ""
	for i = 1 to intLevel
		strSubFolders = strSubFolders & "<img src='" & gsSiteRoot & "images/dot_clear.gif' width='10' height='1' border='0'>"
	next
	
	strEditIcons = "<a href=""default.asp?folder="&fid&"&act=folder_edit&folder_id="&fid&"""><img valign='top' src='" & gsSiteRoot & "admin/images/edit.gif' WIDTH='10' HEIGHT='10' border='0' alt=""" & dictLanguage("Edit_Folder") & """></a>&nbsp;<a onClick=""javascript: return confirm('" & dictLanguage("Confirm_Delete_Folder") & "');"" href=""hnd_repository.asp?act=folder_delete&folder_id="&fid&"""><img src='" & gsSiteRoot & "admin/images/delete.gif' WIDTH='10' HEIGHT='10' border='0' alt=""" & dictLanguage("Delete_Folder") & """></a>&nbsp;<a href=""default.asp?folder="&fid&"&act=folder_add&sub_to="&fid&"""><img valign='top' src='" & gsSiteRoot & "admin/images/new.gif' WIDTH='10' HEIGHT='10' border='0' alt=""" & dictLanguage("Add_SubFolder") & """></a>"
	if boolOpen then
		strSubFolders = strSubFolders &"<img src='" & gsSiteRoot & "images/open_folder.gif' WIDTH='19' HEIGHT='19' border='0'>"
	else 
		strSubFolders = strSubFolders &"<img src='" & gsSiteRoot & "images/close_folder.gif' WIDTH='19' HEIGHT='19' border='0'>"
	end if
	strSubFolders = strSubFolders & "<a href='" & thispage & "?folder=" & fid & "' class='smallhome'>" & fName & "</a>&nbsp;<font class='smallhome'>(" & intCount & ")</font>&nbsp;"&strEditIcons&"<br>"
	strFolders = strSubFolders & strFolders 
end sub

Function AreMyDescendantsKing(fid, kingID, boolCheck)
	if not boolCheck then
		sql = sql_GetSubFoldersByFolderID(fid)
		Call RunSQL(sql, rsWhatzit)
		while not rsWhatzit.eof and not boolCheck
			if trim(rsWhatzit("folder_id")) = trim(kingID) then
				boolCheck = TRUE
			else
				boolCheck = AreMyDescendantsKing(rsWhatzit("folder_id"), kingID, boolCheck)	
			end if
			rsWhatzit.movenext
		wend
		rsWhatzit.close
		set rsWhatzit = nothing				
	end if
	AreMyDescendantsKing = boolCheck
End Function


%>