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
<!--#include file="../../includes/SiteConfig.asp"-->
<!--#include file="../../includes/connection_open.asp"-->

<%
strAct		= sqlEncode(request("act"))
strID	    = sqlEncode(request("id"))
strDocID    = sqlEncode(request("doc_id"))
strDescr    = sqlEncode(request("description"))
strEmpID    = sqlEncode(request("employee_id"))
strFoldID   = sqlEncode(request("folder_id"))
strFoldName = sqlEncode(request("folder_name"))
strSub_to   = sqlEncode(request("sub_to"))

'used for keeping folder "session"
cur_folder = request("cur_folder")



select case strAct
	case "doc_edit"
		sql = sql_UpdateDocument(strID, strEmpID,strDescr,strFoldID)
		Call DoSQL(sql)
		session("msg") = dictLanguage("Updated_Document")
		Response.Redirect("default.asp?folder="&cur_folder)
	case "doc_delete"
		if session("permAdminFileRepository") then
			sql = sql_DeleteDocument(strDocID)
			Call DoSQL(sql)
			session("msg") = dictLanguage("Deleted_Document")
		else
			session("msg") = dictLanguage("NoDeleted_Document")
		end if
		Response.Redirect("default.asp?folder="&cur_folder)	
	case "folder_add"
		if strSub_to <> "" then
			cur_folder = strSub_to
		end if	
		sql = sql_InsertFolder(strFoldName, strSub_to)
		Call DoSQL(sql)
		session("msg") = dictLanguage("Added_Folder")
		Response.Redirect("default.asp?folder="&cur_folder)			
	case "folder_edit"
		sql = sql_UpdateFolder(strFoldID, strFoldName,strSub_to)
		Call DoSQL(sql)
		session("msg") = dictLanguage("Updated_Folder")
		Response.Redirect("default.asp?folder="&strFoldID)
	case "folder_delete"
		if session("permAdminFileRepository") then
			sql = sql_DeleteFolder(strFoldID)
			'Response.Write sql
			Call DoSQL(sql)
			session("msg") = dictLanguage("Deleted_Folder")
		else
			session("msg") = dictLanguage("NoDeleted_Folder")
		end if
		Response.Redirect("default.asp?folder="&strFoldID)											
    case else
end select    


%>

<!--#include file="../../includes/connection_close.asp"-->
