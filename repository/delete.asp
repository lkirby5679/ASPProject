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

<%  if request("file") = "" then
		response.redirect (Request.ServerVariables("HTTP_REFERER"))
	end if%>
<!--#include file="../includes/SiteConfig.asp"-->
<!--#include file="../includes/connection_open.asp"-->
<%  sqlFile = sql_GetDocumentByDocumentID(request("file"))
	Call RunSQL(sqlFile, rsFile)
	if not rsFile.eof then
		strFilename = rsFile("filename")
	end if
	rsFile.close
	set rsFile = nothing
	if strFileName <> "" then
		strFileName = server.MapPath(gsSiteRoot & "repository/library/" & strFileName) 
		Set MyFileObject = Server.CreateObject("Scripting.FileSystemObject")
		if MyFileObject.FileExists(strFilename) then
			boolDelete = MyFileObject.DeleteFile(strFilename, TRUE)
		end if
		set MyFileObject = nothing
	end if	
	sqlDelete = sql_DeleteDocument(request("file"))
	Call DoSQL(sqlDelete) %>
<!--#include file="../includes/connection_close.asp"-->	
<%	response.redirect "default.asp" %>	