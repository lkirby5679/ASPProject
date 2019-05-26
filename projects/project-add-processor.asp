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

<%
dim iEmployeeType(150)
dim iEmployeeID(150)
dim iRate(150)
dim iHours(150)
for each i in request.form 
    session(i) = SQLEncode(request.form(i))
	'response.write i & " = " & session(i) & "<BR>"
	if right(i,4) = "Rate" then
		xp=split(i,"x")
		ie = xp(0)
		iEmployeeType(ie) = TRUE
		iRate(ie) = session(i)
	end if
	if right(i,5) = "Hours" then
		xp=split(i,"x")
		ie = xp(0)
		iEmployeeType(ie) = TRUE
		iHours(ie) = session(i)
	end if		
	if right(i,11) = "Employee_ID" then
		xp=split(i,"x")
		ie = xp(0)
		iEmployeeType(ie) = TRUE
		if session(i)<>"" then
			iEmployeeID(ie) = session(i)
		else
			iEmployeeID(ie) = 0
		end if
	end if	
next

If session("ProjectName") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoName")
End If

If session("Client_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoClient")
Else 'make sure this client is active
	sql = sql_GetActiveClientsByID(session("Client_ID"))
	Call RunSQL(sql, rsClient)
	if rsClient.EOF then 'client not active
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectInactiveClient")
	end if
	rsClient.close
	set rsClient = nothing
End if

If session("AccountExec_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoRep")
End If

If session("WorkOrder_Number") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoWO")
End If

If session("Start_Date") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoStart")
else
	If not isdate(session("Start_Date")) then
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectInvalidStart")
	end if
End If

If session("launch_date") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoEnd")
else
	If not isdate(session("Launch_Date")) then
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectInvalidEnd")
	end if
End If

If session("ProjectType_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoType") 
End If

If Session("strErrorMessage") <> "" then
	response.redirect "project-add.asp"
End If

If request.form("color")="" then
	session("color") = "Green"
end if

for each i in request.form 
	if session(i) = "" then
		session(i) = 0
	end if
next

if DB_BOOLACCESS then
	sql = sql_InsertProject( _
		session("client_id"), _
		session("ProjectName"), _
		session("workorder_number"), _
		session("accountexec_id"), _
		1, _
		session("start_date"), _
		session("launch_date"), _
		session("projecttype_id"), _
		1, _
		session("comments"), _
		session("color"))	
	'Response.Write sql & "<BR>"	
	Call DoSQL(sql)
	sql = sql_SelectLatestProjectID()
	Call RunSQL(sql, rs)
	if not rs.eof then
		qid = rs("maxID")
	end if
	rs.close
	set rs = nothing
else 
	sql = sql_InsertProjectSQLServer( _
		session("client_id"), _
		session("ProjectName"), _
		session("workorder_number"), _
		session("accountexec_id"), _
		1, _
		session("start_date"), _
		session("launch_date"), _
		session("projecttype_id"), _
		1, _
		session("comments"), _
		session("color"))
	'Response.Write sql & "<BR>"	
	Call RunSQL(sql, rs)
	if not rs.eof then
		qid = rs("IdentityInsert")
	end if
	rs.close
	set rs = nothing
end if

'Now update the project budgets
'first delete all the existing
sql = sql_DeleteProjectBudgetsByID(qid)
Call DoSQL(sql)

'then enter all the new ones
for i = 0 to 149
	if iEmployeeType(i) then
		sql = sql_InsertProjectsBudget(qid, i, iEmployeeID(i), iRate(i), iHours(i)) 
		'Response.Write sql & "<BR>"
		Call DoSQL(sql)
	end if
next

'first we need to get the client name for the next two steps
sql = sql_GetClientsByID(session("client_id"))
Call RunSQL(sql, rs)
client_name = "Unknown"
if not rs.eof then
	client_name = rs("client_name")
end if
rs.close
set rs = nothing

forum_id = 0
if gsDiscussion then
	'create a discussion forum for the project and get its id
	forumName = left("Project: " & client_name & " - " & session("ProjectName") & "", 50)
	forumDesc = left("project discussion forum", 50)
	sql = sql_InsertForum(forumName, forumDesc, date(), "none", 1)
	Call DoSQL(sql)
	sql = sql_GetMyLatestForum(forumName, forumDesc, date())
	Call RunSQL(sql, rs)
	if not rs.eof then
		forum_id = rs("forum_id")
	end if
	rs.close
	set rs = nothing
end if

folder_id = 0
if gsFileRepository then
	'create a document repository folder for the project
	'11 is the projects folder - if you have changed it, removed it, whatever then you need to 
	'  change it below 
	folder_id = 11
	folderName = left(client_name & ": " & session("ProjectName") & "", 30)
	sql = sql_InsertFolder(folderName, folder_id)
	Call DoSQL(sql)
	sql = sql_GetMyLatestFolder(folderName, folder_id)
	Call RunSQL(sql, rs)
	if not rs.eof then
		folder_id = rs("folder_id")
	end if
	rs.close
	set rs = nothing
end if

if gsDiscussion or gsFileRepository then
	'update the forum_id, and folder_id fields in the project
	sql = sql_UpdateProjectForumAndFolder(qid, forum_id, folder_id)
	Call DoSQL(sql)
end if
%>

<!--#include file="../includes/main_page_open.asp"-->


<p align="center">
<%=dictLanguage("Added_Project")%><br><br>
<a href="project-view.asp"><%=dictLanguage("View_Projects")%></a><br>
<a href="project-add.asp"><%=dictLanguage("Add_Project")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br></p>

<%
for each i in request.form 
    session(i) = ""
next
%>

<!--#include file="../includes/main_page_close.asp"-->


