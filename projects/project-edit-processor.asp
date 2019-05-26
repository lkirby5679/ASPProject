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

If session("Description") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoName")
End If

If session("Client_ID") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectNoClient")
Elseif session("active") = 1 then 'make sure this client is active
	strClientSQL = sql_GetActiveClientsByID(session("Client_ID"))
	set rsClient = Conn.Execute(strClientSQL)
	if rsClient.EOF then 'client not active
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_ProjectInactiveClient")
	end if
	rsCLient.close
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

If session("strErrorMessage") <> "" then
	response.redirect "project-edit.asp?project_id=" & session("project_id") & ""
End If

If session("color")="" then
	session("color") = "Green"
end if

for each i in request.form 
	if session(i) = "" then
		session(i) = 0
	end if
next

sql = sql_UpdateProjects( _
	session("client_id"), _
	session("description"), _
	session("workorder_number"), _
	session("accountexec_id"), _
	session("start_date"), _
	session("launch_date"), _
	session("projecttype_id"), _
	session("active"), _
	session("comments"), _
	session("color"), _
	session("project_id"))
'response.write sql & "<br>"
Call DoSQL(sql)

'Now update the project budgets
'first delete all the existing
sql = sql_DeleteProjectBudgetsByID(session("project_id"))
'Response.Write sql & "<br>"
Call DoSQL(sql)

'then enter all the new ones
for i = 0 to 149
	if iEmployeeType(i) then
		sql = sql_InsertProjectsBudget(session("project_id"), i, iEmployeeID(i), iRate(i), iHours(i)) 
		'Response.Write sql & "<BR>"
		Call DoSQL(sql)
	end if
next
%>

<!--#include file="../includes/main_page_open.asp"-->

<p align="center">
<%=dictLanguage("Updated_Project")%><br><br>
<a href="project-view.asp"><%=dictLanguage("View_Projects")%></a><br>
<a href="project-add.asp"><%=dictLanguage("Add_Project")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<%for each i in request.form 
    session(i) = ""
  next %>

<!--#include file="../includes/main_page_close.asp"-->

