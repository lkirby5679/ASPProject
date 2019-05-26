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
for each i in Request.Form
	session(i) = SQLEncode(Request.Form(i))
next

If session("Rep") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_No_ClientRep")
End If

If session("Name") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_No_ClientName")
End If

If session("client_id") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("No_Client_ID")
End If

If session("Client_Since") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_No_ClientStart")
else
	If not isdate(session("Client_Since")) then
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_Invalid_ClientStart")
	end if
End If
If session("Standard_Rate") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_No_ClientRate")
else
	If not isnumeric(session("Standard_Rate")) then
		Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_Invalid_ClientRate")
	else
		Session("Standard_Rate") = int(Session("Standard_Rate"))
	End If
End If

If Session("strErrorMessage") <> "" then
	response.redirect "client-edit.asp?client_id=" & session("client_id") & ""
End If

if session("Active") = "True" then
	Active = 1
else 
	Active = 0
end if

sql = sql_UpdateClient( _
	session("Name"), _
	session("Rep"), _
	session("Client_Since"), _
	session("Standard_Rate"),  _
	session("Address1"), _
	session("Address2"), _
	session("City_State_Zip"), _
	session("LiveSite_URL"), _
	session("Contact_Name"), _
	session("Contact_Phone"), _
	session("Contact_Email"), _
	session("Contact2_Name"), _
	session("Contact2_Phone"), _
	session("Contact2_Email"), _
	Active, _
	session("client_id"))
'response.write(sql)
Call DoSQL(sql)
%>

<!--#include file="../includes/main_page_open.asp"-->

<p align="center">
<%=dictLanguage("Updated_Client")%><br><br>
<a href="client-view.asp"><%=dictLanguage("View_Clients")%></a><br>
<a href="client-add.asp"><%=dictLanguage("Add_Client")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<%
for each i in Request.Form
	session(i) = ""
next
%>

<!--#include file="../includes/main_page_close.asp"-->

