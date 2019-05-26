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
for each i in request.form 
    session(i) = SQLEncode(request.form(i))
next

If session("Name") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_No_ClientName")
End If
If session("Rep") = "" then
	Session("strErrorMessage") = Session("strErrorMessage") & "<br>" & dictLanguage("Error_No_ClientRep")
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
	response.redirect "client-add.asp"
End If

If Session("Standard_Rate")="" then
	Session("Standard_Rate")=0
end if
If Session("Client_Since")="" then
	Session("Client_Since")= Date()
end if

sql = sql_InsertClient ( _
	Session("Name"), _
	Session("Rep"), _
	1, _
	Session("Client_Since"), _
	Session("Standard_Rate"), _
	Session("Address1"), _
	Session("Address2"), _
	Session("City_State_Zip"), _
	Session("LiveSite_URL"), _
	Session("Contact_Name"), _
	Session("Contact_Phone"), _
	Session("Contact_Email"), _
	Session("Contact2_Name"), _
	Session("Contact2_Phone"), _
	Session("Contact2_Email"), _
	Session("Notes"))
'response.write(sql)
Call DoSQL(sql)
%>

<!--#include file="../includes/main_page_open.asp"-->

<P align="center">
<b><%=dictLanguage("Client_Added")%></b><br><br>
<a href="client-view.asp"><%=dictLanguage("View_Clients")%></a><br>
<a href="client-add.asp"><%=dictLanguage("Add_Client")%></a><br>
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<%
for each i in request.form 
    session(i) = ""
next
%>

<!--#include file="../includes/main_page_close.asp"-->
