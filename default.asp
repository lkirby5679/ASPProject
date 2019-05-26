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
<!--#include file="includes/SiteConfig.asp"-->

<%
'check to see if the database is setup
On Error Resume Next
strConn = DB_CONNECTIONSTRING	
set conn=server.createobject("adodb.connection")
conn.Open strConn
If conn.errors.count <> 0 then
'	Response.Redirect gsSiteRoot & "admin/dbsetup.asp"
end if
conn.close
set conn = nothing
On Error GoTo 0
	
strOption = Request.QueryString("opt")

if gsAutoLogin then
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Enable cookie based logon if possible
	'''''''''''''''''''''''''''''''''''''''
	if request.Cookies("shark_user_id") <> "" and request.Cookies("shark_password") <> "" and strOption="" then
		response.redirect "logoncheck.asp"
	end if
end if
%>

<!--#include file="includes/style.asp"-->
<!--Transworld Interactive Projects-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>
<head>
	<meta http-equiv="Pragma" content="no-cache">
	<TITLE>Website Hosting | Website Design | Transworld Interactive</TITLE>
<META name="resource-type" content="document">
<META name="robots" content="index,follow">
<META name="Revisit-After" content="1 days">
<META name="Author" content="Professor Landon T. Kirby,Ph.D,D.Sc,DBA">
<META name="keywords" content="project,projects,transworld interactive project,transworld interactive,transworld,interactive,website hosting,website design,graphic design,seo,datacenter,web development,project management,search engine optimization">
<META name="description" content="Transworld Interactive for Project Hosting, Project Design, Project management, Datacenter. Transworld Interactive also does Network Intergration.">
<META name="distribution" content="global">
<META name="copyright" content="Transworld Interactive website copyrighted in 2011 by Professor Landon T. Kirby,Ph.D,D.Sc,DBA">
<meta name="no-email-collection" content="http://www.transworldinteractive.net/nospambots.htm">
<META NAME="page-topic" CONTENT="Transworld Interactive project management website hosting, Project management.">
</head>
<body OnLoad="document.form1.userid.focus();">
<table width="100%" border="0" cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top">
			<table border="0" cellpadding="2" cellspacing="0" width="100%" bgColor="<%=gsColorHighlight%>">
				<tr>
					<td valign="top"><a href="<%=gsSiteRoot%>"><img src="<%=gsSiteRoot%>images/<%=gsSiteLogo%>" alt="<%=gsSiteName%>" border="0"></a></td>
					<td align="right"><a href="<%=gsSiteRoot%>" target="_blank"><img src="images/poweredbyti.gif" border="0" width="160" height="40" alt="powered by Transworld Interactive Projects"></a></td>
				</tr>
			</table>
			<table border="0" cellpadding="2" cellspacing="2" align="center" width="300">
				<tr bgcolor="<%=gsColorHighlight%>"><td align="center" class="homeheader" colspan="2"><%=dictLanguage("Intranet_Login")%></td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>
				<form action="logoncheck.asp" method="POST" id="form1" name="form1">
				<input type="hidden" name="hname" value="hidden">
				<input type="hidden" name="ccokie" value="false">
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("ID")%>: </td>
					<td><input type="text" name="userid" class="formstyleShort"></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Password")%>: </td>
					<td><input type="password" name="password" class="formstyleShort"></td>
				</tr>
				<tr><td align="center" colspan="2"><input type="submit" value="logon" class="formButton"></td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>
				<tr><td align="center" colspan="2"><a href="forgotpassword.asp"><%=dictLanguage("ForgetPassword")%></a></td></tr>
				</form>
				<tr>
					<td align="center" colspan="2" class="alert">
						<p>
						<%=session("msg")%>
						<%session("msg") = ""%>
						<%If session("msg") = "" then
							session.abandon
						  End If %>
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>

