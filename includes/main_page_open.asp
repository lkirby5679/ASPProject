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

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>
<head>
	<title><%=gsSiteName%></title>
	<meta name="description" content="<%=gsMetaDescription%>">
	<meta name="keywords" content="<%=gsMetaKeywords%>">
</head>

<!--#include file="js.asp"-->

<!--#include file="style.asp"-->

<body <%if gsDHTMLNavigation then%>onLoad="javascript: checkNavTimer();"<%end if%>>

<%	if gsNavigation then %>
<!--#include file="navSpans.asp"-->
<%  end if %>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<form name="timeClock" action="timeClock.asp">
	<tr>
		<td valign="top">
<% if session("cookieMessage")<>"" then 
		Response.Write "<b class='bolddark'>&nbsp;" & session("cookieMessage") & "</b>"
		session("cookieMessage") = "" 
	end if %>
		</td>
		<td align="right" valign="top">&nbsp;
<% if boolHourly and lcase(Right(Request.ServerVariables("URL"),8))="main.asp" then %>
				<input type="hidden" name="employeeID" value="<%=strEmpID%>">		
				<input type="Submit" name="submit" value="<%=strTimeClockValue%>" class="formButton">
<% end if %>
		</td>
	</form>
	</tr>
	<tr>
		<td colspan="2">
<% if session("error_message")<>"" then 
		Response.Write "<b class='alert'>&nbsp;" & session("error_message") & "</b><br>"
		session("error_message") = ""
   end if
   if Session("strErrorMessage")<>"" then 
		Response.Write "<b class='alert'>&nbsp;" & session("strErrorMessage") & "</b><br>"
		session("strErrorMessage") = ""
   end if %>		
		</td>
	</tr>
</table>


<table width="100%" border="0" cellpadding="0" cellspacing="1" align="center">
	<tr>
		<td bgColor="<%=gsColorHighlight%>" colspan="2" valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="top"><a href="<%=gsSiteRoot%>"><img src="<%=gsSiteRoot%>images/<%=gsSiteLogo%>" alt="<%=gsSiteName%>" border="0"></a></td>
					<td align="right" valign="top">
						<table border="0" cellpadding="2" cellspacing="2">
							<tr>
								<td valign="top">
									<%=dictLanguage("Todays_Date")%>:&nbsp;&nbsp;<%=FormatDateTime(Now(),2)%>&nbsp;&nbsp;&nbsp;&nbsp;
								</td>
								<td align="center" valign="top">
									<%=dictLanguage("Logged_In")%>:&nbsp;&nbsp;<font class="boldwhite"><%=session("userid")%>&nbsp;&nbsp;&nbsp;&nbsp;</font>
								</td>
								<td><a href="http://www.transworldinteractive.net" target="_blank"><img src="<%=gsSiteRoot%>images/poweredbyti.gif" border="0" width="160" height="40" alt="powered by Transworld Interactive Projects"></a></td>
							</tr>
						</table>
					</td>
				</tr>
<%	if gsNavigation then %>
				<tr>
					<td colspan="2" bgcolor="<%=gsColorBackground%>">
						<!-- #include file="nav.asp"-->
					</td>
				</tr>
<%	end if %>
			</table>
		</td>
	</tr>
	<tr>
		<td valign="top" name="strMain" id="strMain" style="display: block;">