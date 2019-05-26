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
<!--#include file="includes/connection_open.asp"-->
<!--#include file="includes/mail.asp"-->
<!--#include file="includes/style.asp"-->

<% 
boolPasswordSent = false
if Request.form("submit")="submit" then
	emailaddress = SQLEncode(Request.Form("emailaddress"))
	if emailaddress = "" then
		strErrorMessage = dictLanguage("Error_PasswordNoEmail")
	else
		sql = sql_GetPasswordReminder(emailaddress)
		Call RunSQL(sql, rs)
		if not rs.eof then
			active = rs("active")
			if active = 1 or active = -1 or active then
				employeelogin = rs("employeelogin")
				emppassword = rs("password")
				msg = dictLanguage("YourUserIDIs") & employeelogin & VBCRLF
				msg = msg & dictLanguage("YourPasswordIs") & emppassword & VBCRLF
				Call SendEmail("", emailaddress, "", gsAdminEmail, "Password Reminder", msg, "", "", _
					"", "", "", FALSE)
				boolPasswordSent = TRUE
			else
				strErrorMessage = dictLanguage("Error_PasswordInactive")
			end if
		else
			strErrorMessage = dictLanguage("Error_PasswordNoRecord")
		end if
		rs.close
		set rs = nothing
	end if
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>
<head>
	<meta http-equiv="Pragma" content="no-cache">
	<title><%=gsSiteName%></title>
</head>

<body>
<table width="100%" border="0" cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top">
			<table border="0" cellpadding="2" cellspacing="0" width="100%" bgColor="<%=gsColorHighlight%>">
				<tr>
					<td valign="top"><a href="<%=gsSiteRoot%>"><img src="<%=gsSiteRoot%>images/intranet.gif" alt="Transworld Interactive Intranet" border="0"></a></td>
					<td align="right"><a href="<%=gsSiteRoot%>" target="_blank"><img src="<%=gsSiteRoot%>images/poweredbydiggersolutionsv11.gif" border="0" width="160" height="40" alt="powered by Transworld Interactive Projects"></a></td>
				</tr>
			</table>
			<table border="0" cellpadding="2" cellspacing="2" align="center" width="300">
				<tr bgcolor="<%=gsColorHighlight%>"><td align="center" class="homeheader" colspan="2"><%=dictLanguage("ForgetPassword")%></td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>
<% if boolPasswordSent then %>
				<tr><td colspan="2" align="center"><%=dictLanguage("Password_Inst2")%></td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>
				<tr><td colspan="2" align="center"><a href="default.asp?opt=1"><%=dictLanguage("Login")%></a></td></tr>
<% else %>
				<tr><td colspan="2"><%=dictLanguage("Password_Inst1")%></td></tr>
				<% if strErrorMessage <> "" then
					Response.Write "<tr><td colspan=""2"" align=""center""><font color=""red"">" & strErrorMessage & "</font></td></tr>"
				   end if %>
				<form action="forgotpassword.asp" method="POST" id="form1" name="form1">
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Email")%>: </td>
					<td><input type="text" name="emailaddress" class="formstyleLong"></td>
				</tr>
				<tr><td align="center" colspan="2"><input type="submit" value="submit" class="formButton" id="submit" name="submit"></td></tr>
				</form>
<% end if %>
			</table>
		</td>
	</tr>
</table>
</body>
</html>

<!--#include file="includes/connection_close.asp"-->
