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

		</td>
	</tr>
</table>

</body>
</html>

<!--#include file="connection_close.asp"-->

<%
Sub ConfirmPermission(strPermissionToCheck, redirectURL)
	if redirectURL = "" then
		redirectURL = gsSiteURL
	end if
	if not session(strPermissionToCheck) then
		Response.Write "<script language=""Javascript"">"
		Response.Write "strOK = alert('You do not have permissions to view this page');"
		Response.Write "window.location = """ & redirectURL & """;"
		Response.Write "</script>"	
	end if
End Sub

%>