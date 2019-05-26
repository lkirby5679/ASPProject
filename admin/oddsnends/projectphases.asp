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
<%
strAct		= trim(request("act"))
strID		= trim(request("id"))
%>


<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<table border="0" cellspacing="2" cellpadding="2">
<tr><td valign="top">

<table border="0" cellspacing="2" cellpadding="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td colspan="4" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Project_Phases")%></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><a href="projectphases.asp?act=add"><IMG SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_Project_Phase")%>" border="0"></a></td>
	</tr>
<% strSQL = sql_GetProjectPhases()
   Call RunSQL(strSQL, rs)
   do while not rs.eof
		pID = trim(rs("projectphaseid"))
		pName = trim(rs("projectphasename")) %>
	<tr>
		<td valign="top"><%=pID%></td>
		<td valign="top" nowrap><%=pName%></td>
		<td valign="top"><a href="projectphases.asp?act=edit&id=<%=pID%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Project_Phase")%>"></a></td>
		<td valign="top"><a href="projectphases.asp?act=delete&id=<%=pID%>" onClick="return confirm('<%=dictLanguage("Confirm_Delete_Project_Phase")%>');"><IMG SRC="../images/delete.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Delete_Project_Phase")%>"></a></td>
	</tr>
<%		rs.movenext
	loop 
	rs.close
	set rs = nothing %>
</table>

</td><td align="center" width="100%" valign="top">

<table border="0" cellspacing="2" cellpadding="2" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td align="Center" class="homeheader"><%=dictLanguage("Workspace")%></td>
	</tr>
</table>

<%if strAct = "add" then %>

<form name="strForm" id="strForm" action="projectphases.asp" method="POST">
<input type="hidden" name="act" value="add2">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Add_Project_Phase")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Project_Phase_ID")%>:</b></td>
		<td valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Project_Phase_Name")%>:</b></td>
		<td valign="top"><input type="text" name="pName" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>			
</table>
</form>

<%elseif strAct = "add2" then
	pName = SQLEncode(request("pName"))
	if pName = "" then
		pName = "New Project Phase"
	end if
	sql = sql_InsertProjectPhase(pName)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Added_Project_Phase")
	Response.Redirect "projectphases.asp?act=hnd_thnk" %>

<%elseif strAct = "edit" and strID <> "" then 
	sql = sql_GetProjectPhasesByID(strID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		pID			= rs("projectphaseid")
		pName		= trim(rs("projectphasename"))
	else
		Response.Redirect "projectphases.asp"
	end if
	rs.close
	set rs = nothing
%>

<form name="strForm" id="strForm" action="projectphases.asp" method="POST">
<input type="hidden" name="act" value="edit2">
<input type="hidden" name="id" value="<%=pID%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Edit_Project_Phase")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Project_Phase_ID")%>:</b></td>
		<td valign="top"><%=pID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Project_Phase_Name")%>:</b></td>
		<td valign="top"><input type="text" name="pName" value="<%=pName%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%		if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%		end if %>			
</table>
</form>

<%elseif strAct = "edit2" and strID <> "" then
	pName = SQLEncode(request("pName"))
	if pName = "" then
		pName = "New Project Phase"
	end if
	sql = sql_UpdateProjectPhase(strID, pName)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Update_Project_Phase")
	Response.Redirect "projectphases.asp?act=hnd_thnk" %>

<%elseif strAct = "delete" and strID <> "" then 
	sql = sql_DeleteProjectPhase(strID)
	Call DoSQL(sql)
	session("strMessage") = dictLanguage("Deleted_Project_Phase")
	Response.Redirect "projectphases.asp?act=hnd_thnk"%>

<%elseif strAct = "hnd_thnk" then %>
	<p align="Center"><%=session("strMessage")%></p>
	<%	session("strMessage") = ""%>

<%end if%>
</td>
</tr>
</table>

<p align="center"><a href="projectphases.asp"><%=dictLanguage("Return_Admin_Home")%></a></p>

<!--#include file="../../includes/main_page_close.asp"-->
