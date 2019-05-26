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
strEventID	= trim(request("id"))
%>


<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<table border="0" cellspacing="2" cellpadding="2">
<tr><td valign="top">

<table border="0" cellspacing="2" cellpadding="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td colspan="3" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Resources")%></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=add&"><IMG SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_New_Resource")%>" border="0"></a></td>
	</tr>
<% strSQL = sql_GetAllResources()
   Call RunSQL(strSQL, rs)
   do while not rs.eof
		rID			= rs("resources_id")
		rURL		= trim(rs("resources_URL"))
		rTitle		= trim(rs("resources_title")) 
		rLive		= rs("resources_live") 
		if rLive = 1 or rLive = -1 or rLive then
			rLive = TRUE
		else
			rLive = FALSE
		end if %>
	<tr>
		<td valign="top" nowrap>
			<a href="<%=rURL%>" target="_blank">
<%		if not rLive then 
			Response.Write "<font color=""red"">" & rTitle & "&nbsp;(Inactive)</font>"
		else
			Response.Write rTitle 
		end if %>
			</a>
		</td>
		<td valign="top"><a href="default.asp?act=edit&id=<%=rID%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Resource")%>"></a></td>
		<td valign="top"><a href="default.asp?act=delete&id=<%=rID%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Delete_Resource")%>');"><IMG SRC="../images/delete.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Delete_Resource")%>"></a></td>
	</tr>
<%		rs.movenext
	loop 
	rs.close
	set rs = nothing %>
</table>

<br><br>

</td><td align="center" width="100%" valign="top">

<table border="0" cellspacing="2" cellpadding="2" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td align="Center" class="homeheader"><%=dictLanguage("Workspace")%></td>
	</tr>
</table>

<%if strAct = "add" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndAdd">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Add_New_Resource")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_ID")%>:</b></td>
		<td valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_Title")%>:</b></td>
		<td valign="top"><input type="text" name="eventHeading" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_URL")%>:</b></td>
		<td valign="top"><input type="text" name="eventURL" value="" class="formstyleLong" maxlength="100"></td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_Caption")%>:</b></td>
		<td valign="top"><textarea name="eventAbstract" rows="4" class="formstyleLong" maxlength="255"></textarea></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Live")%>:</b></td>
		<td valign="top"><input type="checkbox" name="eventLive" value="1" Checked></td>
	</tr>		
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminResources") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>			
</table>
</form>

<%elseif strAct = "hndAdd" then 
	eventDate			= date()
	eventHeading		= SQLEncode(trim(request("eventHeading")))
	eventURL			= SQLEncode(trim(request("eventURL")))
	eventAbstract		= left(SQLEncode(trim(request("eventAbstract"))), 255)
	eventLive			= request("eventLive")
	eventDeleted		= 0
	eventEnteredBy		= session("employee_id")
	
	if eventLive <> 1 then
		eventLive = 0
	end if
	
	sql = sql_InsertResource(eventDate, eventHeading, eventURL, eventAbstract, _
		eventLive, eventDeleted, eventEnteredBy)
	Call DoSQL(sql)
	
	session("strMessage") = dictLanguage("Resource_Added")
	Response.Redirect "default.asp?act=hnd_thnk"
	

  elseif strAct = "edit" and strEventID <> "" then 
	sql = sql_GetResourcesByID(strEventID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		rID				= rs("resources_id")
		rTitle			= trim(rs("resources_title"))
		rURL			= trim(rs("resources_URL"))
		rCaption		= trim(rs("resources_caption"))
		rLive 			= rs("resources_live")
		if rLive = 1 or rLive = -1 or rLive then
			rLive = TRUE
		else
			rLive = FALSE
		end if
	else
		Response.Redirect "default.asp"
	end if
	rs.close
	set rs = nothing
%>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEdit">
<input type="hidden" name="id" value="<%=strEventID%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Edit_Resource")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_ID")%>:</b></td>
		<td valign="top"><%=rID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_Title")%>:</b></td>
		<td valign="top"><input type="text" name="eventHeading" value="<%=rTitle%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_URL")%>:</b></td>
		<td valign="top"><input type="text" name="eventURL" value="<%=rURL%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resource_Caption")%>:</b></td>
		<td valign="top"><textarea name="eventAbstract" rows="4" class="formstyleLong" maxlength="255"><%=rCaption%></textarea></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Live")%>:</b></td>
		<td valign="top"><input type="checkbox" name="eventLive" value="1" <%if rLive then Response.write "Checked"%>></td>
	</tr>		
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminResources") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>		
</table>
</form>

<%elseif strAct = "hndEdit" and strEventID<>"" then 
	eventHeading		= SQLEncode(trim(request("eventHeading")))
	eventURL			= SQLEncode(trim(request("eventURL")))
	eventAbstract		= left(SQLEncode(trim(request("eventAbstract"))), 255)
	eventLive			= request("eventLive")
	
	if eventLive <> 1 then
		eventLive = 0
	end if
	
	sql = sql_UpdateResource(eventHeading, eventURL, eventAbstract, eventLive, strEventID)
	Call DoSQL(sql)
	
	session("strMessage") = dictLanguage("Resource_Updated")
	Response.Redirect "default.asp?act=hnd_thnk"
	

  elseif strAct = "delete" and strEventID <> "" then  
		sql = sql_DeleteResource(strEventID)
		Call DoSQL(sql)
		
		session("strMessage") = dictLanguage("Resource_Deleted")
		Response.Redirect "default.asp?act=hnd_thnk"


  elseif strAct = "hnd_thnk" then %>
<p align="Center"><%=session("strMessage")%></p>
	<%session("strMessage") = ""%>

<%end if%>

</td>
</tr>
</table>

<p align="center"><a href=".."><%=dictLanguage("Return_Admin_Home")%></a></p>

<!--#include file="../../includes/main_page_close.asp"-->
