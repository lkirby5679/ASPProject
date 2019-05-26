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
		<td colspan="4" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Timecard_Types")%></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><a href="timecardtypes.asp?act=add"><IMG SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_Timecard_Type")%>" border="0"></a></td>
	</tr>
<% strSQL = sql_GetTimecardTypes()
   Call RunSQL(strSQL, rs)
   do while not rs.eof
		tcID = trim(rs("timecardtype_id"))
		tcName = trim(rs("timecardtypedescription")) %>
	<tr>
		<td valign="top"><%=tcID%></td>
		<td valign="top" nowrap><%=tcName%></td>
		<td valign="top"><a href="timecardtypes.asp?act=edit&id=<%=tcID%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Timecard_Type")%>"></a></td>
		<td valign="top"><a href="timecardtypes.asp?act=delete&id=<%=tcID%>" onClick="return confirm('<%=dictLanguage("Confirm_Delete_Timecard_Type")%>');"><IMG SRC="../images/delete.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Delete_Timecard_Type")%>"></a></td>
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

<form name="strForm" id="strForm" action="timecardtypes.asp" method="POST">
<input type="hidden" name="act" value="add2">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Add_Timecard_Type")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Timecard_Type_ID")%>:</b></td>
		<td valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Timecard_Type")%>:</b></td>
		<td valign="top"><input type="text" name="tcType" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Active")%>:</b></td>
		<td valign="top"><input type="checkbox" name="tcActive" value="1" checked></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Default_Employee_Type")%>:</b></td>
		<td valign="top"><select name="tcEmpType" class="formstylelong">
							<option value="0"><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee_Type")%></option>
<%	sql = sql_GetEmployeeTypes()
	Call RunSQL(sql, rs)
	while not rs.eof
		Response.Write "<option value=""" & trim(rs("employeetype_id")) & """"
		Response.Write ">" & trim(rs("employeetype")) & "</option>"
		rs.movenext
	wend
	rs.close
	set rs = nothing %>
			</select>
		</td>
	</tr>	
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>			
</table>
</form>

<%elseif strAct = "add2" then
	strType = SQLEncode(request("tcType"))
	if strType = "" then
		strType = "New Timecard Type"
	end if
	strActive = request("tcActive")
	if strActive = 1 or strActive then
		strActive = 1
	else
		strActive = 0
	end if
	strEmpType = request("tcEmpType")
	sql = sql_InsertTimecardType(strType, strActive, strEmpType)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Added_Timecard_Type")
	Response.Redirect "timecardtypes.asp?act=hnd_thnk" %>

<%elseif strAct = "edit" and strID <> "" then 
	sql = sql_GetTimecardTypesByID(strID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		tcID			= rs("timecardtype_id")
		tcType			= trim(rs("timecardtypedescription"))
		tcActive		= rs("active")
		if tcActive = 1 or tcActive = -1 or tcActive then
			tcActive = TRUE
		else
			tcActive = FALSE
		end if
		tcEmpType		= rs("employeetype_id")
	else
		Response.Redirect "timecardtypes.asp"
	end if
	rs.close
	set rs = nothing
%>

<form name="strForm" id="strForm" action="timecardtypes.asp" method="POST">
<input type="hidden" name="act" value="edit2">
<input type="hidden" name="id" value="<%=strID%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Edit_Timecard_Type")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Timecard_Type_ID")%>:</b></td>
		<td valign="top"><%=strID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Timecard_Type")%>:</b></td>
		<td valign="top"><input type="text" name="tcType" value="<%=tcType%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Active")%>:</b></td>
		<td valign="top"><input type="checkbox" name="tcActive" value="1" <%if tcActive then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Default_Employee_Type")%>:</b></td>
		<td valign="top"><select name="tcEmpType" class="formstylelong">
							<option value="0"><%=dictLanguage("Select")%>&nbsp;<%=dictLanguage("Employee_Type")%></option>
<%	sql = sql_GetEmployeeTypes()
	Call RunSQL(sql, rs)
	while not rs.eof
		Response.Write "<option value=""" & trim(rs("employeetype_id")) & """"
		if trim(rs("employeetype_id")) = trim(tcEmpType) then 
			Response.Write " Selected"
		end if 
		Response.Write ">" & trim(rs("employeetype")) & "</option>"
		rs.movenext
	wend
	rs.close
	set rs = nothing %>
			</select>
		</td>
	</tr>		
	<tr><td colspan="2">&nbsp;</td></tr>	
<%		if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%		end if %>			
</table>
</form>

<%elseif strAct = "edit2" and strID <> "" then
	strType = SQLEncode(request("tcType"))
	if strType = "" then
		strType = "New Timecard Type"
	end if
	strActive = request("tcActive")
	if strActive = 1 or strActive then
		strActive = 1 
	else
		strActive = 0
	end if
	strEmpType = request("tcEmpType")
	sql = sql_UpdateTimecardType(strID, strType, strActive, strEmpType)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Updated_Timecard_Type")
	Response.Redirect "timecardtypes.asp?act=hnd_thnk" %>

<%elseif strAct = "delete" and strID <> "" then 
	sql = sql_DeleteTimecardType(strID)
	Call DoSQL(sql)
	session("strMessage") = dictLanguage("Deleted_Timecard_Type")
	Response.Redirect "timecardtypes.asp?act=hnd_thnk"%>

<%elseif strAct = "hnd_thnk" then %>
	<p align="Center"><%=session("strMessage")%></p>
	<%	session("strMessage") = ""%>

<%end if%>
</td>
</tr>
</table>

<p align="center"><a href="timecardtypes.asp"><%=dictLanguage("Return_Admin_Home")%></a></p>


<!--#include file="../../includes/main_page_close.asp"-->
