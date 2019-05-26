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
strEmpID	= trim(request("id"))
%>


<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<table border="0" cellspacing="2" cellpadding="2">
<tr><td valign="top">

<table border="0" cellspacing="2" cellpadding="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td colspan="4" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Employee_Types")%></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><a href="employeetypes.asp?act=add"><IMG SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_Emp_Type")%>" border="0"></a></td>
	</tr>
<% strSQL = sql_GetEmployeeTypes()
   Call RunSQL(strSQL, rs)
   do while not rs.eof
		empID = trim(rs("employeetype_id"))
		empName = trim(rs("employeetype")) %>
	<tr>
		<td valign="top"><%=empID%></td>
		<td valign="top" nowrap><%=empName%></td>
		<td valign="top"><a href="employeetypes.asp?act=edit&id=<%=empID%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Emp_Type")%>"></a></td>
		<td valign="top"><a href="employeetypes.asp?act=delete&id=<%=empID%>" onClick="return confirm('<%=dictLanguage("Confirm_Delete_Emp_Type")%>');"><IMG SRC="../images/delete.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Delete_Emp_Type")%>"></a></td>
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

<form name="strForm" id="strForm" action="employeetypes.asp" method="POST">
<input type="hidden" name="act" value="add2">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Add_Emp_Type")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_Type_ID")%>:</b></td>
		<td valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_Type")%>:</b></td>
		<td valign="top"><input type="text" name="empType" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Rate")%>:</b> (<%=dictLanguage("Integer")%>)</td>
		<td valign="top"><input type="text" name="empRate" value="0" class="formstyleShort" maxlength="10"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Cost")%>:</b> (<%=dictLanguage("Integer")%>)</td>
		<td valign="top"><input type="text" name="empCost" value="0" class="formstyleShort" maxLength="10"></td>
	</tr>	
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>			
</table>
</form>

<%elseif strAct = "add2" then
	strType = SQLEncode(request("empType"))
	if strType = "" then
		strType = "New Employee Type"
	end if
	strRate = request("empRate")
	if not isNumeric(strRate) then
		strRate = 0
	else
		strRate = int(strRate)
	end if
	strCost = request("empCost")
	if not isNumeric(strCost) then
		strCost = 0
	else
		strCost = int(strCost)
	end if
	sql = sql_InsertEmployeeType(strType, strRate, strCost)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Added_Emp_Type")
	Response.Redirect "employeetypes.asp?act=hnd_thnk" %>

<%elseif strAct = "edit" and strEmpID <> "" then 
	sql = sql_GetEmployeeTypesByID(strEmpID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		empID			= rs("employeetype_id")
		empType			= trim(rs("employeetype"))
		empRate			= int(trim(rs("rate")))
		empCost			= int(trim(rs("cost")))
	else
		Response.Redirect "employeetypes.asp"
	end if
	rs.close
	set rs = nothing
%>

<form name="strForm" id="strForm" action="employeetypes.asp" method="POST">
<input type="hidden" name="act" value="edit2">
<input type="hidden" name="id" value="<%=empID%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Edit_Emp_Type")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_Type_ID")%>:</b></td>
		<td valign="top"><%=empID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Employee_Type")%>:</b></td>
		<td valign="top"><input type="text" name="empType" value="<%=empType%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Rate")%>:</b> (<%=dictLanguage("Integer")%>)</td>
		<td valign="top"><input type="text" name="empRate" value="<%=empRate%>" class="formstyleShort" maxlength="10"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Cost")%>:</b> (<%=dictLanguage("Integer")%>)</td>
		<td valign="top"><input type="text" name="empCost" value="<%=empCost%>" class="formstyleShort" maxLength="10"></td>
	</tr>		
	<tr><td colspan="2">&nbsp;</td></tr>	
<%		if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%		end if %>			
</table>
</form>

<%elseif strAct = "edit2" and strEmpID <> "" then
	strType = SQLEncode(request("empType"))
	if strType = "" then
		strType = "New Employee Type"
	end if
	strRate = request("empRate")
	if not isNumeric(strRate) then
		strRate = 0
	else
		strRate = int(strRate)
	end if
	strCost = request("empCost")
	if not isNumeric(strCost) then
		strCost = 0
	else
		strCost = int(strCost)
	end if
	sql = sql_UpdateEmployeeType(strEmpID, strType, strRate, strCost)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Updated_Emp_Type")
	Response.Redirect "employeetypes.asp?act=hnd_thnk" %>

<%elseif strAct = "delete" and strEmpID <> "" then 
	sql = sql_DeleteEmployeeType(strEmpID)
	Call DoSQL(sql)
	session("strMessage") = dictLanguage("Deleted_Emp_Type")
	Response.Redirect "employeetypes.asp?act=hnd_thnk"%>

<%elseif strAct = "hnd_thnk" then %>
	<p align="Center"><%=session("strMessage")%></p>
	<%	session("strMessage") = ""%>

<%end if%>
</td>
</tr>
</table>

<p align="center"><a href="employeetypes.asp"><%=dictLanguage("Return_Admin_Home")%></a></p>

<!--#include file="../../includes/main_page_close.asp"-->
