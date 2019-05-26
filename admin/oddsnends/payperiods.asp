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
		<td colspan="4" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Employee_Pay_Periods")%></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><a href="payperiods.asp?act=add"><IMG SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_Pay_Period")%>" border="0"></a></td>
	</tr>
<% strSQL = sql_GetPayPeriods()
   Call RunSQL(strSQL, rs)
   do while not rs.eof
		pID = rs("id")
		pStartDate = rs("startdate") 
		pEndDate = rs("enddate")%>
	<tr>
		<td valign="top"><%=pID%></td>
		<td valign="top" nowrap><%=pStartDate%> - <%=pEndDate%></td>
		<td valign="top"><a href="payperiods.asp?act=edit&id=<%=pID%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Pay_Period")%>"></a></td>
		<td valign="top"><a href="payperiods.asp?act=delete&id=<%=pID%>" onClick="return confirm('<%=dictLanguage("Confirm_Delete_Pay_Period")%>');"><IMG SRC="../images/delete.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Delete_Pay_Period")%>"></a></td>
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

<form name="strForm" id="strForm" action="payperiods.asp" method="POST">
<input type="hidden" name="act" value="add2">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Add_Pay_Period")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Pay_Period_ID")%>:</b></td>
		<td valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Start_Date")%>:</b></td>
		<td valign="top"><input type="text" name="pStartDate" value="<%=date()%>" class="formstyleShort" maxlength="10"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("End_Date")%>:</b></td>
		<td valign="top"><input type="text" name="pEndDate" value="<%=date()%>" class="formstyleShort" maxlength="10"></td>
	</tr>	
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>			
</table>
</form>

<%elseif strAct = "add2" then
	pStartDate = request("pStartDate")
	if not isDate(pStartDate) then
		pStartDate = date()
	end if
	pEndDate = request("pEndDate")
	if not isDate(pEndDate) then
		pEndDate = date()
	end if
	sql = sql_InsertPayPeriod(pStartDate, pEndDate)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Added_Pay_Period")
	Response.Redirect "payperiods.asp?act=hnd_thnk" %>

<%elseif strAct = "edit" and strID <> "" then 
	sql = sql_GetPayPeriodsByID(strID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		pID			= rs("id")
		pStartDate	= rs("startdate")
		pEndDate    = rs("enddate")
	else
		Response.Redirect "payperiods.asp"
	end if
	rs.close
	set rs = nothing
%>

<form name="strForm" id="strForm" action="payperiods.asp" method="POST">
<input type="hidden" name="act" value="edit2">
<input type="hidden" name="id" value="<%=pID%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Edit_Pay_Period")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Pay_Period_ID")%>:</b></td>
		<td valign="top"><%=pID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Start_Date")%>:</b></td>
		<td valign="top"><input type="text" name="pStartDate" value="<%=pStartDate%>" class="formstyleShort" maxlength="10"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("End_Date")%>:</b></td>
		<td valign="top"><input type="text" name="pEndDate" value="<%=pEndDate%>" class="formstyleShort" maxlength="10"></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%		if session("permAdminEmployees") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%		end if %>			
</table>
</form>

<%elseif strAct = "edit2" and strID <> "" then
	pStartDate = request("pStartDate")
	if not isDate(pStartDate) then
		pStartDate = date()
	end if
	pEndDate = request("pEndDate")
	if not isDate(pEndDate) then
		pEndDate = date()
	end if
	sql = sql_UpdatePayPeriod(strID, pStartDate, pEndDate)
	Call DoSQL(sql)
	session("strMessage")= dictLanguage("Updated_Pay_Period")
	Response.Redirect "payperiods.asp?act=hnd_thnk" %>

<%elseif strAct = "delete" and strID <> "" then 
	sql = sql_DeletePayPeriod(strID)
	Call DoSQL(sql)
	session("strMessage") = dictLanguage("Deleted_Pay_Period")
	Response.Redirect "payperiods.asp?act=hnd_thnk"%>

<%elseif strAct = "hnd_thnk" then %>
	<p align="Center"><%=session("strMessage")%></p>
	<%	session("strMessage") = ""%>

<%end if%>
</td>
</tr>
</table>

<p align="center"><a href="payperiods.asp"><%=dictLanguage("Return_Admin_Home")%></a></p>

<!--#include file="../../includes/main_page_close.asp"-->
