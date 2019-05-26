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
<!--#include file="../includes/main_page_open.asp"-->

<%Call ConfirmPermission("permPTOAdmin", "")%>

<% start_date = request.form("start_date")
   end_date = request.form("end_date")
%>

<table width="100%" border=0 cellpadding=2 cellspacing=2 style="border-style: solid; border-width: 2; border-color: <%=gsColorBlack%>">
	<tr>
		<td valign=top colspan=8>
			<b class="homeheader"><%=dictLanguage("PTO_Requests_By_Date")%></b><br>
			<%=now()%><br>
		</td>
	</tr>
<%

if start_date = "" and end_date = "" then
	date_query = ""
elseif start_date <>"" and end_date<>"" then
	date_query =  " and start_date >= " & DB_DATEDELIMITER & start_date & DB_DATEDELIMITER & " and start_date <= " & DB_DATEDELIMITER & end_date & DB_DATEDELIMITER & ""
elseif start_date <>"" then
	date_query = " and start_date >= " & DB_DATEDELIMITER & start_date & DB_DATEDELIMITER & ""
else 
	date_query = " and start_date <= " & DB_DATEDELIMITER & end_date & DB_DATEDELIMITER & ""
end if

sql = sql_GetEmployeePTOByDateRange(date_query)
set rs = Conn.execute(sql)
intRowCounter = 0
do while not rs.eof
	intRowCounter = intRowCounter + 1
	New_EmployeeName = rs("EmployeeName")
	if New_EmployeeName <> Old_EmployeeName then %>

	<tr>
		<td colspan=8><br><b><%=rs("EmployeeName")%></b></td>
	</tr>
	<tr bgcolor="<%=gsColorHighlight%>">
		<td valign=top class="columnheader">&nbsp;</td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("From")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("To")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Total_Hours")%></td>
		<td valign=top class="columnheader" nowrap><%=dictlanguage("Paid")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Approved")%></td>
		<td valign=top class="columnheader" nowrap><%=dictLanguage("Excused")%></td>
		<td valign="top" class="columnheader" width="100%"><%=dictLanguage("Reason")%></td>
	</tr>

<%	end if%>

	<tr <%If intRowcounter MOD 2 = 1 then %>bgcolor="<%=gsColorWhite%>"<%Else%>bgcolor="#FFFFFF"<%End If%>>
		<td valign="top" nowrap><a href="ptoapproval.asp?pto_ID=<%=rs("pto_ID")%>"><%=dictLanguage("Edit")%></a></td>
		<td valign="top" nowrap><%=rs("Start_Date")%> - <%=timevalue(rs("Start_Time"))%></td>
		<td valign=top nowrap><%=rs("End_Date")%> - <%=timevalue(rs("End_Time"))%></td>
		<td valign=top align="right" nowrap><%=rs("Total_Hours")%></td>
		<td valign=top nowrap><%=rs("Paid")%></td>
		<td valign=top nowrap><%=rs("Approved")%></td>
		<td valign=top nowrap><%=rs("Excused")%></td>
		<td valign=top><%=rs("Reason")%></td>
	</tr>

<%
	Old_EmployeeName = New_EmployeeName
	rs.movenext
loop
rs.close
set rs = nothing
%>
</table>

<p align="center"><a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a></p>

<br>

<!--#include file="../includes/main_page_close.asp"-->

