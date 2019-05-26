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

<%
sql = sql_GetTimecardViewByEmployeeID(session("employee_id"), DateAdd("d",Date(),-30))
'response.write sql
Call RunSQL(sql, rsViewTimeCard)

'********************* This section counts the total hours for today only **************************
sql = sql_GetTimecardHoursTodayByEmployeeID(session("employee_id"))
Call RunSQL(sql, rsToday)
If not rsToday.EOF Then
	if rsToday("timecharged")<>"" then
		TimeCount1 = rsToday("timecharged")
	else
		TimeCount1 = 0
	End if
	if rsToday("numtimecards")<>"" then
		RecordCount1 = rsToday("numtimecards")
	else
		RecordCount1 = 0
	end if
End If
rsToday.close
Set rsToday = nothing
%>

<!--#include file="../includes/main_page_open.asp"-->

<table border="0" cellpadding=2 cellspacing=2 align="center">
	<tr>
		<td valign=top colspan=6>
			<b class="homeheader"><%=dictLanguage("View_My_Timecards")%>&nbsp; (<%=dictLanguage("Last_30_Days")%>) -- <%=session("userid")%></b><br>
			<font class="small"><%= "(" & dictLanguage("Total_Hours_Today") & ": " & TimeCount1 & ")" %></font>
			
		</td>
	</tr>
	<tr bgcolor="<%=gsColorHighlight%>">
		<td valign=top class="columnHeader" nowrap><%=dictLanguage("Date")%></td>
		<td valign=top class="columnHeader" nowrap><%=dictLanguage("Client")%></td>
		<td valign=top class="columnHeader" nowrap><%=dictLanguage("Project")%></td>
		<td valign=top class="columnHeader" nowrap><%=dictLanguage("Hours")%></td>
		<td valign=top class="columnHeader" width="100%"><%=dictLanguage("Description")%></td>
	</tr>
<%
intRowCounter = 0
Do While Not rsViewTimeCard.EOF
	intRowCounter = intRowCounter + 1 %>

	<tr <%If intRowcounter MOD 2 = 1 then %>bgcolor="<%=gsColorWhite%>"<%Else%>bgcolor="#ffffff"<%End If%>>
		<td valign=top class="small">
			<a href="timecard-edit.asp?timecard_id=<%=rsViewTimeCard("TimeCard_ID")%>"><%=rsViewTimeCard("DateWorked")%></a>
		</td>
		<td valign=top class="small"><%=rsViewTimeCard("Client_Name")%></td>
		<td valign=top class="small">
<%	if len(rsViewTimeCard("Description"))>0 then
		response.write rsViewTimeCard("Description")
	else
		response.write dictLanguage("None")
	end if %></td>
		<td valign=top class="small"><%=rsViewTimeCard("TimeAmount")%></td>
		<td valign=top class="small"><%=rsViewTimeCard("WorkDescription")%></td>
	</tr>
<%
	rsViewTimeCard.MoveNext
Loop 
rsViewTimeCard.close
set rsViewTimeCard = nothing %>
	<tr bgcolor="<%=gsColorHighlight%>">
		<td valign=top></td>
		<td valign=top></td>
		<td valign=top></td>
		<td valign=top></td>
		<td valign=top></td>
	</tr>
</table>
<p align="center">
<a href="../main.asp"><%=dictLanguage("Return_Business_Console")%></a><br>
</p>

<!--#include file="../includes/main_page_close.asp"-->