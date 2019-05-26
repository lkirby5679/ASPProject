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
strStartDate = trim(request("startdate"))
strEndDate = trim(request("enddate"))

if strStartDate <> "" then
	if not isDate(strStartDate) then
		strStartDate = CDate("January 1, 2001")
	end if
else
	strStartDate = CDate("January 1, 2001")
end if
if strEndDate <> "" then
	if not isDate(strEndDate) then
		strEndDate = date()
	end if
else
	strEndDate = date()
end if
%>


<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<table border="0" cellspacing="2" cellpadding="2">
<tr><td valign="top">

<table border="0" cellspacing="2" cellpadding="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td colspan="4" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("News")%></td>
	</tr>
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=add&startdate=<%=strStartDate%>&enddate=<%=strEndDate%>"><IMG SRC="../images/new.gif" height="10" width="10" alt="<%=dictLanguage("Add_New_Event")%>" border="0"></a></td>
	</tr>
<% strSQL = sql_GetNewsByDateSpanAdmin(strStartDate, strEndDate)
   Call RunSQL(strSQL, rs)
   do while not rs.eof
		eventID = trim(rs("id"))
		calendarDate = trim(rs("news_date"))
		calendarHeading = left(trim(rs("news_heading")),40) 
		newsURL = trim(rs("news_url"))
		newsSource = trim(rs("news_source"))%>
	<tr>
		<td valign="top"><%=calendarDate%></td>
		<td valign="top" nowrap>
<%		if newsURL <> "" then %>
			<a href="<%=newsURL%>" target="_blank"><%=calendarHeading%></a> (<%=newsSource%>)
<%		else %>
			<a href="javascript: popup('<%=gsSiteRoot%>news/event.asp?ID=<%=eventID%>');"><%=calendarHeading%></a>
<%		end if %>		
		</td>
		<td valign="top"><a href="default.asp?act=edit&id=<%=eventID%>&startdate=<%=strStartDate%>&enddate=<%=strEndDate%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Edit_Event")%>"></a></td>
		<td valign="top"><a href="default.asp?act=delete&id=<%=eventID%>&startdate=<%=strStartDate%>&enddate=<%=strEndDate%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Delete_Event")%>');"><IMG SRC="../images/delete.gif" height="10" width="10" border="0" alt="<%=dictLanguage("Delete_Event")%>"></a></td>
	</tr>
<%		rs.movenext
	loop 
	rs.close
	set rs = nothing %>
</table>

<br><br>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<table cellpadding="4" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Search_News")%></b></td></tr>
	<tr>
		<td valign="top"><%=dictLanguage("Start_Date")%>:</td>
		<td valign="top"><input type="text" name="startdate" value="<%=strStartDate%>" class="formstyleshort" maxlength="10"></td>
	</tr>
	<tr>
		<td valign="top"><%=dictLanguage("End_Date")%>:</td>
		<td valign="top"><input type="text" name="enddate" value="<%=strEndDate%>" class="formstyleshort" maxlength="10"></td>
	</tr>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
</table>
</form>

</td><td align="center" width="100%" valign="top">

<table border="0" cellspacing="2" cellpadding="2" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td align="Center" class="homeheader"><%=dictLanguage("Workspace")%></td>
	</tr>
</table>

<%if strAct = "add" then %>

<form name="strForm" id="strForm" action="default.asp?startdate=<%=strStartDate%>&enddate=<%=strEndDate%>" method="POST">
<input type="hidden" name="act" value="hndAdd">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Add_New_Event")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("News_ID")%>:</b></td>
		<td valign="top">&nbsp;</td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("News_Date")%>:</b></td>
		<td valign="top"><input type="text" name="eventDate" value="<%=date%>" class="formstyleShort" maxlength="10"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Heading")%>:</b></td>
		<td valign="top"><input type="text" name="eventHeading" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("URL")%>:</b></td>
		<td valign="top"><input type="text" name="eventURL" value="" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Source")%>:</b></td>
		<td valign="top"><input type="text" name="eventSource" value="" class="formstyleLong" maxlength="100"></td>
	</tr>		
	<tr>
		<td valign="top"><b><%=dictLanguage("Abstract")%>:</b></td>
		<td valign="top"><textarea name="eventAbstract" rows="4" class="formstyleLong" maxlength="255"></textarea></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Details")%>:</b></td>
		<td valign="top"><textarea name="eventContent" rows="8" class="formstyleLong"></textarea></td>
	</tr>	

	<tr>
		<td valign="top"><b><%=dictLanguage("Live")%>:</b></td>
		<td valign="top"><input type="checkbox" name="eventLive" value="1" Checked></td>
	</tr>		
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminNews") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>			
</table>
</form>

<%elseif strAct = "hndAdd" then 
	eventDateEntered	= date()
	eventDate			= trim(request("eventDate"))
	eventHeading		= SQLEncode(trim(request("eventHeading")))
	eventURL			= SQLEncode(trim(request("eventURL")))
	eventSource			= SQLEncode(trim(request("eventSource")))
	eventAbstract		= left(SQLEncode(trim(request("eventAbstract"))), 255)
	eventContent		= SQLEncode(trim(request("eventContent")))
	eventLive			= request("eventLive")
	eventDeleted		= 0
	eventEnteredBy		= session("employee_id")
	
	if eventDate = "" then 
		eventDate = date()
	end if
	if not isDate(eventDate) then
		eventDate = date()
	end if
	if eventLive <> 1 then
		eventLive = 0
	end if
	
	sql = sql_InsertNewsEvent(eventDateEntered, eventDate, eventHeading, eventURL, _
		eventSource, eventAbstract, eventContent, eventLive, eventDeleted, eventEnteredBy)
	Call DoSQL(sql)
	
	session("strMessage") = dictLanguage("Event_Added")
	Response.Redirect "default.asp?act=hnd_thnk&startdate=" & strStartDate & "&enddate=" & strEndDate & ""
	

  elseif strAct = "edit" and strEventID <> "" then 
	sql = sql_GetNewsByID(strEventID)
	Call RunSQL(sql, rs)
	if not rs.eof then
		eventID			= rs("id")
		eventDate		= trim(rs("news_date"))
		eventHeading	= trim(rs("news_heading"))
		eventURL		= trim(rs("news_url"))
		eventSource		= trim(rs("news_source"))
		eventAbstract   = trim(rs("news_abstract"))
		eventContent	= trim(rs("news_content"))
		eventLive 		= rs("news_live")
		if eventLive = 1 or eventLive = -1 or eventLive then
			eventLive = TRUE
		else
			eventLive = FALSE
		end if
	else
		Response.Redirect "default.asp"
	end if
	rs.close
	set rs = nothing
%>

<form name="strForm" id="strForm" action="default.asp?startdate=<%=strStartDate%>&enddate=<%=strEndDate%>" method="POST">
<input type="hidden" name="act" value="hndEdit">
<input type="hidden" name="id" value="<%=strEventID%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="2" align="center"><b><%=dictLanguage("Edit_Event")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("News_ID")%>:</b></td>
		<td valign="top"><%=eventID%></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("News_Date")%>:</b></td>
		<td valign="top"><input type="text" name="eventDate" value="<%=eventDate%>" class="formstyleShort" maxlength="10"></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Heading")%>:</b></td>
		<td valign="top"><input type="text" name="eventHeading" value="<%=eventHeading%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("URL")%>:</b></td>
		<td valign="top"><input type="text" name="eventURL" value="<%=eventURL%>" class="formstyleLong" maxlength="100"></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Source")%>:</b></td>
		<td valign="top"><input type="text" name="eventSource" value="<%=eventSource%>" class="formstyleLong" maxlength="100"></td>
	</tr>			
	<tr>
		<td valign="top"><b><%=dictLanguage("Abstract")%>:</b></td>
		<td valign="top"><textarea name="eventAbstract" rows="4" class="formstyleLong" maxlength="255"><%=eventAbstract%></textarea></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Details")%>:</b></td>
		<td valign="top"><textarea name="eventContent" rows="8" class="formstyleLong"><%=eventContent%></textarea></td>
	</tr>	

	<tr>
		<td valign="top"><b><%=dictLanguage("Live")%>:</b></td>
		<td valign="top"><input type="checkbox" name="eventLive" value="1" <%if eventLive then Response.write "Checked"%>></td>
	</tr>		
	<tr><td colspan="2">&nbsp;</td></tr>	
<%	if session("permAdminCalendar") then %>
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%	end if %>		
</table>
</form>

<%elseif strAct = "hndEdit" and strEventID<>"" then 
	eventDate			= trim(request("eventDate"))
	eventHeading		= SQLEncode(trim(request("eventHeading")))
	eventURL			= SQLEncode(trim(request("eventURL")))
	eventSource			= SQLEncode(trim(request("eventSource")))
	eventAbstract		= left(SQLEncode(trim(request("eventAbstract"))), 255)
	eventContent		= SQLEncode(trim(request("eventContent")))
	eventLive			= request("eventLive")
	
	if eventDate = "" then 
		eventDate = date()
	end if
	if not isDate(eventDate) then
		eventDate = date()
	end if
	if eventLive <> 1 then
		eventLive = 0
	end if
	
	sql = sql_UpdateNewsEvent(eventDate, eventHeading, eventURL, eventSource, _
			eventAbstract, eventContent, eventLive, strEventID)
	Call DoSQL(sql)
	
	session("strMessage") = dictLanguage("Event_Updated")
	Response.Redirect "default.asp?act=hnd_thnk&startdate=" & strStartDate & "&enddate=" & strEndDate & ""
	

  elseif strAct = "delete" and strEventID <> "" then  
		sql = sql_DeleteNewsEvent(strEventID)
		Call DoSQL(sql)
		
		session("strMessage") = dictLanguage("Event_Deleted")
		Response.Redirect "default.asp?act=hnd_thnk&startdate=" & strStartDate & "&enddate=" & strEndDate & ""


  elseif strAct = "hnd_thnk" then %>
<p align="Center"><%=session("strMessage")%></p>
	<%session("strMessage") = ""%>

<%end if%>

</td>
</tr>
</table>

<p align="center"><a href=".."><%=dictLanguage("Return_Admin_Home")%></a></p>

<!--#include file="../../includes/main_page_close.asp"-->
