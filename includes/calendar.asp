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
Function GetDaysInMonth(iMonth, iYear)
    Dim dTemp
    dTemp = DateAdd("d", -1, DateSerial(iYear, iMonth + 1, 1))
    GetDaysInMonth = Day(dTemp)
End Function

Function GetWeekdayMonthStartsOn(dAnyDayInTheMonth)
    Dim dTemp
    dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
    GetWeekdayMonthStartsOn = WeekDay(dTemp)
End Function

Function SubtractOneMonth(dDate)
    SubtractOneMonth = DateAdd("m", -1, dDate)
End Function

Function AddOneMonth(dDate)
    AddOneMonth = DateAdd("m", 1, dDate)
End Function

Dim iDIM      ' Days In Month
Dim iDOW      ' Day Of Week that month starts on
Dim iCurrent  ' Variable we use to hold current day of month as we write table
Dim iPosition ' Variable we use to hold current position in table

disDate = date()
curDate = request("date")
if curDate = "" or not isDate(curDate) then
	curDate = date()
end if

curDateShow = Ucase(MonthName(Month(curDate))) & " " & Year(curDate)

iDIM = GetDaysInMonth(Month(curDate), Year(curDate))
iDOW = GetWeekdayMonthStartsOn(curDate)

navDate = CDate(MonthName(month(curdate)) & " 1, " & year(curdate))
navDateLast = CDate(MonthName(Month(navDate)) & " " & iDIM & ", " & Year(navDate)) 
	
lastMonth = SubtractOneMonth(curDate)
nextMonth = AddOneMonth(curDate) %>
	
<table BORDER="0" CELLSPACING="0" CELLPADDING="0"  width="100%" BGCOLOR="#EEEEEE">
	<tr BGCOLOR="#BBBBBB">
		<td ALIGN="left"><a HREF="main.asp?date=<%=lastMonth%>"><img src="<%=gsSiteRoot%>images/left.gif" border="0" WIDTH="25" HEIGHT="18"></a></td>
		<td ALIGN="center" colspan="5" CLASS="homeHeader"><%=curDateShow%></td>
		<td ALIGN="right"><a HREF="main.asp?date=<%=nextMonth%>"><img src="<%=gsSiteRoot%>images/right.gif" border="0" WIDTH="25" HEIGHT="18"></a></td>
	</tr>
	<tr>
		<td ALIGN="center" width="35" class="calendardates"><%=dictLanguage("Sun")%></td>
		<td ALIGN="center" width="35" class="calendardates"><%=dictLanguage("Mon")%></td>
		<td ALIGN="center" width="35" class="calendardates"><%=dictLanguage("Tue")%></td>
		<td ALIGN="center" width="35" class="calendardates"><%=dictLanguage("Wed")%></td>
		<td ALIGN="center" width="35" class="calendardates"><%=dictLanguage("Thu")%></td>
		<td ALIGN="center" width="35" class="calendardates"><%=dictLanguage("Fri")%></td>
		<td ALIGN="center" width="35" class="calendardates"><%=dictLanguage("Sat")%></td>
	</tr>

<%	sql = sql_GetEventsByDateSpan(navDate, navDateLast)
	Call RunSQL(sql, rs)
		
	if gsPTO then
		sql = sql_GetPTOByDateSpan(navDate, navDateLast)
		Call RunSQL(sql, rsPTO)
	end if	
		
	if gsCalendarClients then
		sql = sql_GetClientEventsByDateSpan(navDate, navDateLast)
		Call RunSQL(sql, rsClients)
	end if

	if gsCalendarProjects then
		sql = sql_GetProjectEventsByDateSpan(navDate, navDateLast)
		Call RunSQL(sql, rsProjects)
	end if

	if gsCalendarTasks then
		sql = sql_GetTaskEventsByDateSpan(navDate, navDateLast, session("employee_id"))
		Call RunSQL(sql, rsTasks)
	end if

	' Write spacer cells at beginning of first row if month doesn't start on a Sunday.
	If iDOW <> 1 Then
	    Response.Write vbTab & "<TR>" & vbCrLf
	    iPosition = 1
	    Do While iPosition < iDOW
	        Response.Write vbTab & vbTab & "<TD>&nbsp;</TD>" & vbCrLf
			iPosition = iPosition + 1
	    Loop
	End If
	' Write days of month in proper day slots
	iCurrent = 1
	iPosition = iDOW
	Do While iCurrent <= iDIM
		'-- open the table row --
		If iPosition = 1 Then
			Response.Write(vbTab & "<tr>" & vbCrLf)
		End If
		boolFoundDay = FALSE
		fDate = ""		    
		If Not (rs.BOF and rs.EOF) Then
			rs.MoveFirst
			Do Until rs.EOF
				if  rs("calendar_date")<>"" then
					eDate     = cdate(rs("calendar_date"))
		    		eDay      = Day(eDate)
		    		eMonth    = Month(eDate)
		    		eYear     = Year(eDate)
					If eYear = Year(curDate) and eMonth = Month(curDate) and eDay = iCurrent Then	
						boolFoundDay = TRUE
						fDate = eDate
					End If
				end if
				rs.MoveNext
			Loop
		End If
		if gsPTO then
			If Not (rsPTO.BOF and rsPTO.EOF) Then
				rsPTO.MoveFirst
				Do Until rsPTO.EOF
					if  rsPTO("start_date")<>"" then
						eDate     = cdate(rsPTO("start_date"))
			    		eDay      = Day(eDate)
			    		eMonth    = Month(eDate)
			    		eYear     = Year(eDate)
			    		
			    		gDate     = cdate(rsPTO("end_date"))
			    		gDay      = Day(gDate)
			    		gMonth    = Month(gDate)
			    		gYear     = Year(gDate)
			    		
			    		aDate     = DateValue(iCurrent & "-" & MonthName(Month(curDate)) & "-" & Year(curDate))
			    		
						If eYear = Year(curDate) and eMonth = Month(curDate) and eDay = iCurrent Then	
							boolFoundDay = TRUE
							fDate = eDate
						End If
						if gYear = Year(curDate) and gMonth = Month(curDate) and gDay = iCurrent then
							boolFoundDay = TRUE
							fDate = gDate
						end if
						if DateValue(eDate) <= dateValue(aDate) and DateValue(gDate) >= DateValue(aDate) then
							boolFoundDay = TRUE
							fDate = aDate
						end if
					end if
					rsPTO.MoveNext
				Loop
			End If
		end if		
		if gsCalendarClients then
			If Not (rsClients.BOF and rsClients.EOF) Then
				rsClients.MoveFirst
				Do Until rsClients.EOF
					if  rsClients("client_since")<>"" then
						eDate     = cdate(rsClients("client_since"))
			    		eDay      = Day(eDate)
			    		eMonth    = Month(eDate)
			    		eYear     = Year(eDate)
						If eYear = Year(curDate) and eMonth = Month(curDate) and eDay = iCurrent Then	
							boolFoundDay = TRUE
							fDate = eDate
						End If
					end if
					rsClients.MoveNext
				Loop
			End If
		end if
		if gsCalendarProjects then
			If Not (rsProjects.BOF and rsProjects.EOF) Then
				rsProjects.MoveFirst
				Do Until rsProjects.EOF
					if  rsProjects("start_date")<>"" then
						eDate     = cdate(rsProjects("start_date"))
			    		eDay      = Day(eDate)
			    		eMonth    = Month(eDate)
			    		eYear     = Year(eDate)
						If eYear = Year(curDate) and eMonth = Month(curDate) and eDay = iCurrent Then	
							boolFoundDay = TRUE
							fDate = eDate
						End If
					end if
					if  rsProjects("launch_date")<>"" then
						eDate     = cdate(rsProjects("launch_date"))
			    		eDay      = Day(eDate)
			    		eMonth    = Month(eDate)
			    		eYear     = Year(eDate)
						If eYear = Year(curDate) and eMonth = Month(curDate) and eDay = iCurrent Then	
							boolFoundDay = TRUE
							fDate = eDate
						End If
					end if					
					rsProjects.MoveNext
				Loop
			End If	
		end if
		if gsCalendarTasks then
			If Not (rsTasks.BOF and rsTasks.EOF) Then
				rsTasks.MoveFirst
				Do Until rsTasks.EOF
					if  rsTasks("datedue")<>"" then
						eDate     = cdate(rsTasks("datedue"))
			    		eDay      = Day(eDate)
			    		eMonth    = Month(eDate)
			    		eYear     = Year(eDate)
						If eYear = Year(curDate) and eMonth = Month(curDate) and eDay = iCurrent Then	
							boolFoundDay = TRUE
							fDate = eDate
						End If
					end if
					rsTasks.MoveNext
				Loop
			End If					
		end if
		if boolFoundDay and iCurrent = day(disDate) and Month(curDate) = Month(disDate) and Year(curDate) = Year(disDate) then 
			Response.Write(vbTab & vbTab & "<td align=left valign=top bgcolor='#993300'><a href=""javascript: popupcalendar('" & gsSiteRoot & "calendar/calendar.asp?date=" & fDate & "&view=day');"" class='calendardatestodaylink'>" & iCurrent & "</a>")
		elseif iCurrent = day(disDate) and Month(curDate) = Month(disDate) and Year(curDate) = Year(disDate) then 
			Response.Write(vbTab & vbTab & "<td align=left valign=top class='calendardatestoday'>" & iCurrent & "")
		elseif boolFoundDay then
			Response.Write(vbTab & vbTab & "<td align=left valign=top bgcolor='#993300'><a href=""javascript: popupcalendar('" & gsSiteRoot & "calendar/calendar.asp?date=" & fDate & "&view=day');"" class='calendardateslink'>" & iCurrent & "</a>")							
		else
			Response.Write(vbTab & vbTab & "<td align=left valign=top class='calendardates'>" & iCurrent & "")
		end if
		Response.Write("</td>" & vbCrLf)
		'-- Close the table row --
		If iPosition = 7 Then
			Response.Write vbTab & "</tr>" & vbCrLf
			iPosition = 0
		End If
	    ' Increment variables
	    iCurrent = iCurrent + 1
	    iPosition = iPosition + 1
	Loop
	' Write spacer cells at end of last row if month doesn't end on a Saturday.
	If iPosition <> 1 Then
	    Do While iPosition <= 7
	        Response.Write vbTab & vbTab & "<TD>&nbsp;</TD>" & vbCrLf
	        iPosition = iPosition + 1
	    Loop
	    Response.Write vbTab & "</TR>" & vbCrLf
	End If 
	rs.close
	set rs = nothing 
	if gsPTO then
		rsPTO.close
		set rsPTO = nothing
	end if
	if gsCalendarClients then
		rsClients.close
		set rsClients = nothing
	end if
	if gsCalendarProjects then
		rsProjects.close
		set rsProjects = nothing 
	end if
	if gsCalendarTasks then
		rsTasks.close
		set rsTasks = nothing
	end if%>
</table>
