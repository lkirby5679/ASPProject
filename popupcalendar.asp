<!-- #include file="includes/SiteConfig.asp"-->
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
	Dim sShowDate
	Dim sCallBack
	Dim sCallBackField

	Dim sMonth(12)
	Dim sDay(7)
	Dim sDayName(7)

	sDayName(1) = WeekdayName(1)
	sDayName(2) = WeekdayName(2)
	sDayName(3) = WeekdayName(3)
	sDayName(4) = WeekdayName(4)
	sDayName(5) = WeekdayName(5)
	sDayName(6) = WeekdayName(6)
	sDayName(7) = WeekdayName(7)

	sDay(1) = left(sDayName(1),1)
	sDay(2) = left(sDayName(2),1)
	sDay(3) = left(sDayName(3),1)
	sDay(4) = left(sDayName(4),1)
	sDay(5) = left(sDayName(5),1)
	sDay(6) = left(sDayName(6),1)
	sDay(7) = left(sDayName(7),1)

	sMonth(1)  = MonthName(1)
	sMonth(2)  = MonthName(2)
	sMonth(3)  = MonthName(3)
	sMonth(4)  = MonthName(4)
	sMonth(5)  = MonthName(5)
	sMonth(6)  = MonthName(6)
	sMonth(7)  = MonthName(7)
	sMonth(8)  = MonthName(8)
	sMonth(9)  = MonthName(9)
	sMonth(10) = MonthName(10)
	sMonth(11) = MonthName(11)
	sMonth(12) = MonthName(12)
	
	dim dAdjDateTime
	dAdjDateTime = DateAdd("h",Application("TimeZoneOffset"),Now)
	
	'Get the Date the calendar should display
	sShowDate = Request.QueryString("Date")
	if Len(sShowDate) = 0 then sShowDate = dAdjDateTime
	
	'Get the function to call when the user clicks selects the date
	sCallback = Request.QueryString("CallBack")
	sCallbackField = Request.QueryString("CallBackField")
	
	DisplayCalender sShowDate, sCallBack, sCallBackField
	
	Sub DisplayCalender(vDate,sCallBack,sCallBackField)%>
		
		<%
			vPrevDate = DateAdd("m",-1,vDate)
			vNextDate = DateAdd("m", 1,vDate)
			vPrevDate = CDate(MonthName(Month(vPrevDate)) & " 1, " & Year(vPrevDate))
			vNextDate = CDate(MonthName(Month(vNextDate)) & " 1, " & Year(vNextDate))
		%>

		<html>
			<head>
				<title>Calendar</title>
				<script language="javascript">
					function doNothing(){;}
				</script>
			</head>
		</html>
		<body>
		<table ALIGN="CENTER" WIDTH="140" BORDER="0" CELLPADDING="2" CELLSPACING="1" BGCOLOR="#FFFFFF">
			
			<tr>
				<td BGCOLOR="<%=gsColorHighlight%>" ALIGN="CENTER" COLSPAN="1"><a HREF="PopupCalendar.asp?date=<%=vPrevDate%>&amp;Callback=<%=sCallBack%>&amp;CallbackField=<%=sCallBackField%>"><img SRC="<%=gsSiteRoot%>images/left.gif" BORDER="0"></a></td>
				<td BGCOLOR="<%=gsColorHighlight%>" ALIGN="CENTER" COLSPAN="5"><b><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sMonth(Month(vDate))%>&nbsp;<%=CStr(Year(vDate))%></font></b></td>
				<td BGCOLOR="<%=gsColorHighlight%>" ALIGN="CENTER" COLSPAN="1"><a HREF="PopupCalendar.asp?date=<%=vNextDate%>&amp;Callback=<%=sCallBack%>&amp;CallbackField=<%=sCallBackField%>"><img SRC="<%=gsSiteRoot%>images/right.gif" BORDER="0"></a></td>
			</tr>
			<tr>
				<td BGCOLOR="#EEEEEE" ALIGN="CENTER"><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sDay(1)%></font></td>
				<td BGCOLOR="#EEEEEE" ALIGN="CENTER"><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sDay(2)%></font></td>
				<td BGCOLOR="#EEEEEE" ALIGN="CENTER"><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sDay(3)%></font></td>
				<td BGCOLOR="#EEEEEE" ALIGN="CENTER"><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sDay(4)%></font></td>
				<td BGCOLOR="#EEEEEE" ALIGN="CENTER"><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sDay(5)%></font></td>
				<td BGCOLOR="#EEEEEE" ALIGN="CENTER"><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sDay(6)%></font></td>
				<td BGCOLOR="#EEEEEE" ALIGN="CENTER"><font FACE="arial" SIZE="1" COLOR="<%=gsColorBackground%>"><%=sDay(7)%></font></td>
			</tr>
			<%
			Dim dDate, iPos, iRemainder, i, dPMDate
			'go to the first of the month
			dDate = CDate(MonthName(Month(vDate)) & " 1, " & CStr(Year(vDate)))
			dNow = CDate(dAdjZoneDateTime)
			iPos = WeekDay(dDate) - 1

			Response.Write "<TR>" 
			'If iPos > 0 Then Response.Write "<TD BGCOLOR=""#DADBD6"" COLSPAN=" & CStr(iPos) & " >&nbsp;</TD>"
			If iPos > 0 Then
				For i = 0 to iPos-1
					dPMDate = DateAdd("d",-1*(iPos-i),dDate)
					Response.Write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER WIDTH=20><FONT FACE=arial SIZE=1><A HREF='doNothing()' onclick='javascript:window.opener." & sCallBack & "(""" & dPMDate & """, """ & sCallBackField & """); window.close();' style=""color:#AAAAAA"">" & Day(dPMDate) & "</A></FONT></TD>"
				Next
			End If
			
			Do While Month(dDate) = Month(vDate)
				If iPos Mod 7 = 0 And iPos > 0 Then Response.Write "</TR><TR>"
				Response.Write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER WIDTH=20><FONT FACE=arial SIZE=1><A HREF='doNothing()' onclick='javascript:window.opener." & sCallBack & "(""" & CDate(MonthName(Month(dDate)) & " " & Day(dDate) & ", " & Year(dDate)) & """, """ & sCallBackField & """); window.close();' style=""color:" & gsColorBackground & """>" & Day(dDate) & "</A></FONT></TD>"
				iPos = iPos + 1
				'add a day to this date
				dDate = dDate + 1
			Loop
			
			iRemainder = iPos Mod 7
			If iRemainder > 0 Then
				For iPos = 1 To 7-iRemainder
					Response.Write "<TD BGCOLOR=""#EEEEEE"" ALIGN=CENTER WIDTH=20><FONT FACE=arial SIZE=1><A HREF='doNothing()' onclick='javascript:window.opener." & sCallBack & "(""" & CDate(DateAdd("d",iPos-1,dDate)) & """, """ & sCallBackField & """); window.close();' style=""color:#AAAAAA"">" & iPos & "</A></FONT></TD>"
				Next
			End If
			Response.Write "</TR></TABLE>"	
	End Sub
%>

