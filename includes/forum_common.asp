<!--#include file="adovbs.inc"-->
<!--#include file="MTS_Forums.asp"-->

<!-- forum code based on Ti Portal's forum sample.  http://www.transworldportal.com -->

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  forum code based on Ti Portal's forum sample.  http://www.transworldportal.com '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub WriteLine(strText)
	Response.Write strText & vbCrLf
End Sub

Function Lineify(strInput)
	Dim strTemp
	strTemp = Server.HTMLEncode(strInput)
	strTemp = Replace(strTemp, "       ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "      ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "     ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "    ", "&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, "   ", "&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strTemp = Replace(strTemp, vbCrLf, "<BR>" & vbCrLf, 1, -1, 1)
	Lineify = strTemp
End Function

Function FormatTimestampDB(dTimestampToFormat)
	' Formats to "m/d/yyyy h:mm:ss AM" format
	' Change as appropriate to match your DB
	Dim strMonth, strDay, strYear
	Dim strHour, strMinute, strSecond
	Dim strAMPM
	
	strMonth = Month(dTimestampToFormat)
	strDay = Day(dTimestampToFormat)
	strYear = Year(dTimestampToFormat)
	'strYear = Right(Year(dTimestampToFormat), 2)

	strHour = Hour(dTimestampToFormat) Mod 12
	If strHour = 0 Then strHour = 12
	
	If Hour(dTimestampToFormat) < 12 Then
		strAMPM = "AM"
	Else
		strAMPM = "PM"
	End If

	strMinute = Minute(dTimestampToFormat)
	If Len(strMinute) = 1 Then strMinute = "0" & strMinute
	
	strSecond = Second(dTimestampToFormat)
	If Len(strSecond) = 1 Then strSecond = "0" & strSecond

	' "d/m/yyyy h:mm:ss AM" for all those who have had problems.
	'FormatTimestampDB = strDay & "/" & strMonth & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM
	
	FormatTimestampDB = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM
End Function

Function FormatTimestampDisplay(dTimestampToFormat)
	' Formats to "m/d/yyyy h:mm:ss AM" format
	' Change as appropriate to match your display wishes
	Dim strMonth, strDay, strYear
	Dim strHour, strMinute, strSecond
	Dim strAMPM
	
	strMonth = Month(dTimestampToFormat)
	strDay = Day(dTimestampToFormat)
	strYear = Year(dTimestampToFormat)
	'strYear = Right(Year(dTimestampToFormat), 2)

	strHour = Hour(dTimestampToFormat) Mod 12
	If strHour = 0 Then strHour = 12
	
	If Hour(dTimestampToFormat) < 12 Then
		strAMPM = "AM"
	Else
		strAMPM = "PM"
	End If

	strMinute = Minute(dTimestampToFormat)
	If Len(strMinute) = 1 Then strMinute = "0" & strMinute
	
	strSecond = Second(dTimestampToFormat)
	If Len(strSecond) = 1 Then strSecond = "0" & strSecond

	' "d/m/yyyy h:mm:ss AM" for all those who have had problems.
	'FormatTimestampDB = strDay & "/" & strMonth & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM

	FormatTimestampDisplay = strMonth & "/" & strDay & "/" & strYear & " " & strHour & ":" & strMinute & ":" & strSecond & " " & strAMPM
End Function

Sub ShowForumLine(iId, sFolderStatus, sName, sDescription, iMessageCount, iPendingCount, blnAdmin)
	Dim strOutput
	if blnAdmin then
		%><a HREF="admin_forum_display.asp?fid=<%=iId%>"><img SRC="<%=gsSiteRoot%>images/forum_folder_<%=sFolderStatus%>.gif" BORDER="0"></a>
		&nbsp;
		<a HREF="admin_forum_display.asp?fid=<%=iId%>"><b><%=sName%></b></a>
		&nbsp;--&nbsp;<%=sDescription%>
		<%	
	else
		%><a HREF="../forum/default.asp?fid=<%=iId%>"><img SRC="<%=gsSiteRoot%>images/forum_folder_<%=sFolderStatus%>.gif" BORDER="0"></a>
		&nbsp;
		<a HREF="../forum/default.asp?fid=<%=iId%>"><b><%=sName%></b></a>
		&nbsp;--&nbsp;<%=sDescription%>
		<%
	end if
	
	if blnAdmin then
		%>&nbsp;<a href="admin_forum_delete.asp?action=verify&amp;fid=<%=iId%>"><%=dictLanguage("Delete_Forum")%></a>
		 &nbsp;<a href="admin_forum_update.asp?action=display&amp;fid=<%=iID%>"><%=dictLanguage("Update_Forum")%></a><%
	end if
	If iMessageCount <> 0 Then
		%>&nbsp;(<%=iMessageCount%>&nbsp;<%=dictLanguage("messages")%>)<%
	End If
	if blnAdmin then
		If iPendingCount <> 0 Then
			%>&nbsp;(<%=iPendingCount%>&nbsp;<%=dictLanguage("messages_pending_approval")%>)<%
		End If	
	end if
	%><br><%
End Sub

Sub ShowPeriodLine(iForumId, strPeriodType, iPeriodsAgo, iMessageCount, iPendingCount, blnAdmin)
	Dim strOutput

	strOutput = strOutput & "<IMG SRC=""" & gsSiteRoot & "images/forum_blank.gif"" BORDER=""0"">"

	strOutput = strOutput & "<A HREF=""../forum/default.asp?fid=" & iForumId & "&pts=" & iPeriodsAgo & """>"
	If strPeriodType = "7days" Then
		Select Case iPeriodsAgo
			Case 0
				strOutput = strOutput & "<B><I>Last 7 Days</I></B></A>"
			Case 1
				strOutput = strOutput & "<B><I>8 to 14 Days Ago</I></B></A>"
			Case 2
				strOutput = strOutput & "<B><I>15 to 21 Days Ago</I></B></A>"
			Case Else
				strOutput = strOutput & "<B><I>" & MonthName(Month(DateAdd("m", -(iPeriodsAgo - 3), Date()))) & "'s Posts</I></B></A>"
		End Select
	Else
		strOutput = strOutput & "<B><I>" & MonthName(Month(DateAdd("m", -iPeriodsAgo, Date()))) & "'s Posts</I></B></A>"
	End If
	If iMessageCount <> 0 Then
		strOutput = strOutput & "&nbsp;("
		strOutput = strOutput & iMessageCount
		strOutput = strOutput & "&nbsp;" & dictLanguage("messages") & ")"
	End If
	if blnAdmin then
		If iPendingCount <> 0 Then
			strOutput = strOutput & "&nbsp;("
			strOutput = strOutput & iPendingCount
			strOutput = strOutput & "&nbsp;" & dictLanguage("messages_pending_approval") & ")"
		End If	
	end if
	WriteLine strOutput & "<BR>"
End Sub

Sub ShowMessageLine(iDepth, iId, sSubject, sAuthor, sEmail, sTime, iReplyCount, sPageType, iActiveMessageId, blnAdmin)
	Dim strOutput
	Dim I

	strOutput = ""
	For I = 0 to iDepth - 1
		strOutput = strOutput & "<IMG SRC=""" & gsSiteRoot & "images/forum_blank.gif"" BORDER=""0"">"
	Next 'I

	If sPageType = "message" Then
		If iActiveMessageId = iId Then
			strOutput = strOutput & "<IMG SRC=""" & gsSiteRoot & "images/forum_arrow.gif"" BORDER=""0"">"
		Else
			strOutput = strOutput & "<IMG SRC=""" & gsSiteRoot & "images/forum_blank.gif"" BORDER=""0"">"
		End If
	Else
		strOutput = strOutput & "<IMG SRC=""" & gsSiteRoot & "images/forum_blank.gif"" BORDER=""0"">"
	End If

	strOutput = strOutput & "<IMG SRC=""" & gsSiteRoot & "images/forum_message.gif"" BORDER=""0"" ALIGN=""absmiddle"">"
	strOutput = strOutput & "&nbsp;"
	if blnAdmin then
		strOutput = strOutput & "<A HREF=""admin_message_display.asp?mid=" & iId & """><B>" & Replace(Server.HTMLEncode(sSubject), " ", "&nbsp;", 1, -1, 1) & "</B></A>"
	else
		strOutput = strOutput & "<A HREF=""../forum/display_message.asp?mid=" & iId & """><B>" & Replace(Server.HTMLEncode(sSubject), " ", "&nbsp;", 1, -1, 1) & "</B></A>"
	end if
	
	strOutput = strOutput & "&nbsp;by&nbsp;"
	strOutput = strOutput & "<I>" & Replace(Server.HTMLEncode(sAuthor), " ", "&nbsp;", 1, -1, 1) & "</I>"
	If sPageType = "message" And sEmail <> "" Then
		strOutput = strOutput & "&nbsp;<A HREF=""mailto:" & Server.HTMLEncode(sEmail) & """><IMG SRC=""" & gsSiteRoot & "images/forum_mail.gif"" BORDER=""0""></A>"
	End If
	strOutput = strOutput & "&nbsp;at&nbsp;" 
	strOutput = strOutput & Replace(sTime, " ", "&nbsp;", 1, -1, 1)
	If sPageType = "forum" Then
		strOutput = strOutput & "&nbsp;("
		strOutput = strOutput & iReplyCount
		strOutput = strOutput & "&nbsp;" & dictLanguage("replies") & ")"
	End If
	strOutput = strOutput & ""
 
	WriteLine strOutput & "<BR>"
End Sub

Function GetNMonthsAgo(iMonthsAgo)
	Dim dPastDate
	
	dPastDate = Date()
	'Response.Write dPastDate & "<BR>"
	dPastDate = DateAdd("m", -iMonthsAgo, dPastDate)
	'Response.Write dPastDate & "<BR>"
	dPastDate = DateAdd("d", -(Day(dPastDate) - 1), dPastDate)
	'Response.Write dPastDate & "<BR>"

	GetNMonthsAgo = CDate(dPastDate)
End Function ' GetNMonthsAgo

Sub SendEmailNotification(iNewMessageId, iThreadId, sPostersEmail)
	' DB object var for email notification
	Dim objNotifyRS
	Dim strSQL

	' Make sure emailing is enabled
	If SEND_EMAIL Then
		' Send Email notify if author has requested it
		' thread_id = 0 -> this is the first post in thread -> no one to notify
		If iThreadId <> 0 Then
			strSQL = "SELECT DISTINCT message_author_email FROM messages WHERE "
			strSQL = strSQL & "message_id <> " & iNewMessageId & " AND "
			strSQL = strSQL & "thread_id = " & iThreadId & " AND "
			strSQL = strSQL & "message_author_notify <> 0 AND "
			strSQL = strSQL & "message_author_email <> '' AND "
			strSQL = strSQL & "message_author_email <> '" & sPostersEmail & "';"

			Set objNotifyRS = GetRecordset(strSQL)
			If Not objNotifyRS.EOF Then
				objNotifyRS.MoveFirst

				Do While Not objNotifyRS.EOF
					SendEmail _
						"Ti Portal Webmaster <webmaster@transworldportal.com>", _
						objNotifyRS.Fields("message_author_email").Value, _
						"A new message has been posted!", _
						"A new message has been posted in a thread you asked us watch for you on Ti Portal's " & _
						"discussion forum.  You can find the forum at http://www.transworldportal.com/forum.  For " & _
						"your convenience, the address of the new message is " & _
						"http://www.transworldportal.com/forum/display_message.asp?mid=" & iNewMessageId & "."
					objNotifyRS.MoveNext
				Loop
			End If
			objNotifyRS.Close
			Set objNotifyRS = Nothing
		End If
	End If
End Sub 'SendEmailNotification

Sub ShowThanks(iNewMessageId, iThreadParent, iForumId, sName, sEmail,blnadmin)
	%>
	<b><%=dictLanguage("Thank_You_Post")%></b><br>
	<br>
	<%if blnAdmin then%>
		<a HREF="admin_message_display.asp?mid=<%=iNewMessageId%>"><i><%=dictLanguage("View_Your_Message")%></i></a><br>
		<%If iThreadParent <> 0 Then%>
			<a HREF="admin_message_display.asp?mid=<%=iThreadParent%>"><i><%=dictLanguage("Back_To_The_Message")%></i></a><br>
		<%End If%>
		<a HREF="../admin.asp?fid=<%=iForumId%>"><i><%=dictLanguage("Back_To_The_Folder")%></i></a><br>
	<%else%>
		<a HREF="display_message.asp?mid=<%=iNewMessageId%>"><i><%=dictLanguage("View_Your_Message")%></i></a><br>
		<%If iThreadParent <> 0 Then%>
			<a HREF="display_message.asp?mid=<%=iThreadParent%>"><i><%=dictLanguage("Back_To_The_Message")%></i></a><br>
		<%End If%>
		<a HREF="default.asp?fid=<%=iForumId%>"><i><%=dictLanguage("Back_To_The_Folder")%></i></a><br>
	<%end if
	End Sub 'ShowThanks

Function InputIsValid(strSituation, iForumId, iThreadId, iThreadParent, iThreadLevel, sName, sSubject, sMessage)
	Dim bEverythingIsCool
	bEverythingIsCool = True
	
	'Validate info
	If IsNumeric(iForumId) Then
		If iForumId <> 0 Then
			iForumId = CLng(iForumId)
		Else
			WriteLine dictLanguage("Error_NoActiveForum") & "<BR>"
			bEverythingIsCool = False
		End If
	Else
		WriteLine dictLanguage("Error_NoActiveForum") & "<BR>"
		bEverythingIsCool = False
	End If

	If IsNumeric(iThreadId) And IsNumeric(iThreadParent) And IsNumeric(iThreadLevel) Then 
		iThreadId = CLng(iThreadId)
		iThreadParent = CLng(iThreadParent)
		If iThreadLevel = 0 Then iThreadLevel = 1
		iThreadLevel = CLng(iThreadLevel)
	Else
		WriteLine dictLanguage("Error_InvalidThread") & "<BR>"
		bEverythingIsCool = False
	End If

	' Do our additional checks if we're about to save!
	If strSituation = "save" Then
		If Len(sName) = 0 Then 
			WriteLine dictLanguate("Error_ForumEmptyName") & "<BR>"
			bEverythingIsCool = False
		End If

		If Len(sSubject) = 0 Then 
			WriteLine dictLanguage("Error_ForumEmptySubject") & "<BR>"
			bEverythingIsCool = False
		End If

		If Len(sMessage) = 0 Then 
			WriteLine dictLanguage("Error_ForumMsgEmpty") & "<BR>"
			bEverythingIsCool = False
		End If
	End If

	InputIsValid = bEverythingIsCool
End Function ' InputIsValid

Sub ShowForm(forum_id, thread_id, thread_parent, thread_level, name, email, subject, message, blnAdmin)
	if blnAdmin then%>
		<form NAME="NewMessage" ACTION="admin_message_post.asp?action=save" METHOD="post">
	<%else%>
		<form NAME="NewMessage" ACTION="post_message.asp?action=save" METHOD="post">
	<%end if
	
	%>
	<input TYPE="hidden" NAME="forum_id" VALUE="<%= forum_id %>">
	<input TYPE="hidden" NAME="thread_id" VALUE="<%= thread_id %>">
	<input TYPE="hidden" NAME="thread_parent" VALUE="<%= thread_parent %>">
	<input TYPE="hidden" NAME="thread_level" VALUE="<%= thread_level %>">

	<table BORDER="0" CELLSPACING="0" CELLPADDING="0">
		<tr>
			<td VALIGN="top"><font class="bolddark"><%=dictLanguage("Name")%>:&nbsp;</td>
			<td><input TYPE="text" NAME="name" MAXLENGTH="50" VALUE="<%= name %>" class="formstyleLong"></td>
		</tr>
		<tr>
			<td VALIGN="top"><font class="bolddark"><%=dictLanguage("Email")%>:&nbsp;</td>
			<td><input TYPE="text" NAME="email" MAXLENGTH="50" VALUE="<%= email %>" class="formstyleLong"> (<%=dictLanguage("Optional")%>)</td>
		</tr>
		<tr>
			<td VALIGN="top"><font class="bolddark"><%=dictLanguage("Subject")%>:&nbsp;</td>
			<td><input TYPE="text" NAME="subject" MAXLENGTH="50" VALUE="<%= subject %>" class="formstyleLong"></td>
		</tr>
		<tr>
			<td VALIGN="top"><font class="bolddark"><%=dictLanguage("Message")%>:&nbsp;</td>
			<td><textarea ROWS="10" NAME="message" WRAP="virtual" class="formstyleLong"><%= message %></textarea></td>
		</tr>
		<% If SEND_EMAIL Then %>
		<tr>
			<td VALIGN="top" ALIGN="right"><input TYPE="checkbox" NAME="notify" VALUE="yes"></td>
			<td COLSPAN="2"><font SIZE="-1" COLOR="#0000FF"><i><%=dictLanguage("Forum_EmailMe")%></i></td>
		</tr>
		<% End If %>
		<% if session("permForumAdd") then%>
		<tr>
			<td COLSPAN="2" ALIGN="center"><input TYPE="reset" VALUE="Reset Form" id="reset1" name="reset1" class="formButton">&nbsp;&nbsp;<input TYPE="submit" VALUE="Post Message" id="submit1" name="submit1" class="formButton"></td>
		</tr>
		<%end if%>
	</table>
	</form>
	<%
	if blnADmin then %>
		<% If thread_parent <> 0 Then %>
			<a HREF="admin_message_display.asp?mid=<%= thread_parent %>"><%=dictLanguage("Back_To_The_Message")%></a><br>
		<% End If %>
		<a HREF="admin.asp?fid=<%= forum_id %>"><%=dictLanguage("Back_To_The_Folder")%></a><br>
	
	<%else%>	
		<% If thread_parent <> 0 Then %>
			<a HREF="display_message.asp?mid=<%= thread_parent %>"><%=dictLanguage("Back_To_The_Message")%></a><br>
		<% End If %>
		<a HREF="default.asp?fid=<%= forum_id %>"><%=dictLanguage("Back_To_The_Folder")%></a><br>
	<%end if
End Sub ' ShowForm

Sub ShowMessageExcerpt(iResultNumber, iId, sSubject, sAuthor, sEmail, sTime, sBody)
	%>
	<dl>
		<dt><font SIZE="-1"><b><%= iResultNumber + 1%>.</b>&nbsp;&nbsp;<a HREF="display_message.asp?mid=<%= iId %>"><font SIZE="-1"><b><%= sSubject %></b></a>&nbsp;&nbsp;<font SIZE="-1"><i>by <%= sAuthor %> on <%= sTime %></i>
		<dd><font SIZE="-1"><%= Server.HTMLEncode(Left(sBody, 255) & "...") %>		
	</dl>
	<%
End Sub 'ShowMessageExcerpt

Sub ShowSearchFormAdvanced(sAuthor, sAuthorType, sEmail, sEmailType, sSubject, sSubjectType, sBody, sBodyType, sStartDate, sEndDate)
	%>
	<form ACTION="search.asp" METHOD="get" ID="advform" NAME="advform">
		<table BORDER="0" CELLSPACING="0" CELLPADDING="2" align="center">
			<!--			<TR>				<TD ALIGN="right">Author:</TD>				<TD><SELECT NAME="a_type"><OPTION>contains</OPTION><OPTION<% If sAuthorType = "is exactly" Then Response.Write " SELECTED" %>>is exactly</OPTION></SELECT></TD>				<TD><INPUT TYPE="text" NAME="a" VALUE="<%= sAuthor %>"></INPUT></TD>			</TR>			<TR>				<TD ALIGN="right">Email:</TD>				<TD><SELECT NAME="e_type"><OPTION>contains</OPTION><OPTION<% If sEmailType = "is exactly" Then Response.Write " SELECTED" %>>is exactly</OPTION></SELECT></TD>				<TD><INPUT TYPE="text" NAME="e" VALUE="<%= sEmail %>"></INPUT></TD>			</TR>			-->
			<tr>
				<td ALIGN="right"><%=dictLanguage("Subject")%>:</td>
				<td><select NAME="s_type" class="formstyleshort"><option value="contains"><%=dictLanguage("contains")%></option><option value="is exactly" <% If sSubjectType="is exactly" Then Response.Write " SELECTED" %>><%=dictLanguage("is_exactly")%></option></select></td>
				<td><input TYPE="text" NAME="s" VALUE="<%= sSubject %>" class="formstyleLong"></td>
			</tr>
			<tr>
				<td ALIGN="right"><%=dictLanguage("Body")%>:</td>
				<td><select NAME="b_type" class="formstyleshort"><option value="contains"><%=dictLanguage("contains")%></option><option value="all words" <% If sBodyType="all words" Then Response.Write " SELECTED" %>><%=dictLanguage("all_words")%></option></select></td>
				<td><input TYPE="text" NAME="b" VALUE="<%= sBody %>" class="formstyleLong"></td>
			</tr>
			<tr>
				<td COLSPAN="2" ALIGN="right"><%=dictLanguage("Posted_Between")%></td>
				<td><input TYPE="text" NAME="startdate" SIZE="10" VALUE="<%= sStartDate %>" class="formstyleShort" TITLE="mm/dd/yy">&nbsp;<%=dictLanguage("and")%>&nbsp;<input TYPE="text" NAME="enddate" SIZE="10" VALUE="<%= sEndDate %>" TITLE="mm/dd/yy" class="formstyleShort">&nbsp;<font COLOR="#999999">(mm/dd/yy)</td>
			</tr>
			<tr>
				<td colspan="3" align="center"><input TYPE="submit" VALUE="Search" onClick="return validateSearch(document.advform.b.value)" name="submit2" class="formButton">&nbsp;<input TYPE="reset" VALUE="Reset Form" id="reset2" name="reset2" class="formButton"></td>
			</tr>
		</table>
	</form>
	<%
End Sub 

Sub ShowChildren(iParentId, iPreviousFilter, iCurrentLevel, iActiveMessageId, objMiscRs, blnAdmin)
	Dim iCurrentLocation
	if isarray(objMiscRs) then
		for intCounter = iCurrentLevel to ubound(objMiscRs)
			if objMiscRs(intcounter)(messages_thread_parent) = iParentId then
				ShowMessageLine iCurrentLevel, _
								objMiscRs(intcounter)(messages_message_id), _
								objMiscRs(intcounter)(messages_message_subject), _
								objMiscRs(intcounter)(messages_message_author), _
								objMiscRs(intcounter)(messages_message_author_email), _
								objMiscRs(intcounter)(messages_message_timestamp), _
								0, "message", iActiveMessageId, blnAdmin

				ShowChildren objMiscRs(intcounter)(messages_message_id), _
							 iParentId, _
							 (iCurrentLevel + 1), _
							 iActiveMessageId, _
							 objMiscRS, blnAdmin
			end if
		next
	end if
	
End Sub ' ShowChildren
%>

<script LANGUAGE="JavaScript">
<!--

  function validateSearch(what)
  {
  
     if (what =='')
     {
     alert('<%=dictLanguage("Error_NoValue")%>');
     return false;
     }
     return true;
  }
    

//-->
</script>

<%
Sub ShowSearchForm()
	%><br><br>
	<form ACTION="<%=gsSiteRoot%>../forum/search.asp" METHOD="get" name="SEARCH">
		<b><%=dictLanguage("Search_Forum_Keyword")%>:</b><br>
		<input TYPE="text" NAME="keyword" class="formstyleMed">
		<input TYPE="submit" VALUE="Search" onClick="return validateSearch(document.SEARCH.keyword.value)" id="submit1" name="submit1" class="formButton">
	</form>
	<%
end sub
%>

