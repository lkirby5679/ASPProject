<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
<!--#include file="../includes/forum_common.asp"-->

<!-- forum code based on Ti's Portal forum sample.  http://www.transworldportal.com -->

<table width="600" border="0" cellpadding="3" cellspacing="3" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Discussion_Forum")%></td></tr>
	<tr>
		<td colspan="2" >  								
		<%
		'Response.Write "<BR>su_id -"&session("su_id")
		'Response.Write "<BR>user_id -"&session("user_id")

		'Response.Write "<BR>su_moderate -"&Session("su_moderate")
		'Response.Write "<BR>su_section-"&session("su_section")&"<BR>"

		boolAdmin = false
		'if session("user_id") <> ""  then
		'	boolAdmin = true
		'end if

		if session("permAdminForum") then
			boolAdmin = true
		end if

		if boolAdmin then
		  %>
		  <A HREF="<%=gsSiteRoot%>admin/forum/admin.asp"><B><%=dictLanguage("Moderate_Forum")%></B></A><BR><BR>
		  <%
		end if


			Dim objForumRS, objMessageRS
			Dim objForumCountRS, objMessageCountRS
			Dim strThreadList
			Dim iActiveForumId, iActiveForumName
			Dim iForumMessageCount

			Dim iPeriodLooper
			Dim iPeriodToShow
			Dim iPeriodsToGoBack
			Dim strForumBreakdownType

			Dim dStartDate
			Dim dEndDate	

			iActiveForumId = Request.QueryString("fid")
			If IsNumeric(iActiveForumId) Then
				iActiveForumId = CInt(iActiveForumId)
			Else
				iActiveForumId = 0
			End If

			iPeriodToShow = Request.QueryString("pts")
			If IsNumeric(iPeriodToShow) Then
				iPeriodToShow = CInt(iPeriodToShow)
			Else
				iPeriodToShow = 0
			End If

			' Get Forum Info and count of messages in the forum
			
			if session("su_section")<>"" then
			    objForumRS= GetForumsBySect(session("su_section"))
			else
			    objForumRS= GetAllForums()
			end if
			objForumCountRS = GetAllForumsMessageCounts()

			If Isarray(objForumRS) Then	
				For intcounter = 0 To UBound(objforumrs)
					strForumBreakdownType = Trim(LCase(objForumRS(intcounter)(forums_forum_grouping)))
					
					'Get Message Count for Forum
					iForumMessageCount=0
					if isarray(objForumCountRS) then
						for intInnerCounter = 0 to ubound(objForumCountRS)
							If objForumCountRS(intInnerCounter)(0) = objForumRS(intcounter)(forums_forum_id) then
								iForumMessageCount = objForumCountRS(intInnerCounter)(1)
							end if
						next
					end if

					' If active forum -> show messages Else just show forum
					If objForumRS(intcounter)(forums_forum_id) = iActiveForumId Then
						ShowForumLine objForumRS(intcounter)(forums_forum_id), "open", _
									  objForumRS(intcounter)(forums_forum_name), _
									  objForumRS(intcounter)(forums_forum_description), _
									  iForumMessageCount,0,false

						' Show links to previous months
						iPeriodsToGoBack = DateDiff("m", objForumRS(intcounter)(forums_forum_start_date), Now())
						
						' Make adjustments to periods to go back and show for non-monthly breakdown
						Select Case strForumBreakdownType
							Case "7days"
								iPeriodsToGoBack = iPeriodsToGoBack + 3
							Case "monthly"
								' Nothing to do!
							Case Else
								iPeriodsToGoBack = 0
								iPeriodToShow = 0
						End Select
								
						For iPeriodLooper = 0 To iPeriodsToGoBack
							If strForumBreakdownType = "7days" Or strForumBreakdownType = "monthly" Then
								'Do period message count here.
								ShowPeriodLine objForumRS(intcounter)(forums_forum_id), _
											   strForumBreakdownType, iPeriodLooper, 0, 0, false
							End If

							If iPeriodLooper = iPeriodToShow Then
								'Show Root Level Posts for the selected period and their reply count
								Select Case strForumBreakdownType
									Case "7days"
										If iPeriodToShow <= 2 Then
											dStartDate = Date() - (7 * (iPeriodToShow + 1)) + 1
											dEndDate = Date() - (7 * iPeriodToShow) + 1
										Else
											dStartDate = GetNMonthsAgo(iPeriodToShow - 3)
											dEndDate = GetNMonthsAgo(iPeriodToShow - 4)
										End If
									Case "monthly"
										dStartDate = GetNMonthsAgo(iPeriodToShow)
										dEndDate = GetNMonthsAgo(iPeriodToShow - 1)
									Case Else
										dStartDate = objForumRS(intcounter)(forums_forum_start_date)
										dEndDate = Date() + 1
								End Select
								
								objMessageRS = GetForumMessages(iActiveForumId)
		   						'Build the list of root posts we need counts for
								If isarray(objMessageRS) Then
									for int2InnerCount = 0 to  ubound(objMessageRS)
										strThreadList = strThreadList & objMessageRS(int2InnerCount)(Message_thread_id) & ","
									next
									strThreadList = Left(strThreadList, Len(strThreadList) - 1)
								Else
									strThreadList = (0)
								End If
					
								objMessageCountRS = GetMessageReplies(strThreadList)
								
								If isarray(objMessageRS) Then
									for int2InnerCount = 0 to  ubound(objMessageRS)
										'Get Message Replies
										intMessageReplies = 0	
										if isarray(objMessageCountRS) then
											for int3InnerCount = 0 to ubound(objMessageCountRS)
												if objMessageCountRS(int3InnerCount)(0) = objMessageRS(int2innercount)(messages_message_id)then
													intMessageReplies = objMessageCountRS(int3InnerCount)(1) - 1
												end if
											next
										end if

										ShowMessageLine 1, objMessageRS(int2innercount)(messages_message_id), _
														   objMessageRS(int2innercount)(messages_message_subject), _
														   objMessageRS(int2innercount)(messages_message_author), _
														   objMessageRS(int2innercount)(messages_message_author_email), _
														   objMessageRS(int2innercount)(messages_message_timestamp), _
														   intMessageReplies, "forum", 0,false
									next												   
								End If

							End If
						Next 'iPeriodLooper
						'Set active Forum Name for later use in post line
						iActiveForumName = objForumRS(intcounter)(forums_forum_name)
					Else
						ShowForumLine objForumRS(intcounter)(forums_forum_id), "closed", _
									  objForumRS(intcounter)(forums_forum_name), _
									  objForumRS(intcounter)(forums_forum_description), _
									  iForumMessageCount, 0, false
					End If
				next
			Else
				WriteLine dictLanguage("No_Forums_Open") & "<BR>"
			End If


			If iActiveForumId <> 0 Then
			%>
			<BR>
			<A HREF="post_message.asp?fid=<%= iActiveForumId %>"><I><%=dictLanguage("Post_Message_To")%> <B><%= iActiveForumName %></B></I></A><BR>
			<%
			End If

			ShowSearchForm
			%>
	</TD>		
	</TR>	
	<tr><td>&nbsp;</td></tr>
	<tr><td align="right"><small><%=dictLanguage("Forum_Code_Based")%> <a href="http://www.transworldportal.com/" target="_blank"><small>Ti Portal</small></a>'s forum.</small></td></tr>		
</table>

<!-- #include file="../includes/main_page_close.asp"-->
