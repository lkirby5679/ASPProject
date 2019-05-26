<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
<!--#include file="../includes/forum_common.asp"-->

<!-- forum code based on Ti's Portal forum sample.  http://www.transworldportal.com -->

<table border="0" cellpadding="3" cellspacing="3" align="center">
	<tr><td colspan="2" align="center" bgcolor="<%=gsColorHighlight%>" class="homeheader"><%=dictLanguage("Discussion_Forum")%> - <%=dictLanguage("Search_Results")%></td></tr>
	<tr>
		<td colspan="2" >  								
							

			<%
			'Response.Write "<BR>su_id -"&session("su_id")
			'Response.Write "<BR>su_section-"&session("su_section")&"<BR><BR>"

				Dim objSearchRS
				Dim strSQL

				Const PAGE_SIZE = 10
				Const MAX_RECORDS = 200
				Dim iCurrentPage, iTotalPages
				Dim iPageController
				Dim iResultNumber
				
				Dim strSearchType ' basic / advanced / bad
				Dim strKeyword
				Dim strAuthor, strEmail, strSubject, strBody, dStartDate, dEndDate

				Dim strTemp
				Dim I

				' Retrieve parameters
				strKeyword = CStr(Request.QueryString("keyword"))

				strAuthor  = CStr(Request.QueryString("a"))
				strEmail   = CStr(Request.QueryString("e"))
				strSubject = CStr(Request.QueryString("s"))
				strBody    = CStr(Request.QueryString("b"))
				dStartDate = Request.QueryString("startdate")
				dEndDate   = Request.QueryString("enddate")
				
				' Validate dates or eliminate them!
				If IsDate(dStartDate) Then
					dStartDate = CDate(dStartDate)
				Else
					dStartDate = Null
				End If
				If IsDate(dEndDate) Then
					dEndDate = CDate(dEndDate)
				Else
					dEndDate = Null
				End If

				' Determine Search Type
				If strKeyword <> "" Then
					strSearchType = "basic"
				Else
					If strAuthor <> "" Or strEmail <> "" Or strSubject <> "" Or strBody <> "" Or dStartDate <> "" Or dEndDate <> "" Then
						strSearchType = "advanced"
					End If
				End If

				' Branch processing based on search type
				' This gets weird and I was too lazy to comment it all so you're on your own!

				' Retrieve requested page
				iCurrentPage = Request.QueryString("page")

				If iCurrentPage = "" Then
					iCurrentPage = 1
				Else
					iCurrentPage = CInt(iCurrentPage)
				End If

				objSearchRS = getSearch(strSearchType)
				
				If isarray(objSearchRS) Then
					Response.Write dictLanguage("Your_Search_Returned") & ubound(objSearchRS)+1 & dictLanguage("matches") & ". (200 max)<BR>" & vbCrLf & "<BR>" & vbCrLf
						
					' Get total page count
					if ubound(objSearchRS) < PAGE_SIZE then
						iTotalPages = 1
					else	
						iTotalPages = ubound(objSearchRS) / PAGE_SIZE
					end if
					
					'If the request page falls outside the acceptable range,
					'give them the closest match (1 or max)
					If 1 > iCurrentPage Then iCurrentPage = 1
					If iCurrentPage > iTotalPages Then iCurrentPage = iTotalPages

					 'Write page number n of x
					
					%><B><%=dictLanguage("Page")%>: <%=iCurrentPage%> </B> <%=dictLanguage("of")%> <B><%=iTotalPages%></B><BR><%
					
					' Move to proper page
					'objSearchRS.AbsolutePage = iCurrentPage

					for intCounter = 0 to ubound(objSearchRs)
						ShowMessageExcerpt (PAGE_SIZE * (iCurrentPage - 1) + iPageController), _
											objSearchRS(IntCounter)(messages_message_id), _
											objSearchRS(IntCounter)(messages_message_subject), _
											objSearchRS(IntCounter)(messages_message_author), _
											objSearchRS(IntCounter)(messages_message_author_email), _
											objSearchRS(IntCounter)(messages_message_timestamp), _
											objSearchRS(IntCounter)(messages_message_body)

						iPageController = iPageController + 1
					Next
				else
				  Response.Write("<font color=red>" & dictLanguage("No_Messages_Found") & "</font>")	
				End If

				' Do the navigation links if there's more than 1 page!
				If iTotalPages > 1 Then
					' GetQS
					strTemp = Request.QueryString
					strTemp = Replace(strTemp, "page=" & iCurrentPage, "", 1, -1, 1)
					If Left(strTemp, 1) <> "&" Then 
						strTemp = "&" & strTemp
					end if

					%><B><%=dictLanguage("Result_Pages")%>:</B>&nbsp;<%
					' Prev link if not first page
					If iCurrentPage <> 1 Then
						%>
						<A HREF="search.asp?page=<%=iCurrentPage - 1 & strTemp%>">[<B>&lt;&lt;&nbsp;<%=dictLanguage("Prev")%></B>]</A>
						<%
					End If

					' Show number links
					I = 1
					Do While I <= iTotalPages And I <=20
						If I = iCurrentPage Then
							If I < 10 Then 
								Response.Write "&nbsp;<B>" & I & "</B>&nbsp;" 
							end if
						Else
							If I < 10 Then 
								Response.Write "&nbsp;"
							end if
							%>
							<A HREF="search.asp?page=<%=I & strTemp%>"><%=I%></A>&nbsp;
							<%
						End If
						I = I + 1
					Loop
							
					' Next link if not last page
					If iCurrentPage <> iTotalPages Then
						%>
						<A HREF="search.asp?page=<%=iCurrentPage + 1 & strTemp%>">[<B><%=dictLanguage("Next")%> &gt;&gt;</B>]</A>
						<%
					End If
					Response.Write "<BR>" & vbCrLf
				End If

				' Show search form again with appropriate parameters
				Select Case strSearchType
					Case "basic"
						WriteLine "<BR><B>" & dictLanguage("Refine_Your_Search") & ":</B>"
						ShowSearchFormAdvanced "", "", "", "", "", "", strKeyword, "", "", ""
					Case "advanced"
						WriteLine "<BR><B>" & dictLanguage("Refine_Your_Search") & ":</B>"
						ShowSearchFormAdvanced strAuthor, Request.QueryString("a_type"), strEmail, Request.QueryString("e_type"), strSubject, Request.QueryString("s_type"), strBody, Request.QueryString("b_type"), dStartDate, dEndDate
					Case Else
						WriteLine "<B>" & dictLanguage("Advanced_Search") & ":</B>"
						ShowSearchFormAdvanced "", "", "", "", "", "", "", "", "", ""
				End Select



			%>

	</TD>		
	</TR>	
	<tr><td>&nbsp;</td></tr>
	<tr><td align="right"><small><%=dictLanguage("Forum_Code_Based")%> <a href="http://www.transworldportal.com/" target="_blank"><small>Ti Portal</small></a>'s forum.</small></td></tr>			
</table>

<!--#include file="../includes/main_page_close.asp"-->