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
<%		sql = sql_GetTop3Forums()
		Call RunSQL(sql, rs)
		if not rs.eof then
			while not rs.eof
				forumName   = rs("forum_name")
				forumDesc   = rs("forum_description")
				forumID     = rs("forum_id")
				forumDate   = rs("forum_start_date")
				sql = sql_GetForumMessageCountByForum(forumID)
				Call RunSQL(sql, rsCount)
				if not rsCount.eof then
					numMessages = rsCount("numMessages")
					if isNull(numMessages) then
						nummessages = 0
					end if
				else
					numMessages = 0
				end if
				rsCount.close
				set rsCount = nothing 
				
				Response.Write "<a HREF=""" & gsSiteRoot & "../forum/default.asp?fid=" & forumID & """>"
				Response.Write "<img SRC=""" & gsSiteRoot & "images/forum_folder_closed.gif"" BORDER=""0"">"
				Response.Write "&nbsp;</a>"
				Response.Write "<a HREF=""" & gsSiteRoot & "../forum/default.asp?fid=" & forumID & """ class=""smallhome"">"
				Response.Write forumName & "</a><font class=""small"">&nbsp;-&nbsp;"
				'Response.Write forumDesc & "&nbsp;"
				Response.Write "(" & numMessages & "&nbsp;" & dictLanguage("posts") & ")&nbsp;-&nbsp;"
				Response.Write forumDate & "</font><br>"
				
				rs.movenext
			wend
		else 
			Response.Write "There are no folders currently open.<BR>"
		end if
		rs.close
		set rs = nothing %>
