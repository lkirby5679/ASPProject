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

						<table cellpadding="2" cellspacing="2" border="0">
<%			sql = sql_GetTop3News()
			Call runSQL(sql, rs)
			if not rs.eof then 			
				while not rs.eof %>
							<tr>
								<td	colspan="2" class="homecontent">
<%					if trim(rs("news_url")) <> "" then %>
									<a class="home" href="<%=trim(rs("news_url"))%>" target="_blank"><%=rs("news_date")%> - <%=rs("news_heading")%></a> <font class="small">(<%=rs("news_source")%>)</font>
<%					else %>
									<a class="home" href="javascript: popup('<%=gsSiteRoot%>news/event.asp?id=<%=rs("id")%>');"><%=rs("news_date")%> - <%=rs("news_heading")%></a>
<%					end if %>		
								</td>
							</tr>
<%					if trim(rs("news_url")) <> "" then 
						'do nothing
					else %>							
							<tr>
								<td>&nbsp;</td>
								<td class="homecontent"><font class="small"><%=left(rs("news_abstract"),100)%>...&nbsp;<a href="javascript: popup('<%=gsSiteRoot%>news/event.asp?id=<%=rs("id")%>');" class="more"><%=dictLanguage("more")%> >></a></font></td>
							</tr>
<%					end if
					rs.movenext 
				wend 
			rs.close
			set rs = nothing 
			end if %>
						</table>