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

<%intNavNumber = 0%>
						<table border="0" cellpadding="2" cellspacing="2" width="100%">
							<tr>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>main.asp"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs();" onMouseOut="javascript: rollOffNav(0);"<%end if%>><%=dictLanguage("Home")%></div></a>
								</td>
<%intNavNumber = intNavNumber + 1
  if session("permAdmin") then %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>admin/"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs();" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Admin")%></div></a>
								</td>
<%	intNavNumber = intNavNumber + 1
  end if
  if gsTimecards then %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>timecards/"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs(); showNavSub('Timecards',event);" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Timecards")%></div></a>
								</td>
<%	intNavNumber = intNavNumber + 1
  end if 
  if gsTasks then %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>tasks/"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs(); showNavSub('Tasks',event);" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Tasks")%></div></a>
								</td>
<%	intNavNumber = intNavNumber + 1
  end if
  if gsProjects then %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>projects/"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs(); showNavSub('Projects',event);" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Projects")%></div></a>
								</td>
<%	intNavNumber = intNavNumber + 1
  end if
  if gsClients then %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>clients/"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs(); showNavSub('Clients',event);" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Clients")%></div></a>
								</td>
<%	intNavNumber = intNavNumber + 1
  end if
  if gsReports then %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>reports/"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs(); showNavSub('Reports',event);" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Reports")%></div></a>
								</td>
<%	intNavNumber = intNavNumber + 1
  end if 
  if gsOther then %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>tools/"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs(); showNavSub('Other',event);" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Tools")%></div></a>
								</td>
<%	intNavNumber = intNavNumber + 1
  end if %>
								<td align="center" bgcolor="<%=gsColorHighlight%>">
									<a href="<%=gsSiteRoot%>abandon.asp"><div class="nav" id="nav" name="nav"<%if gsDHTMLNavigation then%> onMouseOver="javascript: rollOnNav(<%=intNavNumber%>); killAllSubs();" onMouseOut="javascript: rollOffNav(<%=intNavNumber%>);"<%end if%>><%=dictLanguage("Logout")%></div></a>
								</td>																
							</tr>
						</table>
