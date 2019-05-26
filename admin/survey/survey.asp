<%@language="VBSCript"%>
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

<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<%
Response.Buffer = "true"
thispage = "survey.asp"

LimitNumofVotes = FALSE
RecordInCookie = TRUE 
WindowSize = 400

sql = sql_GetCurrentSurvey()
Call RunSQL(sql, rsPolls)
if not rsPolls.eof then
	PollQuestion = rsPolls("fdPollID")	%>
		<table border="0" width="<%=WindowSize%>" cellpadding="2" cellspacing="2" bgcolor="<%=gsColorBackground%>" align="Center">
			<tr><td bgcolor="<%=gsColorHighlight%>" align="center"><b class="homeheader"><%=dictlanguage("Active_Survey")%></b></td></tr>
			<tr style="display: <%=strDisplay%>;" name="poll" id="poll">
				<td>
				<!--#include file="../../includes/survey.asp"-->				
				</td>
			</tr>
		</table>	
<% 	
else response.write ("<br>" & dictLanguage("No_Active_Survey") & "<br>")
end if
rsPolls.close
set rsPolls = nothing %>	

<br>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>				
</p>	
		
<!--#include file="../../includes/main_page_close.asp"-->