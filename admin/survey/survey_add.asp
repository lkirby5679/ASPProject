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
thispage = "survey_add.asp"

poll = trim(request("pollid"))

if poll = "" then %>
	<script language="JavaScript">
	// <!--
	function isBlank(field) {
		if ((field == "") || (field == null)) {
			return true
		}
		return false
	}

	function submitIt(form){
		if (isBlank(form.pollid.value)) {
		    alert("<%=dictLanguage("Error_No_Topic")%>")
			form.pollid.focus()
			return false
		}
		return true
	}
	//  -->
	</SCRIPT>

<center>
<form method="post" action="<%=thispage%>" name="form1">
<br>
<table border="0">
	<tr>
		<td><%=dictLanguage("Survey_Topic")%>:</td>
		<td><input type="text" name="pollid" class="formstylelong"></td>
	</tr>
	<tr><td colspan="2" align="center"><input type="Submit" value="Submit" onClick="return submitIt(form1)" id=Submit1 name=Submit1 class="formButton"></td></tr>	
</table>
</form>
</center>
<%
end if

if poll <> "" then
	sql = sql_InsertSurvey(SQLEncode(poll))
	Call DoSQL(sql)
	response.write ("<br><Center>" & dictLanguage("SurveyInst_1") & "<a href='survey_edit.asp'>" & dictLanguage("Here") & "</a>, " & dictLanguage("SurveyInst_2") & ".<br></center>")	
end if%>
		
<br>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>				
</p>

<!--#include file="../../includes/main_page_close.asp"-->
