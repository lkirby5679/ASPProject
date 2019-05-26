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
thispage = "answer_add.asp"

poll = trim(request("pollid"))
question = trim(request("question"))

if question <> "" then %>

	<script language="JavaScript">
	// <!--
	function isBlank(field) {
		if ((field == "") || (field == null)) {
			return true
		}
		return false
	}

	function submitIt(form){
		if (isBlank(form.answer.value)) {
		    alert("<%=dictLanguage("Error_No_Answer")%>")
			form.answer.focus()
		  	form.answer.select()
			return false
		}
		return true
	}
	//  -->
	</SCRIPT>

<center>
<form method="post" action="answer_add_processor.asp" name="form1">
<input type="hidden" name="pollid" value="<%=poll%>">
<input type="hidden" name="question" value="<%=question%>">
<br>
<table border="0">
	<tr>
		<td><%=dictLanguage("Answer")%>:</td>
		<td><input type="text" name="answer" class="formstylelong"></td>
	</tr>
	<tr><td colspan="2" align="center"><input type="Submit" value="Submit" class="formbutton" onClick="return submitIt(form1)" id=Submit1 name=Submit1></td></tr>	
</table>
</form>
</center>
<%end if%>


<br>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>				
</p>


<!--#include file="../../includes/main_page_close.asp"-->
