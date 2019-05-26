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
thispage = "question_edit.asp"

poll = trim(request("pollid"))
question = trim(request("question"))

if question <> "" then
	sql = sql_GetQuestionByID(question)
	Call RunSQL(sql, rsQuestion)
	if not rsQuestion.eof then
		question_text = rsQuestion("fdPollQuestion")
		question_type = rsQuestion("fdPollQuestionType")	
	end if
	rsQuestion.close
	set rsQuestion = nothing %>

	<script language="JavaScript">
	// <!--
	function isBlank(field) {
		if ((field == "") || (field == null)) {
			return true
		}
		return false
	}
	
	function submitIt(form){
		if (isBlank(form.question_text.value)) {
		    alert("<%=dictLanguage("Error_No_Question")%>")
			form.question_text.focus()
		  	form.question_text.select()
			return false
		}
		return true
	}
	//  -->
	</SCRIPT>

<center>
<form method="post" action="question_edit_processor.asp" name="form1">
<input type="hidden" name="pollid" value="<%=poll%>">
<input type="hidden" name="question" value="<%=question%>">
<input type="hidden" name="old_type" value="<%=question_type%>">
<br>
<table border="0">
	<tr>
		<td><%=dictLanguage("Question")%>:</td>
		<td><input type="text" name="question_text" value="<%=question_text%>" class="formstylelong"></td>
	</tr>
	<tr>
		<td><%=dictLanguage("Type")%>:</td>
		<td><select name="type" class="formstylelong">
			<option value="1"<%if question_type=1 then%> Selected<%end if%>><%=dictLanguage("Radio_Buttons")%></option>
			<option value="2"<%if question_type=2 then%> Selected<%end if%>><%=dictLanguage("Checkboxes")%></option>
			<option value="3"<%if question_type=3 then%> Selected<%end if%>><%=dictLanguage("Text_Field")%></option>
  			<option value="0"<%if question_type=0 then%> Selected<%end if%>><%=dictLanguage("Label")%></option>
			</select>
		</td>	
	</tr>	
	<tr><td colspan="2" align="center"><input type="Submit" value="Submit" class="formbutton" onClick="return submitIt(form1)" id=Submit1 name=Submit1></td></tr>	
</table>
</form>
<a href="question_delete_processor.asp?pollid=<%=poll%>&question=<%=question%>" onClick="javascript: return confirm('<%=dictLanguage("Confirm_Delete_Question")%>');"><%=dictLanguage("Click_Delete_Question")%></a>
<br>
</center>
<%end if%>

<br>
<p align="center">
<a href="default.asp"><%=dictLanguage("Return_Survey_Admin_Home")%></a><br>	
<a href=""><%=dictLanguage("Return_Admin_Home")%></a><br>				
</p>


<!--#include file="../../includes/main_page_close.asp"-->
