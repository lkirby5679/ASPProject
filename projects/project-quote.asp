<%@ LANGUAGE="VBSCRIPT" %>
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
<!--#include file="../includes/main_page_header.asp"-->


<%
project_id = request("pid")
if project_id = "" then
	response.redirect ("main.asp")
end if

pqid = request("pqid")

if Request("Submit") = "Submit" then
	total_hours = request("QuoteTotalHours")
	total_price = request("QuoteTotalPrice")
	if total_hours = "" then
		total_hours = 0
	end if
	if total_price = "" then
		total_price = 0
	end if
	sql = sql_InsertProjectQuote(project_id, date(), SQLEncode(request("project_name")), _
		SQLEncode(request("client_name")), SQLEncode(request("account_rep")), _
		SQLEncode(request("work_order")), request("start_date"), request("end_date"), _
		SQLEncode(request("project_type")), SQLEncode(request("comments")), _
		total_hours, total_price)
	'Response.Write sql & "<BR><br>"
	Call DoSQL(sql)
	sql = sql_GetLatestProjectQuote()
	'Response.Write sql & "<BR><br>"
	Call RunSQL(sql, rs)
	if not rs.eof then
		pqid = rs("maxid")
	else
		pqid = 0
	end if
	rs.close
	set rs = nothing
	sql = sql_DeleteProjectQuoteItems(pqid)
	'Response.Write sql & "<BR><br>"
	Call DoSQL(sql)

	intvarfields = request("intvarfields")
	for i = 1 to intvarfields
		item_name = SQLEncode(request("Hours" & i)) 
		item_rate = request(i & "xRate")
		if item_rate = "" then
			item_rate = 0
		end if
		item_hours = request(i & "xHours")
		if item_hours = "" then
			item_hours = 0
		end if
		item_total = request(i & "xTotal")
		if item_total = "" then
			item_total = 0
		end if
		if item_name <> "" or item_total<>0 then
			sql = sql_InsertProjectQuoteItems(pqid, project_id, item_name, item_rate, item_hours, item_total)
			'Response.Write sql & "<BR><br>"
			Call DoSQL(sql)
		end if
	next
	for i = 1 to 3
		item_name = SQLEncode(request("AdditionalHours" & i)) 
		item_rate = request("AdditionalHours" & i & "xRate")
		if item_rate = "" then
			item_rate = 0
		end if
		item_hours = request("AdditionalHours" & i & "xHours")
		if item_hours = "" then
			item_hours = 0
		end if
		item_total = request("AdditionalHours" & i & "xTotal")
		if item_total = "" then
			item_total = 0
		end if
		if item_name <> "" or item_total<>0 then		
			sql = sql_InsertProjectQuoteItems(pqid, project_id, item_name, item_rate, item_hours, item_total)
			'Response.Write sql & "<BR><br>"
			Call DoSQL(sql)
		end if
	next	
	Response.Redirect "project-quote.asp?Submit=view&pid=" & project_id & "&pqid=" & pqid & ""

elseif request("Submit")="Delete" then
	sql = sql_DeleteProjectQuoteItems(pqid)
	Call DoSQL(sql) 
	sql = sql_DeleteProjectQuote(pqid)
	Call DoSQL(sql)%>
	
	<!--#include file="../includes/popup_page_open.asp"-->
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr><td align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("Project_Quote")%></b></td></tr>
		<tr><td align="center">The quote has been deleted.</td></tr>
	</table>		<%


elseif request("Submit")="view" then 
	sql = sql_GetProjectQuoteByID(pqid)
	Call RunSQL(sql, rs) 
	if not rs.eof then %>

	<!--#include file="../includes/popup_page_open.asp"-->
	<form method="post" action="project-quote.asp" name="strForm" id="strForm">
	<input type="hidden" name="pid" value="<%=project_id%>">
	<input type="hidden" name="pqid" value="<%=rs("projectquote_id")%>">
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr><td colspan="4" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("Project_Quote")%></b></td></tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Name")%>:</b></td>
			<td colspan="3"><%=rs("project_name")%></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client")%>:</b></td>
			<td colspan="3"><%=rs("client_name")%></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Date_Quote_Prepared")%>:</b></td>
			<td colspan="3"><%=rs("dateentered")%></td>
	    </tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Account_Rep")%>:</b></td>
			<td colspan="3"><%=rs("account_rep")%></td>
	    </tr>	
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Work_Order")%>:</b></td>
			<td colspan="3"><%=rs("work_order")%></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Date")%>:</b></td>
			<td colspan="3"><%=rs("start_date")%></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Date")%>:</b></td>
			<td colspan="3"><%=rs("end_date")%></td>
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Project_Type")%>:</b></td>
			<td colspan="3"><%=rs("Project_Type")%></td>
		</tr>
		<tr>
			<td valign="top"><b class="bolddark"><%=dictLanguage("Comments")%>:</b></td>
			<td colspan="3"><%=rs("comments")%></td>
		</tr> 	 	

		<tr>
			<td>&nbsp;</td>
			<td align="center" width="50"><%=dictLanguage("Rate")%></td>
			<td align="center" width="50"><%=dictLanguage("Hours")%></td>
			<td align="center" width="50"><%=dictLanguage("Total")%></td>
		</tr>
	
<%	sql = sql_GetProjectQuoteItemsByProjectQuoteID(pqid)
	'Response.Write sql & "<BR>"
	Call RunSQL(sql, rsRates) 
	while not rsRates.eof %>
		<tr>
			<td><%=rsRates("itemname")%></td>
			<td align="right"><%=rsRates("itemrate")%></td>
			<td align="right"><%=formatNumber(rsRates("itemhours"),2,-1,0,0)%></td>
			<td align="right"><%=formatNumber(rsRates("itemtotal"),2,-1,0,0)%></td>
		</tr>		
<%		rsRates.movenext
	wend 
	rsRates.close
	set rsRates = nothing %>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Total")%></b></td>
			<td>&nbsp;</td>
			<td align="right"><b class="bolddark"><%=formatNumber(rs("total_hours"),2,-1,0,0)%></b></td>
			<td align="right"><b class="bolddark"><%=formatNumber(rs("total_price"),2,-1,0,0)%></b></td>
		</tr>	
	</table>
<%	if session("permProjectsEdit") then%>	
	<p align="center"><input type="Button" name="Print" value="Print" class="formButton" onClick="javascript: window.print();">&nbsp;<input type="Submit" name="Submit" value="Delete" class="formButton"></p>
<%  end if %>
</form>
<%	end if
    rs.close
    set rs = nothing 
	
else
	sql = sql_GetProjectsByID(project_id)
	Call RunSQL(sql, rsProject)
		start_date = rsProject("Start_Date")
		if start_date = "" then
			start_date = date()
		else
			if not isdate(start_date) then
				start_date = date()
			end if
		end if
		launch_date = rsProject("Launch_Date")
		if launch_date = "" then
			launch_date = date()
		else
			if not isdate(launch_date) then
				launch_date = date()
			end if
		end if
		
	sql = sql_GetAllClients()
	Call RunSQL(sql, rsClientList)

	sql = sql_GetProjectTypes()
	Call RunSQL(sql, rsProjectType)

	sql = sql_GetEmployeeTypes()
	Call RunSQL(sql, rsRates)

	sql = sql_GetActiveEmployees()
	Call RunSQL(sql, rsEmployees)

	sql = sql_GetProjectBudgetsByID(project_id)
	Call RunSQL(sql, rsBudget)
%>

<!--#include file="../includes/popup_page_open.asp"-->

<script language="javascript">
//<--
isMoneyFormatAmount = 0.00;
isMoneyFormatString = "0.00";

function addRow(rowname) {
	rowRate = document.strForm[rowname + "xRate"].value
	rowRateBool  = isMoneyFormat(rowRate);
	rowRateAmount = isMoneyFormatAmount;
	
	rowHours = document.strForm[rowname + "xHours"].value
	rowHoursBool = isMoneyFormat(rowHours);
	rowHoursAmount = isMoneyFormatAmount;
	
	rowTotal = rowRateAmount * rowHoursAmount;
	rowTotalBool = isMoneyFormat(rowTotal);
	//alert(isMoneyFormatString);
	document.strForm[rowname + "xTotal"].value = isMoneyFormatString;
	}
	
function addHours() {
	var newValue = 0;
	for (var i = 0; i<document.strForm.elements.length; i++)   {
		var varStr = document.strForm.elements[i].name;
		var varStrLoc = varStr.indexOf("xHours");
		if (varStrLoc > 0) {
			newValue = eval(newValue) + eval(document.strForm.elements[i].value);
		}	
	}
	document.strForm["QuoteTotalHours"].value = newValue;
	}
	
function addTotals() {
	var newValue = 0;
	for (var i = 0; i<document.strForm.elements.length; i++)   {
		var varStr = document.strForm.elements[i].name;
		var varStrValue = document.strForm.elements[i].value
		var varStrValueBool = isMoneyFormat(varStrValue);
		var varStrValueString = "" + isMoneyFormatAmount;
		var varStrLoc = varStr.indexOf("xTotal");
		var varStrValueLoc = varStrValueString.indexOf(".00");
		if (varStrLoc > 0) {
			if (varStrValueLoc < 0) {
				strValueBool = isMoneyFormat(varStrValue);
				document.strForm.elements[i].value = isMoneyFormatString;
			}
			newValue = eval(eval(newValue) + eval(varStrValueString));
		}	
	}
	newValueBool = isMoneyFormat(newValue);
	document.strForm["QuoteTotalPrice"].value = isMoneyFormatString;
	}	

function isMoneyFormat(str) {
   isMoneyFormatAmount = 0.00;
   isMoneyFormatString = "0.00";
<% if formatNumber(0,2) = "0,00" then %>
		eur = 1;
<% else %>
		eur = 0;
<% end if %>
   //alert("eur=" + eur);
   if(!str) return false;
   str = "" + str; // force string
   for (var i=0; i<str.length;i++) {
      var ch = str.charAt(i);
      if (!isNum(ch) && ch!='.' && ch != ',' && ch !='-') return false
   }

   var sign = 1;
   var signChar = '';
   isMoneyFormatAmount = 0.00;
   isMoneyFormatString = "0.00";

   if (str.length > 1) {
      signChar = str.substring(0,1);
      if (signChar == '-' || signChar == '+' ) {
         if (signChar == '-') sign = -1;         
         str = str.substring(1);
      }
      else signChar = '';
   }
   var decimalPoint = '.';
   var thDelim = ',';
   if (eur) { 
      decimalPoint = ','
      thDelim = '.';
   }
   //alert("decimal=" + decimalPoint);
   test1 = str.split(decimalPoint);
   if (test1.length == 2) { // Decimals found
      if (test1[1].length > 2) return false; // more than 2 decimals
      if (isNum(test1[1])) {
         if (test1[1] < 9 && test1[1].charAt(0) > 0) test1[1] = new String(test1[1]+"0");
      }
      else return false; 
   }
   else test1[1] = "00"; // force decimals
   if (test1[0] == '') test1[0] = 0;
   if (test1[0] && test1[0].indexOf(thDelim) != -1) {
      test2 = test1[0].split(thDelim);  
      if (test2.length >= 2) { // thousands found
         var thError = false;
         for (var i=0;i<test2.length;i++) {
            if (test2[i].length < 3 && i != 0) { thError = true; break; } // all thousands exept the first.
            if (!isNum(test2[i])) { thError = true; break; } // all numbers
         }
         if (thError) return false;
         test1[0] = test2.join('')
      }
   }
   isMoneyFormatAmount = (parseInt(test1[0]) + parseFloat('.'+test1[1]))*sign;
   isMoneyFormatString = new String(""+signChar+""+parseInt(test1[0])) + decimalPoint + test1[1];
   return true;
} 

function isNum(str) {
  if(!str) return false;
  for(var i=0; i<str.length; i++){
    var ch=str.charAt(i);
    if ("0123456789".indexOf(ch) ==-1) return false;
  }
  return true;
} 

//-->
</script>

<form method="post" action="project-quote.asp" name="strForm" id="strForm">
<input type="hidden" name="pid" value="<%=project_id%>">
	<table border="0" cellpadding="2" cellspacing="2" align="center">
		<tr><td colspan="4" align="center" bgcolor="<%=gsColorHighlight%>" width="100%"><b class="homeHeader"><%=dictLanguage("Project_Quote")%></b></td></tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Name")%>:</b></td>
			<td colspan="3"><%=rsProject("Description")%></td>
			<input type="hidden" name="project_name" value="<%=rsProject("Description")%>">
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Client")%>:</b></td>
			<td colspan="3"><%=GetClientName(rsProject("Client_id"))%></td>
			<input type="hidden" name="client_name" value="<%=GetClientName(rsProject("Client_id"))%>">
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Account_Rep")%>:</b></td>
			<td colspan="3"><%=GetEmpName(rsProject("AccountExec_ID"))%></td>
			<input type="hidden" name="account_rep" value="<%=GetEmpName(rsProject("AccountExec_ID"))%>">
	    </tr>	
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Work_Order")%>:</b></td>
			<td colspan="3"><%=rsProject("WorkOrder_Number")%></td>
			<input type="hidden" name="work_order" value="<%=rsProject("WorkOrder_Number")%>">
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Start_Date")%>:</b></td>
			<td colspan="3"><%=start_date%></td>
			<input type="hidden" name="start_date" value="<%=start_date%>">
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("End_Date")%>:</b></td>
			<td colspan="3"><%=launch_date%></td>
			<input type="hidden" name="end_date" value="<%=launch_date%>">
		</tr>
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Project_Type")%>:</b></td>
			<td colspan="3"><%=GetProjectTypeName(rsProject("ProjectType_id"))%></td>
			<input type="hidden" name="project_type" value="<%=GetProjectTypeName(rsProject("ProjectType_id"))%>">
		</tr>
		<tr>
			<td valign="top"><b class="bolddark"><%=dictLanguage("Comments")%>:</b></td>
			<td colspan="3"><%=rsProject("comments")%></td>
			<input type="hidden" name="comments" value="<%=rsProject("comments")%>">
		</tr> 	 	

		<tr>
			<td>&nbsp;</td>
			<td align="center"><%=dictLanguage("Rate")%></td>
			<td align="center"><%=dictLanguage("Hours")%></td>
			<td align="center"><%=dictLanguage("Total")%></td>
		</tr>
	
<%quotetotal = 0
  hourstotal = 0
  intVarFields = 0
  while not rsRates.eof 
	intVarFields = intVarFields+1
	boolThisBudget = FALSE
	if not rsBudget.eof then
		if rsBudget("employeetype_id") = rsRates("employeetype_id") then 
			boolThisBudget = TRUE %>
		<tr>
			<td><input type="text" name="Hours<%=intVarFields%>" class="formstyleMed" size="20" value="<%=rsRates("EmployeeType")%>"></td>
			<td><input type="text" name="<%=intVarFields%>xRate" class="formstyleTiny" size="4" value="<%=rsBudget("rate")%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('<%=intVarFields%>');addTotals();"></td>
			<td><input type="text" name="<%=intVarFields%>xHours" class="formstyleTiny" size="4" value="<%=rsBudget("Hours")%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('<%=intVarFields%>');addHours();addTotals();"></td>
			<td><input type="text" name="<%=intVarFields%>xTotal" class="formstyleShort" size="6" value="<%=formatNumber(rsBudget("rate")*rsBudget("Hours"),2,-1,0,0)%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addTotals();"></td>
		</tr>
<%			hourstotal = hourstotal + rsBudget("Hours")
			quotetotal = quotetotal + (rsBudget("rate")*rsBudget("Hours"))
			rsBudget.movenext
		else %>
		<tr>
			<td><input type="text" name="Hours<%=intVarFields%>" class="formstyleMed" size="20" value="<%=rsRates("EmployeeType")%>"></td>
			<td><input type="text" name="<%=intVarFields%>xRate" class="formstyleTiny" size="4" value="<%=rsRates("Rate")%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('<%=intVarFields%>');addTotals();"></td>
			<td><input type="text" name="<%=intVarFields%>xHours" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('<%=intVarFields%>');addHours();addTotals();"></td>
			<td><input type="text" name="<%=intVarFields%>xTotal" class="formstyleShort" size="6" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addTotals();"></td>
		</tr>		
<%		end if
	else %>
		<tr>
			<td><input type="text" name="Hours<%=intVarFields%>" class="formstyleMed" size="20" value="<%=rsRates("EmployeeType")%>"></td>
			<td><input type="text" name="<%=intVarFields%>xRate" class="formstyleTiny" size="4" value="<%=rsRates("Rate")%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('<%=intVarFields%>');addTotals();"></td>
			<td><input type="text" name="<%=intVarFields%>xHours" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('<%=intVarFields%>');addHours();addTotals();"></td>
			<td><input type="text" name="<%=intVarFields%>xTotal" class="formstyleShort" size="6" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addTotals();"></td>
		</tr>	
<%	end if
	rsRates.movenext 
wend %>
		<input type="hidden" name="intVarFields" value="<%=intVarFields%>">
		<tr>
			<td><input type="text" name="AdditionalHours1" class="formstyleMed" value="Additional Hours" size="20"></td>
			<td><input type="text" name="AdditionalHours1xRate" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('AdditionalHours1');addTotals();"></td>
			<td><input type="text" name="AdditionalHours1xHours" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('AdditionalHours1');addHours();addTotals();"></td>
			<td><input type="text" name="AdditionalHours1xTotal" class="formstyleShort" size="6" value="<%=formatNumber(0,2,-1,0,0)%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addTotals();"></td>
		</tr>	
		<tr>
			<td><input type="text" name="AdditionalHours2" class="formstyleMed" value="Additional Hours" size="20"></td>
			<td><input type="text" name="AdditionalHours2xRate" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('AdditionalHours2');addTotals();"></td>
			<td><input type="text" name="AdditionalHours2xHours" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('AdditionalHours2');addHours();addTotals();"></td>
			<td><input type="text" name="AdditionalHours2xTotal" class="formstyleShort" size="6" value="<%=formatNumber(0,2,-1,0,0)%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addTotals();"></td>
		</tr>
		<tr>
			<td><input type="text" name="AdditionalHours3" class="formstyleMed" value="Additional Hours" size="20"></td>
			<td><input type="text" name="AdditionalHours3xRate" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('AdditionalHours3');addTotals();"></td>
			<td><input type="text" name="AdditionalHours3xHours" class="formstyleTiny" size="4" value="0" onkeypress="txtDate_onKeypress();" onChange="javascript: addRow('AdditionalHours3');addHours();addTotals();"></td>
			<td><input type="text" name="AdditionalHours3xTotal" class="formstyleShort" size="6" value="<%=formatNumber(0,2,-1,0,0)%>" onkeypress="txtDate_onKeypress();" onChange="javascript: addTotals();"></td>
		</tr>					
		<tr>
			<td><b class="bolddark"><%=dictLanguage("Total")%></b></td>
			<td>&nbsp;</td>
			<td><input type="text" name="QuoteTotalHours" class="formstyleTiny" size="4" value="<%=hourstotal%>"></td>
			<td><input type="text" name="QuoteTotalPrice" class="formstyleShort" size="6" value="<%=formatnumber(quotetotal,2,-1,0,0)%>"></td>
		</tr>	
	</table>
	<%if session("permProjectsEdit") then%>
	<p align="center"><input type="Submit" name="Submit" value="Submit" class="formButton"> <input type="Button" name="Print" value="Print" class="formButton" onClick="javascript: window.print();"> <input type="Reset" name="Reset" value="Reset" class="formButton"></p>
	<%end if%>
</form>

<%
	rsEmployees.Close
	set rsEmployees=nothing
	rsClientList.Close
	set rsClientList=nothing
	rsProjectType.Close
	set rsProjectType=nothing
	rsProject.close
	set rsProject = nothing
	rsRates.close
	set rsRates = nothing
	rsBudget.close
	set rsBudget = nothing 

	Function GetEmpName(ID)
		if ID <> "" then
			sql = sql_GetEmployeesByID(ID)
			'response.write sql
			Call RunSQL(sql, rs)
			if not rs.eof then
				empName = rs("EmployeeName")
			else
				empName = ""
			end if
			rs.close
			set rs = nothing
		else
			empName = ""
		end if
		GetEmpName = empName
	End Function
	
	Function GetClientName(ID)
		if ID <> "" then
			sql = sql_GetClientsByID(ID)
			'response.write sql
			Call RunSQL(sql, rs)
			if not rs.eof then
				clientName = rs("Client_Name")
			else
				clientName = ""
			end if
			rs.close
			set rs = nothing
		else
			clientName = ""
		end if
		GetClientName = clientName
	End Function	
	
	Function GetProjectTypeName(ID)
		if ID <> "" then
			sql = sql_GetProjectTypeByID(ID)
			'response.write sql
			Call RunSQL(sql, rs)
			if not rs.eof then
				projecttypename = rs("projecttypedescription")
			else
				projecttypename = ""
			end if
			rs.close
			set rs = nothing
		else
			projectypename = ""
		end if
		GetProjectTypeName = projecttypename
	End Function		
end if
%>

<!--#include file="../includes/popup_page_close.asp"-->