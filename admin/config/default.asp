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

<%
strAct		= trim(request("act"))
strOldSetting = trim(request("oldSetting"))
%>

<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<script language="JavaScript">
//<!--
function foreColor(strField){
   	var arr = showModalDialog("selcolor.htm","","font-family:Verdana; font-size:12; dialogWidth:30em; dialogHeight:34em" );
   	if (arr != null){ 
   		document.strForm(strField).value=arr;
   	}
}   
//-->
</script>

<table border="0" cellspacing="2" cellpadding="2">
<tr><td valign="top">

<table border="0" cellspacing="2" cellpadding="2">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td colspan="4" align="Center" class="homeheader"><%=dictLanguage("Admin")%>--<%=dictLanguage("Site_Configuration")%></td>
	</tr>
	<tr>
		<td valign="top"><%=dictLanguage("Background")%>:</td>
		<td valign="top" nowrap><%=gsColorBackground%></td>
		<td valign="top" bgcolor="<%=gsColorBackground%>" width="10">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=editBackground&oldSetting=<%=gsColorBackground%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Background"></a></td>
	</tr>
	<tr>
		<td valign="top"><%=dictLanguage("Highlight")%>:</td>
		<td valign="top" nowrap><%=gsColorHighlight%></td>
		<td valign="top" bgcolor="<%=gsColorHighlight%>" width="10">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=editHighlight&oldSetting=<%=gsColorHighlight%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Highlight"></a></td>
	</tr>	
	<tr>
		<td valign="top"><%=dictLanguage("Alt_Highlight")%>:</td>
		<td valign="top" nowrap><%=gsColorAltHighlight%></td>
		<td valign="top" bgcolor="<%=gsColorAltHighlight%>" width="10">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=editAltHighlight&oldSetting=<%=gsColorAltHighlight%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Alt Highlight"></a></td>
	</tr>
	<tr>
		<td valign="top"><%=dictLanguage("White")%>:</td>
		<td valign="top" nowrap><%=gsColorWhite%></td>
		<td valign="top" bgcolor="<%=gsColorWhite%>" width="10">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=editWhite&oldSetting=<%=gsColorWhite%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit White"></a></td>
	</tr>	
	<tr>
		<td valign="top"><%=dictLanguage("Black")%>:</td>
		<td valign="top" nowrap><%=gsColorBlack%></td>
		<td valign="top" bgcolor="<%=gsColorBlack%>" width="10">&nbsp;</td>
		<td valign="top"><a href="default.asp?act=editBlack&oldSetting=<%=gsColorBlack%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Black"></a></td>
	</tr>	
	<tr><td colspan="4">&nbsp;</td></tr>
	<tr>
		<td valign="top"><%=dictLanguage("Site_Name")%>:</td>
		<td valign="top" colspan="2" nowrap><%=gsSiteName%></td>
		<td valign="top"><a href="default.asp?act=editSiteName&oldSetting=<%=gsSiteName%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Site Name"></a></td>
	</tr>	
	<tr>
		<td valign="top"><%=dictLanguage("Site_Logo")%>:</td>
		<td valign="top" colspan="2" nowrap><a href="<%=gsSiteRoot%>images/<%=gsSiteLogo%>" target="_blank"><%=gsSiteLogo%></a></td>
		<td valign="top"><a href="default.asp?act=editSiteLogo&oldSetting=<%=gsSiteLogo%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Site Logo"></a></td>
	</tr>	
	<tr>
		<td valign="top"><%=dictLanguage("Admin_Email")%>:</td>
		<td valign="top" colspan="2" nowrap><%=gsAdminEmail%></td>
		<td valign="top"><a href="default.asp?act=editAdminEmail&oldSetting=<%=gsAdminEmail%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Admin Email"></a></td>
	</tr>	
	<tr>
		<td valign="top"><%=dictLanguage("Mail_Host")%>:</td>
		<td valign="top" colspan="2" nowrap><%=gsMailHost%></td>
		<td valign="top"><a href="default.asp?act=editMailHost&oldSetting=<%=gsMailHost%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Mail Host"></a></td>
	</tr>	
	<tr>
		<td valign="top"><%=dictLanguage("Site_URL")%>:</td>
		<td valign="top" colspan="2" nowrap><%=gsSiteURL%></td>
		<td valign="top"><a href="default.asp?act=editSiteURL&oldSetting=<%=gsSiteURL%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Site URL"></a></td>
	</tr>							
	<tr>
		<td valign="top"><%=dictLanguage("Site_Root")%>:</td>
		<td valign="top" colspan="2" nowrap><%=gsSiteRoot%></td>
		<td valign="top"><a href="default.asp?act=editSiteRoot&oldSetting=<%=gsSiteRoot%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Site Root"></a></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2"><%=dictLanguage("Mail_Component")%>:</td>
		<td valign="top"><%=gsMailComponent%></td>
		<td valign="top"><a href="default.asp?act=editMailComponent&oldSetting=<%=gsEmailComponent%>"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Mail Component"></a></td>
	</tr>
	<tr>
		<td valign="top" colspan="3"><%=dictLanguage("Meta_Description")%></td>
		<td valign="top"><a href="default.asp?act=editMetaDescription"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Meta Description"></a></td>
	</tr>
	<tr>
		<td valign="top" colspan="3"><%=dictLanguage("Meta_Keywords")%></td>
		<td valign="top"><a href="default.asp?act=editMetaKeywords"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Meta Keywords"></a></td>
	</tr>
	<tr>
		<td valign="top" colspan="3"><%=dictLanguage("Functionality_Options")%>:</td>
		<td valign="top"><a href="default.asp?act=editFunctionality"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Site Functionality"></a></td>
	</tr>		
	<tr>
		<td valign="top" colspan="3"><%=dictLanguage("Home_Page_Options")%>:</td>
		<td valign="top"><a href="default.asp?act=editHomePage"><IMG SRC="../images/edit.gif" height="10" width="10" border="0" alt="Edit Home Page Options"></a></td>
	</tr>	
</table>

<br><br>

</td><td align="center" width="100%" valign="top">

<table border="0" cellspacing="2" cellpadding="2" width="100%">
	<tr bgcolor="<%=gsColorHighlight%>">
		<td align="Center" class="homeheader"><%=dictLanguage("Workspace")%></td>
	</tr>
</table>

<%if strAct = "editFunctionality" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditFunctionality">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="4" align="center"><b><%=dictLanguage("Functionality_Options")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Top_Nav")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsNavigation" value="1" <%if gsNavigation then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Production_Report_Grid")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsProductionReportGrid" value="1" <%if gsProductionReportGrid then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("DHTML_Nav")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsDHTMLNavigation" value="1" <%if gsDHTMLNavigation then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Reports")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsReports" value="1" <%if gsReports then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Timecards")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsTimecards" value="1" <%if gsTimecards then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Calendar")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsCalendar" value="1" <%if gsCalendar then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Tasks")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsTasks" value="1" <%if gsTasks then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Discussion")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsDiscussion" value="1" <%if gsDiscussion then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Projects")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsProjects" value="1" <%if gsProjects then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Employee_Directory")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsEmployeeDirectory" value="1" <%if gsEmployeeDirectory then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Clients")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsClients" value="1" <%if gsClients then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("File_Repository")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsFileRepository" value="1" <%if gsFileRepository then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Client_Contact_Log")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsClientContactLog" value="1" <%if gsClientContactLog then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsNews" value="1" <%if gsNews then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Timesheets")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsTimesheets" value="1" <%if gsTimesheets then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Resources")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsResources" value="1" <%if gsResources then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Project_Dashboard")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsProjectDashboard" value="1" <%if gsProjectDashboard then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Options")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsOptions" value="1" <%if gsOptions then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Project_Quotes")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsProjectQuotes" value="1" <%if gsProjectQuotes then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Auto_Login")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsAutoLogin" value="1" <%if gsAutoLogin then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Host_Non_Billable")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHostNonBillable" value="1" <%if gsHostNonBillable then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Show_Clients_in_Calendar")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsCalendarClients" value="1" <%if gsCalendarClients then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Send_Task_Emails")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsTaskEmails" value="1" <%if gsTaskEmails then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Show_Projects_in_Calendar")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsCalendarProjects" value="1" <%if gsCalendarProjects then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Generate_Work_Order_Numbers")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsDynWorkOrderNumber" value="1" <%if gsDynWorkOrderNumber then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("Show_Tasks_in_Calendar")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsCalendarTasks" value="1" <%if gsCalendarTasks then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top">&nbsp;</td>
		<td valign="top">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("PTO")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsPTO" value="1" <%if gsPTO then Response.Write "Checked"%>></td>
	</tr>													
	<tr><td colspan="4">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="4" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditFunctionality" then

	strLookingFor = "const gsNavigation "
	strComments = "'whether to use the top nav or not"
	strNewSetting = trim(request("gsNavigation"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Top_Nav_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Top_Nav_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsDHTMLNavigation "
	strComments = "'whether to use the DHTML of the navigation or not"
	strNewSetting = trim(request("gsDHTMLNavigation"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("DHTML_Nav_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("DHTML_Nav_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsTimecards "
	strComments = "'Use Timecards or not.  If using timecards Clients and Projects must be enabled."
	strNewSetting = trim(request("gsTimecards"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Timecards_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Timecards_Not_Updated") & strErrorMessage & "<BR>"
	end if		

	strLookingFor = "const gsTasks "
	strComments = "'Use Tasks or not. If using tasks Clients and Projects must be enabled."
	strNewSetting = trim(request("gsTasks"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Tasks_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Tasks_Not_Updated") & strErrorMessage & "<BR>"
	end if		

	strLookingFor = "const gsProjects "
	strComments = "'Use Projects or not"
	strNewSetting = trim(request("gsProjects"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Projects_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Projects_Not_Updated") & strErrorMessage & "<BR>"
	end if	
	
	strLookingFor = "const gsClients "
	strComments = "'Use Clients or not"
	strNewSetting = trim(request("gsClients"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Clients_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Clients_Not_Updated") & strErrorMessage & "<BR>"
	end if	
	
	strLookingFor = "const gsClientContactLog "
	strComments = "'Use Client Contact Log or not.  You must use clients to use this."
	strNewSetting = trim(request("gsClientContactLog"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Client_Contact_Log_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Client_Contact_Log_Not_Updated") & strErrorMessage & "<BR>"
	end if			

	strLookingFor = "const gsTimeSheets "
	strComments = "'Use Timesheets for hourly employees or not"
	strNewSetting = trim(request("gsTimesheets"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Timesheets_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Timesheets_Not_Updated") & strErrorMessage & "<BR>"
	end if			

	strLookingFor = "const gsProjectDashboard "
	strComments = "'Use Project Dashboard page or not."
	strNewSetting = trim(request("gsProjectDashboard"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Project_Dashboard_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Project_Dashboard_Not_Updated") & strErrorMessage & "<BR>"
	end if
	
	strLookingFor = "const gsProjectQuotes "
	strComments = "'Use Project Quotes or not."
	strNewSetting = trim(request("gsProjectQuotes"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Project_Quotes_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Project_Quotes_Not_Updated") & strErrorMessage & "<BR>"
	end if	
	
	strLookingFor = "const gsHostNonBillable "
	strComments = "'Is Time charged agains the host company Non-billable?"
	strNewSetting = trim(request("gsHostNonBillable"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Host_Non_Billable_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Host_Non_Billable_Not_Updated") & strErrorMessage & "<BR>"
	end if	
	
	strLookingFor = "const gsTaskEmails "
	strComments = "'Send Task Emails"
	strNewSetting = trim(request("gsTaskEmails"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Task_Emails_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Task_Emails_Not_Updated") & strErrorMessage & "<BR>"
	end if	
	
	strLookingFor = "const gsDynWorkOrderNumber "
	strComments = "'If True all work order numbers will be generated on the fly "
	strNewSetting = trim(request("gsDynWorkOrderNumber"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Dynamic_Work_Order_Number_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Dynamic_Work_Order_Number_Not_Updated") & strErrorMessage & "<BR>"
	end if	
		
	strLookingFor = "const gsProductionReportGrid "
	strComments = "'Use the production report grid or not. You must use clients, projects, and timecards to use this."
	strNewSetting = trim(request("gsProductionReportGrid"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Production_Report_Grid_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Production_Report_Grid_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsReports "
	strComments = "'Use Reports or not. You must use clients, projects and timecards to use this."
	strNewSetting = trim(request("gsReports"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Reports_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Reports_Not_Updated") & strErrorMessage & "<BR>"
	end if		

	strLookingFor = "const gsCalendar "
	strComments = "'Use Calendar or not"
	strNewSetting = trim(request("gsCalendar"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Not_Updated") & "<BR>"
	end if	

	strLookingFor = "const gsDiscussion "
	strComments = "'Use Discussion Forum or not"
	strNewSetting = trim(request("gsDiscussion"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Discussion_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Discussion_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsEmployeeDirectory "
	strComments = "'Use Employee Directory or not"
	strNewSetting = trim(request("gsEmployeeDirectory"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Employee_Directory_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Employee_Directory_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsFileRepository "
	strComments = "'Use File Repository or not"
	strNewSetting = trim(request("gsFileRepository"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("File_Repository_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("File_Repository_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsNews "
	strComments = "'Use News or not"
	strNewSetting = trim(request("gsNews"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("News_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsResources "
	strComments = "'Use Resources or not"
	strNewSetting = trim(request("gsResources"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Resources_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Resources_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsOptions "
	strComments = "'Use User defined options or not, they can change colors, etc"
	strNewSetting = trim(request("gsOptions"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Options_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Options_Not_Updated") & strErrorMessage & "<BR>"
	end if	
	
	strLookingFor = "const gsAutoLogin "
	strComments = "'Does the site Auto Login the user if their cookie exists?"
	strNewSetting = trim(request("gsAutoLogin"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Auto_Login_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Auto_Login_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsCalendarClients "
	strComments = "'Show Client Start date in Calendar	"
	strNewSetting = trim(request("gsCalendarClients"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Clients_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Clients_Not_Updated") & strErrorMessage & "<BR>"
	end if
	
	strLookingFor = "const gsCalendarProjects "
	strComments = "'Show Project Start and end dates in calendar "
	strNewSetting = trim(request("gsCalendarProjects"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Projects_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Projects_Not_Updated") & strErrorMessage & "<BR>"
	end if
		
	strLookingFor = "const gsCalendarTasks "
	strComments = "'Show Uncompleted Task due dates in calendar for assignor and assignee "
	strNewSetting = trim(request("gsCalendarTasks"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Tasks_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Tasks_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsPTO "
	strComments = "'Use PTO or not. "
	strNewSetting = trim(request("gsPTO"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("PTO_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("PTO_Not_Updated") & strErrorMessage & "<BR>"
	end if

	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editHomePage" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditHomePage">
<table cellpadding="2" cellspacing="2" align="center">
	<tr><td colspan="4" align="center"><b><%=dictLanguage("Home_Page_Options")%></b></td></tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Minimize_Maximize_Buttons")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeMinimize" value="1" <%if gsHomeMinimize then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNews" value="1" <%if gsHomeExtNews then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Thoughts")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeThoughts" value="1" <%if gsHomeThoughts then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News_US_News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsUS" value="1" <%if gsHomeExtNewsUS then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Production_Report")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeProductionReport" value="1" <%if gsHomeProductionReport then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News_World_News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsWorld" value="1" <%if gsHomeExtNewsWorld then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeNews" value="1" <%if gsHomeNews then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News_Tech_News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsTech" value="1" <%if gsHomeExtNewsTech then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Resources")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeResources" value="1" <%if gsHomeResources then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News_Web_News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsWeb" value="1" <%if gsHomeExtNewsWeb then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top"><b><%=dictLanguage("Calendar")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeCalendar" value="1" <%if gsHomeCalendar then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News_Business_News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsBusiness" value="1" <%if gsHomeExtNewsBusiness then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Discussion")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeDiscussion" value="1" <%if gsHomeDiscussion then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News_Science_News")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsScience" value="1" <%if gsHomeExtNewsScience then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top"><b><%=dictLanguage("Survey")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeSurvey" value="1" <%if gsHomeSurvey then Response.Write "Checked"%>></td>
		<td valign="top"><b><%=dictLanguage("External_News_Sports")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsSports" value="1" <%if gsHomeExtNewsSports then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Gaming")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsGaming" value="1" <%if gsHomeExtNewsGaming then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Linux")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsLinux" value="1" <%if gsHomeExtNewsLinux then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Linux_Gaming")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsLinuxGaming" value="1" <%if gsHomeExtNewsLinuxGaming then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Linux_Software")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsLinuxSoftware" value="1" <%if gsHomeExtNewsLinuxSoftware then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Mac")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsMac" value="1" <%if gsHomeExtNewsMac then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Mac_Gaming")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsMacGaming" value="1" <%if gsHomeExtNewsMacGaming then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Mac_Software")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsMacSoftware" value="1" <%if gsHomeExtNewsMacSoftware then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Show_Title")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsShowTitle" value="1" <%if gsHomeExtNewsShowTitle="yes" then Response.Write "Checked"%>></td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Show_Time")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsShowTime" value="1" <%if gsHomeExtNewsShowTime="yes" then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Show_Indent")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsShowIndent" value="1" <%if gsHomeExtNewsShowIndent="yes" then Response.Write "Checked"%>></td>
	</tr>
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Show_Number")%>:</b></td>
		<td valign="top">
			<select name="gsHomeExtNewsShowNumber" size="1" class="formstyletiny">
<% for i = 1 to 10 
	Response.Write "<option value=""" & i & """"
	if trim(gsHomeExtNewsShowNumber) = trim(i) then 
		Response.Write " Selected"
	end if
	Response.Write ">" & i & "</option>"
   next %>	</select>
		</td>
	</tr>	
	<tr>
		<td valign="top" colspan="2">&nbsp;</td>
		<td valign="top"><b><%=dictLanguage("External_News_Open_New_Window")%>:</b></td>
		<td valign="top"><input type="checkbox" name="gsHomeExtNewsShowNewWindow" value="1" <%if gsHomeExtNewsShowNewWindow="yes" then Response.Write "Checked"%>></td>
	</tr>	
	<tr><td colspan="4">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="4" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditHomePage" then

	strLookingFor = "const gsHomeMinimize "
	strComments = "'whether the feature boxes on the home page are able to be collapsed/expanded "
	strNewSetting = trim(request("gsHomeMinimize"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Minimize_Maximize_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Minimize_Maximize_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeThoughts "
	strComments = "'which of the features on the home page do you want turned on? "
	strNewSetting = trim(request("gsHomeThoughts"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Thoughts_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Thoughts_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeProductionReport "
	strComments = ""
	strNewSetting = trim(request("gsHomeProductionReport"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Production_Report_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Production_Report_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeNews "
	strComments = ""
	strNewSetting = trim(request("gsHomeNews"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("News_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeResources "
	strComments = ""
	strNewSetting = trim(request("gsHomeResources"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Resources_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Resources_Not_Updated") & strErrorMessage & "<BR>"
	end if
	
	strLookingFor = "const gsHomeCalendar "
	strComments = ""
	strNewSetting = trim(request("gsHomeCalendar"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Calendar_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsHomeDiscussion "
	strComments = ""
	strNewSetting = trim(request("gsHomeDiscussion"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Discussion_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Discussion_Not_Updated") & strErrorMessage & "<BR>"
	end if
	
	strLookingFor = "const gsHomeSurvey "
	strComments = ""
	strNewSetting = trim(request("gsHomeSurvey"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("Survey_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("Survey_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNews "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNews"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsUS "
	strComments = "'these are options for the included news feed"
	strNewSetting = trim(request("gsHomeExtNewsUS"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_US_News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_US_News_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsWorld "
	strComments = "'these options are what categories you would"
	strNewSetting = trim(request("gsHomeExtNewsWorld"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_World_News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_World_News_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsTech "
	strComments = "'like displayed"
	strNewSetting = trim(request("gsHomeExtNewsTech"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Tech_News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Tech_News_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsWeb "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsWeb"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Web_News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Web_News_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsBusiness "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsBusiness"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Business_News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Business_News_Not_Updated") & strErrorMessage & "<BR>"
	end if
	
	strLookingFor = "const gsHomeExtNewsScience "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsScience"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Science_News_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Science_News_Not_Updated") & strErrorMessage & "<BR>"
	end if
	
	strLookingFor = "const gsHomeExtNewsSports "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsSports"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Sports_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Sports_Not_Updated") & strErrorMessage & "<BR>"
	end if
	
	strLookingFor = "const gsHomeExtNewsGaming "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsGaming"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Gaming_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Gaming_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsLinux "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsLinux"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Linux_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Linux_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsLinuxGaming "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsLinuxGaming"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Linux_Gaming_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Linux_Gaming_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsLinuxSoftware "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsLinuxSoftware"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Linux_Software_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Linux_Software_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsMac "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsMac"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Mac_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Mac_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsMacGaming "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsMacGaming"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Mac_Gaming_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Mac_Gaming_Not_Updated") & strErrorMessage & "<BR>"
	end if	
	
	strLookingFor = "const gsHomeExtNewsMacSoftware "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsMacSoftware"))
	if strNewSetting = "1" then
		strNewSetting = "TRUE"
	else
		strNewSetting = "FALSE"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Mac_Software_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Mac_Software_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsHomeExtNewsShowTitle "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsShowTitle"))
	if strNewSetting = "1" then
		strNewSetting = "yes"
	else
		strNewSetting = "no"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Title_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Title_Not_Updated") & strErrorMessage & "<BR>"
	end if	

	strLookingFor = "const gsHomeExtNewsShowTime "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsShowTime"))
	if strNewSetting = "1" then
		strNewSetting = "yes"
	else
		strNewSetting = "no"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Time_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Time_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsShowIndent "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsShowIndent"))
	if strNewSetting = "1" then
		strNewSetting = "yes"
	else
		strNewSetting = "no"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Indent_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Indent_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsShowNumber "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsShowNumber"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Number_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_Number_Not_Updated") & strErrorMessage & "<BR>"
	end if

	strLookingFor = "const gsHomeExtNewsShowNewWindow "
	strComments = ""
	strNewSetting = trim(request("gsHomeExtNewsShowNewWindow"))
	if strNewSetting = "1" then
		strNewSetting = "yes"
	else
		strNewSetting = "no"
	end if
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_New_Window_Updated") & "<BR>"
	else
		session("strMessage") = session("strMessage") & dictLanguage("External_News_Show_New_Window_Not_Updated") & strErrorMessage & "<BR>"
	end if

	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editBackground" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditBackground">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Background")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsColorBackground" value="<%=gsColorBackground%>" class="formstyleshort">
			<a href="javascript: foreColor('gsColorBackground');"><img src="../../images/chartblock.gif" width="16" height="13" border="0" alt="<%=dictLanguage("Select_Color")%>"></a>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditBackground" then

	strLookingFor = "gsColorBackground "
	strComments = "'darker color used throughout the site and for text in some places"
	strNewSetting = trim(request("gsColorBackground"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editHighlight" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditHighlight">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Highlight")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsColorHighlight" value="<%=gsColorHighlight%>" class="formstyleshort">
			<a href="javascript: foreColor('gsColorHighlight');"><img src="../../images/chartblock.gif" width="16" height="13" border="0" alt="<%=dictLanguage("Select_Color")%>"></a>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditHighlight" then

	strLookingFor = "gsColorHighlight "
	strComments = "'light color used throughout the site"
	strNewSetting = trim(request("gsColorHighlight"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editAltHighlight" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditAltHighlight">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Alt_Highlight")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsColorAltHighlight" value="<%=gsColorAltHighlight%>" class="formstyleshort">
			<a href="javascript: foreColor('gsColorAltHighlight');"><img src="../../images/chartblock.gif" width="16" height="13" border="0" alt="<%=dictLanguage("Select_Color")%>"></a>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditAltHighlight" then

	strLookingFor = "gsColorAltHighlight "
	strComments = "'alt light color used throughout the site"
	strNewSetting = trim(request("gsColorAltHighlight"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editWhite" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditWhite">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("White")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsColorWhite" value="<%=gsColorWhite%>" class="formstyleshort">
			<a href="javascript: foreColor('gsColorWhite');"><img src="../../images/chartblock.gif" width="16" height="13" border="0" alt="<%=dictLanguage("Select_Color")%>"></a>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditWhite" then

	strLookingFor = "const gsColorWhite	"
	strComments = "'alt color for white used throughout the site, especially in alternating rows"
	strNewSetting = trim(request("gsColorWhite"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editBlack" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditBlack">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Black")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsColorBlack" value="<%=gsColorBlack%>" class="formstyleshort">
			<a href="javascript: foreColor('gsColorBlack');"><img src="../../images/chartblock.gif" width="16" height="13" border="0" alt="<%=dictLanguage("Select_Color")%>"></a>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditBlack" then

	strLookingFor = "const gsColorBlack	"
	strComments = "'alt color used for black throughout the site, black is so harsh sometimes"
	strNewSetting = trim(request("gsColorBlack"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editSiteName" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditSiteName">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Site_Name")%></b></td>
		<td valign="top">
			<input type="text" name="gsSiteName" value="<%=gsSiteName%>" class="formstylelong">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditSiteName" then

	strLookingFor = "const gsSiteName "
	strComments = ""
	strNewSetting = trim(request("gsSiteName"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editSiteLogo" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditSiteLogo">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Site_Logo")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsSiteLogo" value="<%=gsSiteLogo%>" class="formstylelong">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditSiteLogo" then

	strLookingFor = "const gsSiteLogo "
	strComments = "'stored in the images folder"
	strNewSetting = trim(request("gsSiteLogo"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editAdminEmail" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditAdminEmail">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Admin_Email")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsAdminEmail" value="<%=gsAdminEmail%>" class="formstylelong">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditAdminEmail" then

	strLookingFor = "const gsAdminEmail "
	strComments = ""
	strNewSetting = trim(request("gsAdminEmail"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editMailHost" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditMailHost">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Mail_Host")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsMailHost" value="<%=gsMailHost%>" class="formstylelong">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditMailHost" then

	strLookingFor = "const gsMailHost "
	strComments = ""
	strNewSetting = trim(request("gsMailHost"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editSiteURL" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditSiteURL">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Site_URL")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsSiteURL" value="<%=gsSiteURL%>" class="formstylelong">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditSiteURL" then

	strLookingFor = "const gsSiteURL "
	strComments = "'full site url, including any folders"
	strNewSetting = trim(request("gsSiteURL"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editSiteRoot" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditSiteRoot">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Site_Root")%>:</b></td>
		<td valign="top">
			<input type="text" name="gsSiteRoot" value="<%=gsSiteRoot%>" class="formstylelong">
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
<%if session("permAdminDatabaseSetup") then %>			
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditSiteRoot" then

	strLookingFor = "const gsSiteRoot "
	strComments = "'the place in the site structure where this site is "
	strNewSetting = trim(request("gsSiteRoot"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editMailComponent" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditMailComponent">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Mail_Component")%>:</b></td>
		<td valign="top">
			<select size="1" name="gsMailComponent" class="formstylelong">
				<option value="CDONTS" <%if gsMailComponent="CDONTS" then Response.Write "Selected"%>>CDONTS</option>
				<option value="CDO" <%if gsMailComponent="CDO" then Response.Write "Selected"%>>CDO</option>
				<option value="ASPXP" <%if gsMailComponent="ASPXP" then Response.Write "Selected"%>>ASPXP</option>
				<option value="PersistsASPMail" <%if gsMailComponent="PersistsASPMail" then Response.Write "Selected"%>>PersistsASPMail</option>
				<option value="Jmail" <%if gsMailComponent="Jmail" then Response.Write "Selected"%>>Jmail</option>
				<option value="ServerObjectsASPMail" <%if gsMailComponent="ServerObjectsASPMail" then Response.Write "Selected"%>>ServerObjectsASPMail</option>
			</select>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditMailComponent" then

	strLookingFor = "const gsMailComponent "
	strComments = "'CDONTS is the default component coded into the release "
	strNewSetting = trim(request("gsMailComponent"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editMetaDescription" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditMetaDescription">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Meta_Description")%>:</b></td>
		<td valign="top">
			<textarea name="gsMetaDescription" class="formstylelong" rows="5"><%=gsMetaDescription%></textarea>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditMetaDescription" then

	strLookingFor = "const gsMetaDescription "
	strComments = "'Meta tag description "
	strNewSetting = trim(request("gsMetaDescription"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>

<%elseif strAct = "editMetaKeywords" then %>

<form name="strForm" id="strForm" action="default.asp" method="POST">
<input type="hidden" name="act" value="hndEditMetaKeywords">
<input type="hidden" name="oldSetting" value="<%=strOldSetting%>">
<table cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top"><b><%=dictLanguage("Meta_Keywords")%>:</b></td>
		<td valign="top">
			<textarea name="gsMetaKeywords" class="formstylelong" rows="5"><%=gsMetaKeywords%></textarea>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>	
<%if session("permAdminDatabaseSetup") then %>		
	<tr><td colspan="2" align="center"><input type="submit" name="submit" value="submit" class="formbutton"></td></tr>
<%end if%>
</table>
</form>

<%elseif strAct = "hndEditMetaKeywords" then

	strLookingFor = "const gsMetaKeywords "
	strComments = "'Meta tag keywords "
	strNewSetting = trim(request("gsMetaKeywords"))
	strErrorMessage = EditSetting(strLookingFor, strComments, strNewSetting)
	if strErrorMessage = "" then
		session("strMessage") = strMessage & dictLanguage("The_setting_was_updated") & "<BR>"
	else
		session("strMessage") = strErrorMessage
	end if
	Response.Redirect "default.asp?act=hnd_thnk"	%>



<%elseif strAct = "hnd_thnk" then %>
<p align="Center"><%=session("strMessage")%></p>
	<%session("strMessage") = ""%>

<%end if%>

</td>
</tr>
</table>

<p align="center"><a href=".."><%=dictLanguage("Return_Admin_Home")%></a></p>

<%
Function EditSetting(strToFind, strToFindComments, strForReplace)
	strToFindSize = len(strToFind)
	dim MyFileLines(200)
	dim strMessage
	On Error Resume Next
	'Do some FSO magic and edit the siteConfig page
	strSiteConfig = Server.MapPath("../../includes/SiteConfig.asp")
	set FSO = Server.CreateObject("Scripting.FileSystemObject")
	if err.number <> 0 then
		strMessage = strMessage & "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR><BR>" 
		err.clear
	End if
		
	if not FSO.FileExists(strSiteConfig) then
		set FSO = nothing
		strMessage = strMessage & dictLanguage("Site_Config_Not_Found") & "<BR><BR>"
	else
		Set MyFile = FSO.GetFile(strSiteConfig)
		if err.number <> 0 then
			strMessage = strMessage & "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR><BR>" 
			err.clear
		End if		
		set MyFileTS = MyFile.OpenAsTextStream(1,-2) 'open for reading
		if err.number <> 0 then
			strMessage = strMessage & "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR><BR>" 
			err.clear
		End if	
		i = 0
		do while not MyFileTS.atEndOfStream
			MyFileLines(i) = MyFileTS.ReadLine
			if left(MyFileLines(i),strToFindSize) = strToFind then
				'Response.Write "FOUND IT<BR><BR>"
				if strForReplace = "TRUE" or strForReplace = "FALSE" then
					MyFileLines(i) = strToFind & " = " & strForReplace & "    " & strToFindComments
				else
					MyFileLines(i) = strToFind & " = """ & strForReplace & """    " & strToFindComments
				end if
			end if																											
			i = i + 1
		loop
		MyFileTS.close
		if err.number <> 0 then
			strMessage = strMessage & "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR><BR>" 
			err.clear
		End if			
		set MyFileTS = MyFile.OpenAsTextStream(2,-2) 'open for writing
		if err.number <> 0 then
			strMessage = strMessage & "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR><BR>" 
			err.clear
		End if	
		for j = 0 to i-1
			MyFileTS.WriteLine MyFileLines(j)
		next
		set MyFileTS = nothing
		Set MyFile = nothing
		set FSO = nothing
		if err.number <> 0 then
			strMessage = strMessage & "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR><BR>" 
			err.clear
		End if	
		'Response.Write "<p align=""left"">"
		'for j = 1 to i-2
		'	Response.Write "" & MyFileLines(j) & "<br>"
		'next 
		'Response.Write "</p>"
	end if
	if strMessage <> "" then 
		strMessage = "Error Occured:<BR><BR>" & strMessage
	end if
	EditSetting = strMessage
End Function	
%>

<!--#include file="../../includes/main_page_close.asp"-->
