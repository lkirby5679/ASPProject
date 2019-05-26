<%@ Language=VBScript %>
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
<% 'with the addition of the MySQL and PostgreSQL options to the SiteConfig file
   'this page needs to be rewritten - Notes: dsh 8/10/2002%>

<!--#include file="../includes/SiteConfig.asp"-->

<html>
<head>
</head>

<script>
<!--
window.resizeTo(750, 550)			
// -->
</script>

<!--#include file="../includes/style.asp"-->

<% mode = trim(request("mode"))%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<html>
<head>
	<meta http-equiv="Pragma" content="no-cache">
	<title>intranet login</title>
</head>

<body>

<table width="100%" border="0" cellpadding="2" cellspacing="2" align="center">
	<tr>
		<td valign="top">
			<table border="0" cellpadding="2" cellspacing="0" width="100%" bgColor="<%=gsColorHighlight%>">
				<tr>
					<td valign="top"><a href="<%=gsSiteRoot%>"><img src="<%=gsSiteRoot%>images/tilogo.gif" alt="Transworld Interactive Project" border="0"></a></td>
					<td align="right"><a href="<%=gsSiteRoot%>" target="_blank"><img src="<%=gsSiteRoot%>images/poweredbydiggersolutionsv11.gif" border="0" width="160" height="40" alt="powered by Transworld Interactive Projects"></a></td>
				</tr>
			</table>
			<table border="0" cellpadding="2" cellspacing="2" align="center" width="500">
				<tr bgcolor="<%=gsColorHighlight%>"><td align="center" class="homeheader" colspan="2">Database Setup</td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>

<% if mode = "" then %>
				<form action="dbsetup.asp" method="post">
				<input type="hidden" name="mode" value="selectdb">
				<tr>
					<td colspan="2"><%=dictLanguage("dbsetup1")%></td>
				</tr>
				<tr>
					<td class="bolddark" valign="top" nowrap>Type of database:</td>
					<td valign="top">
						<input type="radio" name="strDBType" value="access" Checked> MS Access<br>
						<input type="radio" name="strDBType" value="sqlserver"> MS SQL Server 7.0 +<br>
						<input type="radio" name="strDBType" value="other"> <%=dictLanguage("Other")%>
					</td>
				</tr> 
				<tr><td colspan="2" align="center"><input type="submit" name="submit" value="Continue" class="formButton"></td></tr>
				</form>
<% end if %>
<% if mode = "selectdb" then 
		if trim(request("strDBType"))= "other" then %>
				<tr><td colspan="2"><%=dictLanguage("dbsetup2")%></td></tr>
<%		elseif trim(request("strDBType"))="access" then %>
				<form action="dbsetup.asp" method="post" id="form1" name="form1">
				<input type="hidden" name="mode" value="enterdata">
				<input type="hidden" name="strDBType" value="access">
				<tr>
					<td colspan="2"><%=dictLanguage("dbsetup3")%></td>
				</tr>
				<tr>
					<td class="bolddark" valign="top" nowrap><%=dictLanguage("Access_File_Name")%>:</td>
					<td valign="top"><input type="text" name="strDBFileName" value="intranet.mdb" class="formstyleLong"></td>
				</tr> 
				<tr><td colspan="2" align="center"><input type="submit" name="submit" value="Continue" class="formButton"></td></tr>
				</form>
<%		elseif trim(request("strDBType"))="sqlserver" then %>
				<form action="dbsetup.asp" method="post" id="form1" name="form1">
				<input type="hidden" name="mode" value="enterdata">
				<input type="hidden" name="strDBType" value="sqlserver">
				<tr>
					<td colspan="2"><%=dictLanguage("dbsetup4")%></td>
				</tr>
				<tr>
					<td class="bolddark" valign="top" nowrap><%=dictLanguage("Data_Source")%>:</td>
					<td valign="top"><input type="text" name="strDBDataSource" value="Whitetip" class="formstyleLong"></td>
				</tr> 
				<tr>
					<td class="bolddark" valign="top" nowrap><%=dictLanguage("Initial_Catalog")%>:</td>
					<td valign="top"><input type="text" name="strDBInitialCatalog" value="intranet" class="formstyleLong"></td>
				</tr> 				
				<tr>
					<td class="bolddark" valign="top" nowrap><%=dictLanguage("User_Name")%>:</td>
					<td valign="top"><input type="text" name="strDBUserName" value="intranet" class="formstyleLong"></td>
				</tr> 
				<tr>
					<td class="bolddark" valign="top" nowrap><%=dictLanguage("Password")%>:</td>
					<td valign="top"><input type="text" name="strDBPassword" value="digger" class="formstyleLong"></td>
				</tr> 
				<tr><td colspan="2" align="center"><input type="submit" name="submit" value="Continue" class="formButton"></td></tr>
				</form>								
<%		end if %>
<% elseif mode = "enterdata" then 
		dim MyFileLines(150)
		if trim(request("strDBType"))="access" then 
			strDBFileName = trim(request("strDBFileName"))
			'Do some FSO magic and edit the siteConfig page
			strSiteConfig = Server.MapPath("../includes/SiteConfig.asp")
			set FSO = Server.CreateObject("Scripting.FileSystemObject")
			if not FSO.FileExists(strSiteConfig) then
				set FSO = nothing
				Response.Write "<tr><td colspan=""2"">The SiteConfig.asp file was not where it was expected.  Please check that the site has been configured properly.</td></tr>"
			else
				Set MyFile = FSO.GetFile(strSiteConfig)
				set MyFileTS = MyFile.OpenAsTextStream(1,-2) 'open for reading
				i = 0
				boolHitIt = FALSE
				boolDoneIt = FALSE
				do while not MyFileTS.atEndOfStream
					MyFileLines(i) = MyFileTS.ReadLine
					if left(MyFileLines(i),16)="'if using access" then
						boolHitIt = TRUE
					end if
					if left(MyFileLines(i),18) = "'DB_DATASOURCENAME" or left(MyFileLines(i),17) = "DB_DATASOURCENAME" then
						MyFileLines(i) = "DB_DATASOURCENAME = """ & strDBFileName & """ 'name of mdb file - should be stored in the data folder outside of the htdocs folder"
					end if
					if left(MyFileLines(i),14) = "'DB_DATASOURCE" or left(MyFileLines(i),13) = "DB_DATASOURCE" then
						if boolHitIt then
							MyFileLines(i) = Replace(MyFileLines(i),"'DB_DATASOURCE","DB_DATASOURCE",1,1)
						elseif left(MyFileLines(i),13) = "DB_DATASOURCE" then
							MyFileLines(i) = Replace(MyFileLines(i),"DB_DATASOURCE","'DB_DATASOURCE")
						end if
					end if
					if left(MyFileLines(i),17) = "'DB_DATEDELIMITER" or left(MyFileLines(i),16) = "DB_DATEDELIMITER" then
						if boolHitIt then
							MyFileLines(i) = Replace(MyFileLines(i),"'DB_DATEDELIMITER","DB_DATEDELIMITER",1,1)
						elseif left(MyFileLines(i),16) = "DB_DATEDELIMITER" then
							MyFileLines(i) = Replace(MyFileLines(i),"DB_DATEDELIMITER","'DB_DATEDELIMITER",1,1)
						end if
					end if	
					if left(MyFileLines(i),14) = "'DB_BOOLACCESS" or left(MyFileLines(i),13) = "DB_BOOLACCESS" then
						if boolHitIt then
							MyFileLines(i) = Replace(MyFileLines(i),"'DB_BOOLACCESS","DB_BOOLACCESS",1,1)
						elseif left(MyFileLines(i),13) = "DB_BOOLACCESS" then
							MyFileLines(i) = Replace(MyFileLines(i),"DB_BOOLACCESS","'DB_BOOLACCESS",1,1)
						end if
					end if		
					if left(MyFileLines(i),20) = "'DB_CONNECTIONSTRING" or left(MyFileLines(i),19) = "DB_CONNECTIONSTRING" then
						if boolHitIt then
							MyFileLines(i) = Replace(MyFileLines(i),"'DB_CONNECTIONSTRING","DB_CONNECTIONSTRING",1,1)
						elseif left(MyFileLines(i),19) = "DB_CONNECTIONSTRING" then
							MyFileLines(i) = Replace(MyFileLines(i),"DB_CONNECTIONSTRING","'DB_CONNECTIONSTRING",1,1)
						end if
					end if
					if left(MyFileLines(i),17) = "DB_INITIALCATALOG" then
							MyFileLines(i) = Replace(MyFileLines(i),"DB_INITIALCATALOG","'DB_INITIALCATALOG",1,1)
					end if	
					if left(MyFileLines(i),11) = "DB_USERNAME" then
							MyFileLines(i) = Replace(MyFileLines(i),"DB_USERNAME","'DB_USERNAME",1,1)
					end if
					if left(MyFileLines(i),11) = "DB_PASSWORD" then
							MyFileLines(i) = Replace(MyFileLines(i),"DB_PASSWORD","'DB_PASSWORD",1,1)
					end if																												
					i = i + 1
				loop
				MyFileTS.close
				set MyFileTS = MyFile.OpenAsTextStream(2,-2) 'open for writing
				for j = 0 to i-1
					MyFileTS.WriteLine MyFileLines(j)
				next
				set MyFileTS = nothing
				Set MyFile = nothing
				set FSO = nothing
				'for j = 1 to i-2
				'	Response.Write "<tr><td colspan=""2"">" & MyFileLines(j) & "</td></tr>"
				'next 
				Response.Redirect "dbsetup.asp?mode=done"	
			end if									
		elseif trim(request("strDBType")) = "sqlserver" then
			strDBDataSource = trim(request("strDBDataSource"))
			strDBInitialCatalog = trim(request("strDBInitialCatalog"))
			strDBUserName = trim(request("strDBUserName"))
			strDBPassword = trim(request("strDBPassword"))
			'Do some FSO magic and edit the siteConfig page
			strSiteConfig = Server.MapPath("../includes/SiteConfig.asp")
			set FSO = Server.CreateObject("Scripting.FileSystemObject")
			if not FSO.FileExists(strSiteConfig) then
				set FSO = nothing
				Response.Write "<tr><td colspan=""2"">The SiteConfig.asp file was not where it was expected.  Please check that the site has been configured properly.</td></tr>"
			else
				Set MyFile = FSO.GetFile(strSiteConfig)
				set MyFileTS = MyFile.OpenAsTextStream(1,-2) 'open for reading
				i = 0
				boolHitIt = FALSE
				do while not MyFileTS.atEndOfStream
					MyFileLines(i) = MyFileTS.ReadLine
					if left(MyFileLines(i),16)="'if using access" then
						boolHitIt = TRUE
					end if
					if left(MyFileLines(i),17) = "DB_DATASOURCENAME" then
						MyFileLines(i) = Replace(MyFileLines(i),"DB_DATASOURCENAME","'DB_DATASOURCENAME",1,1)
					end if
					if (left(MyFileLines(i),14) = "'DB_DATASOURCE" or left(MyFileLines(i),13) = "DB_DATASOURCE") and left(MyFileLines(i),18)<>"'DB_DATASOURCENAME" and left(MyFileLines(i),17)<>"DB_DATASOURCENAME" then
						if boolHitIt then
							if left(MyFileLines(i),13) = "DB_DATASOURCE" then
								MyFileLines(i) = Replace(MyFileLines(i),"DB_DATASOURCE","'DB_DATASOURCE",1,1)
							end if
						else
							MyFileLines(i) = "DB_DATASOURCE = """ & strDBDataSource & """ 'Data Source or Server name"
						end if
					end if
					if left(MyFileLines(i),17) = "'DB_DATEDELIMITER" or left(MyFileLines(i),16) = "DB_DATEDELIMITER" then
						if boolHitIt then
							if left(MyFileLines(i),16) = "DB_DATEDELIMITER" then
								MyFileLines(i) = Replace(MyFileLines(i),"DB_DATEDELIMITER","'DB_DATEDELIMITER",1,1)
							end if
						else
							MyFileLines(i) = Replace(MyFileLines(i),"'DB_DATEDELIMITER","DB_DATEDELIMITER",1,1)
						end if
					end if	
					if left(MyFileLines(i),14) = "'DB_BOOLACCESS" or left(MyFileLines(i),13) = "DB_BOOLACCESS" then
						if boolHitIt then
							if left(MyFileLines(i),13) = "DB_BOOLACCESS" then
								MyFileLines(i) = Replace(MyFileLines(i),"DB_BOOLACCESS","'DB_BOOLACCESS",1,1)
							end if
						else
							MyFileLines(i) = Replace(MyFileLines(i),"'DB_BOOLACCESS","DB_BOOLACCESS",1,1)
						end if
					end if		
					if left(MyFileLines(i),20) = "'DB_CONNECTIONSTRING" or left(MyFileLines(i),19) = "DB_CONNECTIONSTRING" then
						if boolHitIt then
							if left(MyFileLines(i),19) = "DB_CONNECTIONSTRING" then
								MyFileLines(i) = Replace(MyFileLines(i),"DB_CONNECTIONSTRING","'DB_CONNECTIONSTRING",1,1)
							end if
						else
							MyFileLines(i) = Replace(MyFileLines(i),"'DB_CONNECTIONSTRING","DB_CONNECTIONSTRING",1,1)
						end if
					end if
					if left(MyFileLines(i),18) = "'DB_INITIALCATALOG" or left(MyFileLines(i),17) = "DB_INITIALCATALOG" then
							MyFileLines(i) = "DB_INITIALCATALOG = """ & strDBInitialCatalog & """ 'initial database to log into"
					end if	
					if left(MyFileLines(i),12) = "'DB_USERNAME" or left(MyFileLines(i),11) = "DB_USERNAME" then
							MyFileLines(i) = "DB_USERNAME = """ & strDBUserName & """"
					end if
					if left(MyFileLines(i),12) = "'DB_PASSWORD" or left(MyFileLines(i),11) = "DB_PASSWORD" then
							MyFileLines(i) = "DB_PASSWORD = """ & strDBPassword & """"
					end if																												
					i = i + 1
				loop
				MyFileTS.close
				set MyFileTS = MyFile.OpenAsTextStream(2,-2) 'open for writing
				for j = 0 to i-1
					MyFileTS.WriteLine MyFileLines(j)
				next
				set MyFileTS = nothing
				Set MyFile = nothing
				set FSO = nothing
				'for j = 1 to i-2
				'	Response.Write "<tr><td colspan=""2"">" & MyFileLines(j) & "</td></tr>"
				'next 
				Response.Redirect "dbsetup.asp?mode=done"	
			end if				
		end if %>
<% elseif mode = "done" then %>
				<tr><td colspan="2">
						<%=dictLanguage("dbsetup5")%>
						<br><br>
						<div align="center">
						<a href="<%=gsSiteRoot%>" target="_blank"><%=dictLanguage("Proceed_To_Login")%></a>
						</div>
				</td></tr>
			
<% end if %>

			</table>
		</td>
	</tr>
</table>
<p align="center"><a href="javascript: window.close()"><%=dictLanguage("Close_This_Window")%></a></p>
</body>
</html>

