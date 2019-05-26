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

<%strType = request("strType")
  if strType="excel" then
	Response.Contenttype="application/vnd.ms-excel"
	Response.AddHeader "content-disposition","attachment;filename=news.xls"
  elseif strType="cdtf" then
 	Response.Contenttype="text"
	Response.AddHeader "content-disposition","attachment;filename=news.txt" 
  end if %>

<!-- #include file="../../includes/siteconfig.asp"-->
<!-- #include file="../../includes/connection_open.asp"-->

<%
if strType <> "cdtf" then
	Response.Write "<html><head></head><body>"
end if
' this default sql statement just gets all the columns and items in the database.  if you 
' would like you could change the sql statement to do joins, etc, and return the values
' you want returned. 
sql = sql_GetAllNews()
Call RunSQL(sql, rs)
if not rs.eof then
	if strType <> "cdtf" then 
		Response.Write "<table cellpadding=1 cellspacing=0 border=1><tr>"
	end if
	for i = 0 to rs.fields.count-1
		if strType <> "cdtf" then 
			Response.Write "<td><font size=2><b>" 
		else 
			Response.Write """"
		end if
		Response.write trim(rs.fields(i).name)
		if strType <> "cdtf" then 
			Response.Write "</b></font></td>"
		else 
			Response.Write """"
			if i <> rs.fields.count - 1 then
				Response.Write ", "
			end if
		end if
	next
	if strType <> "cdtf" then 
		Response.Write "</tr>"
	else 
		Response.Write VBCRLF
	end if
	while not rs.eof
		if strType <> "cdtf" then 
			Response.Write "<tr>"
		end if
		for i = 0 to rs.fields.count-1
			if strType <> "cdtf" then 
				Response.Write "<td><font size=2>"
			else
				Response.Write """"
			end if
			Response.write trim(rs.fields(i).value)
			if strType <> "cdtf" then 
				Response.Write "</font></td>"
			else 
				Response.Write """"
				if i <> rs.fields.count - 1 then
					Response.Write ", "
				end if
			end if
		next		
		if strType <> "cdtf" then 
			Response.Write "</tr>"
		else
			Response.Write VBCRLF
		end if
		rs.movenext
	wend
else
	Response.Write dictLanguage("No_records_exist")
end if
if strType <> "cdtf" then
	Response.Write "</body></html>"
end if 
%>