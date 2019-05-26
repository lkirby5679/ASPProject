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
<!--#include file="SiteSQL.asp"-->
<%
'** CONNECTION STRING
Dim conn
strConn = DB_CONNECTIONSTRING	
set conn=server.createobject("adodb.connection")
conn.Open strConn

Private function RunSQL(ByVal sql,ByRef myRS)
	'Response.Write sql & "<BR>"
	On Error Resume Next
	if lcase(left(trim(sql),6)) = "select" then
			set myRS = server.CreateObject("adodb.recordset")
			myRS.open sql,conn,3,3
	else
		select case lcase(left(trim(sql),6))
			case "update", "delete"
				if instr(lcase(sql),"where") = 0 then
					response.write "Dork, you tried to run a "& lcase(left(trim(sql),6)) &" query without the where clause!"
					boolDoNotRunQuery = true
					response.end
				end if
		end select
		if boolDoNotRunQuery <> true then
			set myRS = conn.Execute(sql)
		end if
	end if
	if err.number <> 0 then
		Response.Clear
		Response.Write "Error Occured:<BR><BR>"
		Response.Write "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR>" 
		Response.Write "SQL = " & sql & "<BR>"	
		Response.End
	End if
	On Error Goto 0	
end Function

Private function DoSQL(ByVal sql)
	'Response.write sql & "<BR>"
	On Error Resume Next
	boolDoNotRunQuery = FALSE
	actionWord = left(trim(sql),6)
	select case lcase(actionWord)
		case "update", "delete"
			if instr(lcase(sql),"where") = 0 then
				response.write "Dork, you tried to run a "& lcase(left(trim(sql),6)) &" query without the where clause!"
				boolDoNotRunQuery = true
				response.end
			end if
	end select
	if boolDoNotRunQuery <> true then
		conn.Execute(sql)
	end if
	if err.number <> 0 then
		Response.Clear
		Response.Write "Error Occured:<BR><BR>"
		Response.Write "Error # " & CStr(Err.Number) & " " & Err.Description & "<BR>" 
		Response.Write "SQL = " & sql & "<BR>"	
		Response.End
	End if
	On Error Goto 0	
end Function

function SQLEncode(strText)
   strText = trim(strText)
   if strText <> "" and isNull(strText) = False then
     strText = replace(strText,"'","''")
   end if
   SQLEncode = strText
end function

function PrepareBit(strBit)
	if strBit <> "" and isNumeric(strBit) then
		if strBit = 1 or strBit=-1 or strBit then
			if DB_BOOLACCESS then
				strBit = -1
			else
				strBit = 1
			end if
		else 
			strBit = 0
		end if
	else
		strBit = 0
	end if
	PrepareBit = strBit
end function

Function MediumDate(str)
    Dim aDay
    Dim aMonth
    Dim aYear

    aDay = Day(str)
    'aMonth = Monthname(Month(str),TRUE)
    aMonth = Month(str)
    Select Case aMonth
		Case 1
			aMonth = "Jan"
		Case 2
			aMonth = "Feb"
		Case 3
			aMonth = "Mar"
		Case 4
			aMonth = "Apr"
		Case 5
			aMonth = "May"
		Case 6
			aMonth = "Jun"
		Case 7
			aMonth = "Jul"
		Case 8
			aMonth = "Aug"
		Case 9
			aMonth = "Sep"
		Case 10
			aMonth = "Oct"
		Case 11
			aMonth = "Nov"
		Case 12
			aMonth = "Dec"
		Case Else
			aMonth = "Jan"
    End Select
    aYear = Year(str)
	
	if DB_BOOLMYSQL then
		MediumDate = aYear & "-" & Month(str) & "-" & aDay
	else	
		MediumDate = aDay & "-" & aMonth & "-" & aYear
	end if
End Function

Function DecimalCommaToPeriod(nmbr)
	if instr(nmbr, ",") then
		nmbr = replace(nmbr, ",", ".")
	end if
	DecimalCommaToPeriod = nmbr
End Function

%>