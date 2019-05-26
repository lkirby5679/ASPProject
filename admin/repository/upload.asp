<!--#include file="../../includes/SiteConfig.asp"-->
<!--#include file="../../includes/connection_open.asp"-->
<%
Response.Buffer = true
Function BuildUpload(RequestBin)
     'Get the boundary
     PosBeg = 1
     PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
     boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
     boundaryPos = InstrB(1,RequestBin,boundary)
     'Get all data inside the boundaries
     Do until (boundaryPos=InstrB(RequestBin,boundary & getByteString("--")))
          'Members variable of objects are put in a dictionary object
          Dim UploadControl
          Set UploadControl = CreateObject("Scripting.Dictionary")
          'Get an object name
          Pos = InstrB(BoundaryPos,RequestBin,getByteString("Content-Disposition"))
          Pos = InstrB(Pos,RequestBin,getByteString("name="))
          PosBeg = Pos+6
          PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
          Name = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
          PosFile = InstrB(BoundaryPos,RequestBin,getByteString("filename="))
          PosBound = InstrB(PosEnd,RequestBin,boundary)
          'Test if object is of file type
          If PosFile<>0 AND (PosFile<PosBound) Then
               'Get Filename, content-type and content of file
               PosBeg = PosFile + 10
               PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(34)))
               FileName = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
               'Add filename to dictionary object
               UploadControl.Add "FileName", FileName
               Pos = InstrB(PosEnd,RequestBin,getByteString("Content-Type:"))
               PosBeg = Pos+14
               PosEnd = InstrB(PosBeg,RequestBin,getByteString(chr(13)))
               'Add content-type to dictionary object
               ContentType = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
               UploadControl.Add "ContentType",ContentType
               'Get content of object
               PosBeg = PosEnd+4
               PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
               Value = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
               Else
               'Get content of object
               Pos = InstrB(Pos,RequestBin,getByteString(chr(13)))
               PosBeg = Pos+4
               PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
               Value = getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
          End If
          UploadControl.Add "Value" , Value
          UploadRequest.Add name, UploadControl
          BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
     Loop
End Function

Function getByteString(StringStr)
     For i = 1 to Len(StringStr)
          char = Mid(StringStr,i,1)
          getByteString = getByteString & chrB(AscB(char))
     Next
End Function

Function getString(StringBin)
     getString =""
     For intCount = 1 to LenB(StringBin)
          getString = getString & chr(AscB(MidB(StringBin,intCount,1)))
     Next
End Function

Response.Clear
byteCount = Request.TotalBytes

RequestBin = Request.BinaryRead(byteCount)

Set UploadRequest = CreateObject("Scripting.Dictionary")

BuildUpload(RequestBin)

URL = UploadRequest.Item("url").Item("Value")
Description = UploadRequest.Item("description").Item("Value")
Folder = UploadRequest.Item("folder").Item("Value")
byEmployee = UploadRequest.Item("byEmployee").Item("Value")

If UploadRequest.Item("blob").Item("Value") <> "" Then
	contentType = UploadRequest.Item("blob").Item("ContentType")
    filepathname = UploadRequest.Item("blob").Item("FileName")
    filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
    filenameShow = filename
    FolderName = UploadRequest.Item("where").Item("Value")
    Path = Server.MapPath(FolderName)
    ToFolder = Path
    value = UploadRequest.Item("blob").Item("Value")
    filename = ToFolder & "\" & filename
    Set MyFileObject = Server.CreateObject("Scripting.FileSystemObject")
	if not MyFileObject.FileExists(filename) then
	    Set objFile = MyFileObject.CreateTextFile(filename)
	    For i = 1 to LenB(value)
			objFile.Write chr(AscB(MidB(value,i,1)))
	    Next
		objFile.Close
	    Set objFile = Nothing
	else
		session("msg") = dictLanguage("Error_FileExists")
		Response.Redirect ("default.asp")
	end if
    Set MyFileObject = Nothing
End If
Set UploadRequest = Nothing
	
if Folder = "" then 
	Folder = 5
end if
if (URL<>"" and URL<>"http://") or fileNameShow<>"" then
	sql = sql_InsertDocument(SQLEncode(fileNameShow), Folder, SQLEncode(URL), _
		SQLEncode(Description), byEmployee, date()) 
	'response.write sql
	Call DoSQL(sql)
else
	session("msg") = dictLanguage("Error_NoFileInput")
	Response.Redirect ("default.asp")
end if
%>
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

<!--#include file="../../includes/connection_close.asp"-->

<%session("msg") = dictLanguage("Document_Uploaded")
response.redirect ("default.asp")%>