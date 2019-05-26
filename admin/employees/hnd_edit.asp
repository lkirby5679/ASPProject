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
if Err <> 0 then
	strMessage = dictLanguage("Upload_Error_1")
	Err.clear
end if 	

BuildUpload(RequestBin)
     
strAct				= UploadRequest.Item("act").Item("Value")
if strAct = "edit" then
	strID				= UploadRequest.Item("empID").Item("Value")
end if
empName				= SQLEncode(UploadRequest.Item("empName").Item("Value"))
empStartDate		= SQLEncode(UploadRequest.Item("empStartDate").Item("Value"))
empLogin			= SQLEncode(UploadRequest.Item("empLogin").Item("Value"))
empTitle			= SQLEncode(UploadRequest.Item("empTitle").Item("Value"))
empEmail			= SQLEncode(UploadRequest.Item("empEmail").Item("Value"))
empEmailAlt			= SQLEncode(UploadRequest.Item("empEmailAlt").Item("Value"))
empCommissionRate	= UploadRequest.Item("empCommissionRate").Item("Value")
if not isNumeric(empCommissionRate) then
	empCommissionRate = 0
else
	empCommissionRate = int(empCommissionRate)
end if
empSalesGoal		= UploadRequest.Item("empSalesGoal").Item("Value")
if not isNumeric(empSalesGoal) then
	empSalesGoal = 0
else
	empSalesGoal = int(empSalesGoal)
end if
empProductionGoal	= UploadRequest.Item("empProductionGoal").Item("Value")
if not isNumeric(empProductionGoal) then
	empProductionGoal = 0
else
	empProductionGoal = int(empProductionGoal)
end if
empPassword			= SQLEncode(UploadRequest.Item("empPassword").Item("Value"))
if UploadRequest.Exists("empActive") then
	empActive		= UploadRequest.Item("empActive").Item("Value")
	if empActive <> 1 then
		empActive = 0
	end if
else
	empActive = 0
end if
if UploadRequest.Exists("empHourly") then
	empHourly			= UploadRequest.Item("empHourly").Item("Value") 
	if empHourly <> 1 then
		empHourly = 0
	end if		
else
	empHourly = 0
end if
empType				= UploadRequest.Item("empType").Item("Value")
if empType = 0 then
	empType = 1
end if
empReportsTo		= UploadRequest.Item("empReportsTo").Item("Value")
if empReportsTo = 0 then
	empReportsTo = 1
end if
empDepartment		= UploadRequest.Item("empDepartment").Item("Value")
empBirthDate		= UploadRequest.Item("empBirthDate").Item("Value")
empHomePhone		= SQLEncode(UploadRequest.Item("empHomePhone").Item("Value"))
empMobilePhone		= SQLEncode(UploadRequest.Item("empMobilePhone").Item("Value"))
empWorkPhone		= SQLEncode(UploadRequest.Item("empWorkPhone").Item("Value"))
empWorkPhoneExt		= SQLEncode(UploadRequest.Item("empWorkPhoneExt").Item("Value"))
empVoiceMail		= SQLEncode(UploadRequest.Item("empVoiceMail").Item("Value"))
empHomeStreet1		= SQLEncode(UploadRequest.Item("empHomeStreet1").Item("Value"))
empHomeStreet2		= SQLEncode(UploadRequest.Item("empHomeStreet2").Item("Value"))
empHomeCity			= SQLEncode(UploadRequest.Item("empHomeCity").Item("Value"))
empHomeState		= SQLEncode(UploadRequest.Item("empHomeState").Item("Value"))
empHomeZip			= SQLEncode(UploadRequest.Item("empHomeZip").Item("Value"))
empIMName			= SQLEncode(UploadRequest.Item("empIMName").Item("Value"))

empNewFile			= SQLEncode(UploadRequest.Item("empImage").Item("FileName"))
if strAct = "edit" then
	empOldFile			= SQLEncode(UploadRequest.Item("empImageOld").Item("Value"))     
end if

'On Error Resume Next     
If empNewFile <> "" Then
    contentType = UploadRequest.Item("empImage").Item("ContentType")
    filepathname = UploadRequest.Item("empImage").Item("FileName")
    filename = Right(filepathname,Len(filepathname)-InstrRev(filepathname,"\"))
    empNewFile = filename
    FolderName = Server.MapPath("employees/images/")
    filename = FolderName & "\" & filename
    'Response.Write filename & "<BR>"
    value = UploadRequest.Item("empImage").Item("Value")
    Set MyFileObject = Server.CreateObject("Scripting.FileSystemObject")
    Set objFile = MyFileObject.CreateTextFile(filename)
    For i = 1 to LenB(value)
         objFile.Write chr(AscB(MidB(value,i,1)))
    Next
    objFile.Close
    Set objFile = Nothing
    Set MyFileObject = Nothing
else
	empNewFile = empOldFile	 
End If
Set UploadRequest = Nothing

if (Err <> 0) then
	Err.clear
	strMessage =  dictLanguage("Upload_Error_2") & strMessage
	strNewFile = ""
end if
On Error Goto 0	
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
<!-- #include file="../../includes/siteconfig.asp"--> 
<!-- #include file="../../includes/connection_open.asp" -->

<%
if (strAct = "add") then
	sql = sql_InsertEmployee( _
			empName, empStartDate, empLogin, empTitle, empEmail, empEmailAlt, _
			empCommissionRate, empSalesGoal, empProductionGoal, empPassword, _
			empActive, empHourly, empType, empReportsTo, empDepartment, _
			empBirthDate, empHomePhone, empMobilePhone, empWorkPhone, _
			empWorkPhoneExt, empVoiceMail, empHomeStreet1, empHomeStreet2, _
			empHomeCity, empHomeState, empHomeZip, empIMName, empNewFile)
	'Response.Write sql & "<BR>"
	Call DoSQL(sql)

	'Check to see if application variables have values
	'if they don't have a value then give them all a value of 0
	if Application("blnpermAdminAll") = "" then
		Application("blnpermAdminAll")				= 0
		Application("blnpermAdmin")					= 0
		Application("blnpermAdminCalendar")			= 0
		Application("blnpermAdminDatabaseSetup")	= 0
		Application("blnpermAdminEmployees")		= 0
		Application("blnpermAdminEmployeesPerms")	= 0
		Application("blnpermAdminFileRepository")	= 0
		Application("blnpermAdminForum")			= 0
		Application("blnpermAdminNews")				= 0
		Application("blnpermAdminResources")		= 0
		Application("blnpermAdminSurveys")			= 0
		Application("blnpermAdminThoughts")			= 0
		Application("blnpermClientsAdd")			= 0
		Application("blnpermClientsEdit")			= 0
		Application("blnpermClientsDelete")			= 0
		Application("blnpermProjectsAdd")			= 0
		Application("blnpermProjectsEdit")			= 0
		Application("blnpermProjectsDelete")		= 0
		Application("blnpermTasksAdd")				= 0
		Application("blnpermTasksEdit")				= 0
		Application("blnpermTasksDelete")			= 0
		Application("blnpermTimecardsAdd")			= 0
		Application("blnpermTimecardsEdit")			= 0
		Application("blnpermTimecardsDelete")		= 0
		Application("blnpermTimesheetsEdit")		= 0
		Application("blnpermForumAdd")				= 0
		Application("blnpermRepositoryAdd")			= 0
		Application("blnpermPTOAdmin")			    = 0
	end if

	sql = sql_SelectLatestEmployeeID
	'Response.Write sql & "<BR>"
	'response.end
	Call RunSQL(sql, rsLargestEmployeeID)
	if rsLargestEmployeeID("maxid") <> NULL or rsLargestEmployeeID("maxid") <> "" then
		if isNumeric(rsLargestEmployeeID("maxid")) then
			strEmpID				= (rsLargestEmployeeID("maxid"))
			permAll 				= Application("blnpermAdminAll")
			permAdmin 				= Application("blnpermAdmin")
			permAdminCalendar		= Application("blnpermAdminCalendar")
		 	permAdminDatabaseSetup	= Application("blnpermAdminDatabaseSetup")
		 	permAdminEmployees		= Application("blnpermAdminEmployees")
		 	permAdminEmployeesPerms	= Application("blnpermAdminEmployeesPerms")
		 	permAdminFileRepository	= Application("blnpermAdminFileRepository")
			permAdminForum			= Application("blnpermAdminForum")
		 	permAdminNews			= Application("blnpermAdminNews")
		 	permAdminResources		= Application("blnpermAdminResources")
		 	permAdminSurveys		= Application("blnpermAdminSurveys")
		 	permAdminThoughts		= Application("blnpermAdminThoughts")
			permClientsAdd			= Application("blnpermClientsAdd")
		 	permClientsEdit			= Application("blnpermClientsEdit")
		 	permClientsDelete		= Application("blnpermClientsDelete")
			permProjectsAdd			= Application("blnpermProjectsAdd")
		 	permProjectsEdit		= Application("blnpermProjectsEdit")
		 	permProjectsDelete		= Application("blnpermProjectsDelete")
		 	permTasksAdd			= Application("blnpermTasksAdd")
		 	permTasksEdit			= Application("blnpermTasksEdit")
		 	permTasksDelete			= Application("blnpermTasksDelete")
		 	permTimecardsAdd		= Application("blnpermTimecardsAdd")
		 	permTimecardsEdit		= Application("blnpermTimecardsEdit")
		 	permTimecardsDelete		= Application("blnpermTimecardsDelete")
		 	permTimesheetsEdit		= Application("blnpermTimesheetsEdit")
		 	permForumAdd			= Application("blnpermForumAdd")
		 	permRepositoryAdd		= Application("blnpermRepositoryAdd")
		 	permPTOAdmin			= Application("blnpermPTOAdmin")
			sql = sql_UpdatePermissions(strEmpID, permAll, permAdmin, permAdminCalendar, _
				permAdminDatabaseSetup, permAdminEmployees, permAdminEmployeesPerms, _
				permAdminFileRepository, permAdminForum, permAdminNews, permAdminResources, _
				permAdminSurveys, permAdminThoughts, permClientsAdd, permClientsEdit, _
				permClientsDelete, permProjectsAdd, permProjectsEdit, permProjectsDelete, _
				permTasksAdd, permTasksEdit, permTasksDelete, permTimecardsAdd, _
				permTimecardsEdit, permTimecardsDelete, permTimesheetsEdit, permForumAdd, _
				permRepositoryAdd, permPTOAdmin)
			Call DoSQL(sql)
		end if
	end if
	rsLargestEmployeeID.close
	set rsLargestEmployeeID = nothing

	Session("strMessage") = dictLanguage("Employee_Added")		
	if strMessage <> "" then 
		session("strMessage") = session("strMessage") & "<BR>" & strMessage
	end if
	Response.redirect "default.asp?act=hnd_thnk" 
	
elseif (strAct = "edit" and strID <> "") then

	sql = sql_UpdateEmployee( _
			strID, empName, empStartDate, empLogin, empTitle, empEmail, empEmailAlt, _
			empCommissionRate, empSalesGoal, empProductionGoal, empPassword, _
			empActive, empHourly, empType, empReportsTo, empDepartment, _
			empBirthDate, empHomePhone, empMobilePhone, empWorkPhone, _
			empWorkPhoneExt, empVoiceMail, empHomeStreet1, empHomeStreet2, _
			empHomeCity, empHomeState, empHomeZip, empIMName, empNewFile)
	'Response.Write sql & "<BR>"
	Call DoSQL(sql)

	Session("strMessage") = dictLanguage("Employee_Updated")		
	if strMessage <> "" then 
		session("strMessage") = session("strMessage") & "<BR>" & strMessage
	end if
	Response.redirect "default.asp?act=hnd_thnk" 

End if 
%>

<!-- #include file="../../includes/connection_close.asp" -->