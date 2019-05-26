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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This include file contains all the sql code used throughout the site as functions that return the
'sql statement. Two reasons for this. 1 - to consolidate all the code into one place so it easy to
'see what is already there and to reduce duplication. 2 - You can write separate code for access 
'and sql (in case you want to use some stored procedures in sql server or queries in access) easily 
'using the DB_BOOLACCESS variable and you can add new code for other databases that you see fit to 
'use.  There are a few exceptions to this rule, namely a few reports that build the sql statements
'depending on the input and most of the forum (discussion) code is stored in the MTS_Fourms.asp page.
'The functions are listed by sections.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Admin
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Calendar
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetAllEvents()
	dim sql
	sql = "SELECT * FROM tbl_calendar ORDER By calendar_date desc, calendar_starttime, calendar_heading"
	sql_GetAllEvents = sql
End Function

Function sql_GetEvents()
	dim sql
	sql = "SELECT id, calendar_date, calendar_heading, calendar_abstract, calendar_starttime, calendar_endtime " & _
		"FROM tbl_calendar " & _
		"WHERE calendar_live = " & PrepareBit(1) & " and calendar_deleted = 0 " & _
		"ORDER By calendar_date desc, calendar_starttime, calendar_heading"
	sql_GetEvents = sql
End Function

Function sql_GetEventsByDate(nDate)
	dim sql
	sql = "SELECT id, calendar_date, calendar_heading, calendar_abstract, calendar_starttime, calendar_endtime " & _
		"FROM tbl_calendar " & _
		"WHERE calendar_live = " & PrepareBit(1) & " and calendar_deleted = 0 and calendar_date = " & DB_DATEDELIMITER & MediumDate(nDate) & DB_DATEDELIMITER & " " & _
		"ORDER By calendar_starttime"
	sql_GetEventsByDate = sql
End Function

Function sql_GetClientEventsByDate(navDate)
	dim sql
	sql = "SELECT client_id, client_name, client_since FROM tbl_clients " & _
	"WHERE client_since = " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
	"ORDER by client_name"
	sql_GetClientEventsByDate = sql
End Function

Function sql_GetProjectEventsByDate(navDate)
	dim sql
	sql = "SELECT project_id, description, start_date, launch_date, client_name " & _
	"FROM tbl_projects, tbl_clients " & _
	"WHERE tbl_projects.client_id = tbl_clients.client_id and " & _
		"(start_date = " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
		"OR launch_date = " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & ")"
	sql_GetProjectEventsByDate = sql
End Function

Function sql_GetTaskEventsByDate(navDate, id)
	dim sql
	sql = "SELECT tbl_clients.client_name, tbl_projects.description as project_name, " & _
		"tbl_tasks.task_id, tbl_tasks.description as task_desc, tbl_tasks.datedue, " & _
		"tbl_employees.employeename as empname, tbl_tasks.assignedto, tbl_tasks.orderedby " & _
	"FROM tbl_tasks, tbl_projects, tbl_clients, tbl_employees " & _ 
	"WHERE tbl_tasks.project_id = tbl_projects.project_id and " & _
		"tbl_tasks.client_id = tbl_clients.client_id and " & _
		"tbl_projects.client_id = tbl_clients.client_id and " & _
		"tbl_tasks.assignedto = tbl_employees.employee_id and " & _
		"datedue = " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "and showit = " & PrepareBit(1) & " " 
	else
		sql = sql & "and [show] = " & PrepareBit(1) & " " 
	end if
	sql = sql & "and done = 0 " & _
		"and (assignedto = " & id & " or orderedby = " & id & ")"
	sql_GetTaskEventsByDate = sql
End Function

Function sql_GetEventsByID(nID)
	dim sql
	sql = "SELECT id, calendar_date, calendar_heading, calendar_content, " & _
		"calendar_starttime, calendar_abstract, calendar_endtime, calendar_live " & _
		"FROM tbl_calendar " & _
		"WHERE id = " & nID & ""
	sql_GetEventsByID = sql
End Function

Function sql_GetEventsByDateSpan(navDate, navDateLast)
	dim sql
	sql = "SELECT * FROM tbl_calendar " & _
	"WHERE calendar_date >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
		"and calendar_date <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & " " & _
		"and calendar_live = " & PrepareBit(1) & " and calendar_deleted = 0"
	sql_GetEventsByDateSpan = sql
End Function

Function sql_GetEventsByDateSpanAdmin(navDate, navDateLast)
	dim sql
	sql = "SELECT * FROM tbl_calendar " & _
	"WHERE calendar_date >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
		"and calendar_date <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & " " & _
		"and calendar_deleted = 0 " & _
	"ORDER BY calendar_date desc, calendar_starttime, calendar_heading"
	sql_GetEventsByDateSpanAdmin = sql
End Function

Function sql_GetClientEventsByDateSpan(navDate, navDateLast)
	dim sql
	sql = "SELECT client_name, client_id, client_since FROM tbl_clients " & _
	"WHERE client_since >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
		"and client_since <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & ""
	sql_GetClientEventsByDateSpan = sql
End Function	

Function sql_GetProjectEventsByDateSpan(navDate, navDateLast)	
	dim sql
	sql = "SELECT project_id, description, start_date, launch_date, client_name " & _
	"FROM tbl_projects, tbl_clients " & _
	"WHERE tbl_projects.client_id = tbl_clients.client_id and " & _
		"((start_date >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
		"and start_date <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & ") " & _
	"OR (launch_date >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
		"and launch_date <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & "))"
	sql_GetProjectEventsByDateSpan = sql
End Function

Function sql_GetTaskEventsByDateSpan(navDate, navDateLast, id)
	dim sql
	sql = "SELECT tbl_clients.client_name, tbl_projects.description as project_name, " & _
		"tbl_tasks.task_id, tbl_tasks.description as task_desc, tbl_tasks.datedue, " & _
		"tbl_employees.employeename as empname, tbl_tasks.assignedto, tbl_tasks.orderedby " & _
	"FROM tbl_tasks, tbl_projects, tbl_clients, tbl_employees " & _ 
	"WHERE tbl_tasks.project_id = tbl_projects.project_id and " & _
		"tbl_tasks.client_id = tbl_clients.client_id and " & _
		"tbl_projects.client_id = tbl_clients.client_id and " & _
		"tbl_tasks.assignedto = tbl_employees.employee_id and " & _
		"(datedue >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " " & _
		"and datedue <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & ") "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql &  "and showit = " & PrepareBit(1) & " and done = 0 "	
	else
		sql = sql &  "and [show] = " & PrepareBit(1) & " and done = 0 "
	end if		
	sql = sql & "and (assignedto = " & id & " or orderedby = " & id & ")"
	sql_GetTaskEventsByDateSpan = sql
End Function

Function sql_InsertCalendarEvent(eventDateEntered, eventDate, eventStartTime, eventEndTime, _
			eventHeading, eventAbstract, eventContent, eventLive, eventDeleted, eventEnteredBy)
	dim sql
	sql = "INSERT INTO tbl_calendar (calendar_dateentered, calendar_date, calendar_starttime, " & _
		"calendar_endtime, calendar_heading, calendar_abstract, calendar_content, " & _
		"calendar_live, calendar_deleted, calendar_enteredby) VALUES(" & _
		"" & DB_DATEDELIMITER & MediumDate(eventDateEntered) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(eventDate) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & eventStartTime & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & eventEndTime & DB_DATEDELIMITER & ", " & _
		"'" & eventHeading & "', '" & eventAbstract & "', " & _
		"'" & eventContent & "', " & eventLive & ", " & eventDeleted & ", " & _
		eventEnteredBy & ")"
	sql_InsertCalendarEvent = sql
End Function

Function sql_UpdateCalendarEvent(eventDate, eventStartTime, eventEndTime, _
			eventHeading, eventAbstract, eventContent, eventLive, id)
	dim sql 
	sql = "UPDATE tbl_calendar set calendar_date = " & DB_DATEDELIMITER & MediumDate(eventDate) & DB_DATEDELIMITER & ", " & _
		"calendar_starttime = " & DB_DATEDELIMITER & eventStartTime & DB_DATEDELIMITER & ", " & _
		"calendar_endtime = " & DB_DATEDELIMITER & eventEndTime & DB_DATEDELIMITER & ", " & _
		"calendar_heading = '" & eventHeading & "', " & _
		"calendar_abstract = '" & eventAbstract & "', " & _
		"calendar_content = '" & eventContent & "', " & _
		"calendar_live = " & PrepareBit(eventLive) & " " & _
	"WHERE id = " & id & ""
	sql_UpdateCalendarEvent = sql
End Function

Function sql_DeleteCalendarEvent(id)
	Dim sql
	sql = "UPDATE tbl_calendar set calendar_deleted = " & PrepareBit(1) & " " & _
		"WHERE id = " & id & ""
	sql_DeleteCalendarEvent = sql
End Function  
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Clients
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetClientsByID(nID)
	dim sql
	sql = "SELECT * FROM tbl_clients WHERE client_id = " & nID & ""
	sql_GetClientsByID = sql
End Function

Function sql_InsertPhoneLog(nEmpID, nDate, nTime, cID, nTalkedTo, nMessage)
	dim sql
	sql = "INSERT INTO tbl_phonelog (sbemp,calldate,calltime,client_id,talkedto,message) " & _
				" values ('" & nEmpID & "', " & _
				"" & DB_DATEDELIMITER & MediumDate(nDate) & DB_DATEDELIMITER & ", " & _
				"" & DB_DATEDELIMITER & nTime & DB_DATEDELIMITER & ", " & _
				"" & cID & ", " & _
				"'" & nTalkedTo & "', " & _
				"'" & nMessage & "')"
	sql_InsertPhoneLog = sql
End Function

Function sql_InsertClient(Client_Name,Rep_ID,Active,Client_Since,Standard_Rate,Address1,Address2, _
	City_State_Zip,LiveSite_URL,Contact_Name,Contact_Phone,Contact_Email,Contact2_Name, _
	Contact2_Phone,Contact2_Email,Notes)
	dim sql
	sql = "INSERT into tbl_clients (client_name,rep_id,active,client_since,standard_rate,address1," & _
		"address2,city_state_zip,livesite_url,contact_name,contact_phone,contact_email,contact2_name," & _
		"contact2_phone,contact2_email,notes) VALUES (" & _
		"'" & Client_Name & "'," & _
		Rep_id & "," & _
		Active & "," & _
		"" & DB_DATEDELIMITER & MediumDate(Client_Since) & DB_DATEDELIMITER & "," & _
		Standard_Rate & "," & _
		"'" & Address1 & "'," & _
		"'" & Address2 & "'," & _
		"'" & City_State_Zip & "'," & _
		"'" & LiveSite_URL & "'," & _
		"'" & Contact_Name & "'," & _
		"'" & Contact_Phone & "'," & _
		"'" & Contact_Email & "'," & _
		"'" & Contact2_Name & "'," & _
		"'" & Contact2_Phone & "'," & _
		"'" & Contact2_Email & "'," & _
		"'" & Notes & "')"
	sql_InsertClient = sql
end Function

Function sql_UpdateClient(name, rep, client_since, standard_rate, address1, address2, _
	city_state_zip, livesite_url, contact_name, contact_phone, contact_email, contact2_name, _
	contact2_phone, contact2_email, active, client_id)
	dim sql
	sql = "Update tbl_clients SET " & _
		"client_name = '" & Name & "'," & _
		"rep_id=" & Rep & ","
		if Client_Since <> "" then
			sql = sql & "client_since=" & DB_DATEDELIMITER & MediumDate(Client_Since) & DB_DATEDELIMITER & ","
		else
			sql = sql & "client_since=NULL,"
		end if
	sql = sql & "standard_rate=" & Standard_Rate & "," & _
		"address1='" & Address1 & "'," & _
		"address2='" & Address2 & "'," & _
		"city_state_zip='" & City_State_Zip & "'," & _
		"livesite_url='" & LiveSite_URL & "'," & _
		"contact_name='" & Contact_Name & "'," & _
		"contact_phone='" & Contact_Phone & "'," & _
		"contact_email='" & Contact_Email & "'," & _
		"contact2_name='" & Contact2_Name & "'," & _
		"contact2_phone='" & Contact2_Phone & "'," & _
		"contact2_email='" & Contact2_Email & "'," & _
		"active=" & Active & " " & _
		"where client_id = " & client_id & ""	
	sql_UpdateClient = sql
End Function

Function sql_GetActiveClientsWithRep()
	dim sql
	sql = "SELECT tbl_clients.client_id, tbl_clients.client_name, tbl_employees.employeename " & _
		"FROM tbl_clients " & _
		"INNER JOIN tbl_employees ON tbl_clients.rep_id = tbl_employees.employee_id " & _
		"WHERE tbl_clients.active=" & PrepareBit(1) & " " & _
		"ORDER BY tbl_clients.client_name"
	sql_GetActiveClientsWithRep = sql
End Function

Function sql_GetAllClientsWithRep()
	dim sql
	sql = "SELECT tbl_clients.client_id, tbl_clients.client_name, tbl_employees.employeename " & _
		"FROM tbl_clients " & _
		"INNER JOIN tbl_employees ON tbl_clients.rep_id = tbl_Employees.Employee_ID " & _
		"ORDER BY tbl_clients.client_name"
	sql_GetAllClientsWithRep = sql
End Function

Function sql_GetActiveClientsByRepID(id)
	dim sql
	sql = "SELECT tbl_clients.client_id, tbl_clients.client_name, tbl_employees.employeename " & _
		"FROM tbl_clients " & _
		"INNER JOIN tbl_employees ON tbl_clients.rep_id = tbl_employees.employee_id " & _
		"WHERE tbl_clients.active=" & PrepareBit(1) & " and " & _
			"tbl_clients.rep_id = tbl_employees.employee_id and " & _
			"tbl_employees.employeelogin = '" & id & "' " & _
		"ORDER BY tbl_clients.client_name"
	sql_GetActiveClientsByRepID = sql
End Function 

Function sql_GetAllClientsByRepID(id)
	dim sql
	sql = "SELECT tbl_clients.client_id, tbl_clients.client_name, tbl_employees.employeename " & _
		"FROM tbl_clients " & _
		"INNER JOIN tbl_employees ON tbl_clients.rep_id = tbl_employees.employee_id " & _
		"WHERE tbl_clients.rep_id = tbl_employees.employee_id and " & _
			"tbl_employees.employeelogin = '" & id & "' " & _
		"ORDER BY tbl_clients.client_name"
	sql_GetAllClientsByRepID = sql
End Function 

Function sql_GetActiveClientsByID(id)
	dim sql
	sql = "SELECT * FROM tbl_clients " & _
		"WHERE client_id = " & id & " and active = " & PrepareBit(1) & ""
	sql_GetActiveClientsByID = sql
End Function 

Function sql_GetAllClients()
	dim sql
	sql = "SELECT * FROM tbl_clients ORDER BY client_name"
	sql_GetAllClients = sql
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Document Repository
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetFoldersCount()
	dim sql
	if DB_BOOLACCESS or DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = "SELECT tbl_folders.folder_id, " & _
			"(Count(tbl_folders_1.sub_to) + Count(tbl_documents.folder)) as rec_count " & _
			"FROM tbl_documents RIGHT JOIN " & _
			"(tbl_folders AS tbl_folders_1 RIGHT JOIN " & _	
			"tbl_folders ON tbl_folders_1.sub_to = tbl_folders.folder_id) " & _
			"ON tbl_documents.folder = tbl_folders.folder_id " & _
			"GROUP BY tbl_folders.folder_id"
	else
		sql = "SELECT tbl_folders.folder_id, " & _
			"(Count(folders.sub_to) + Count(tbl_documents.folder)) as rec_count " & _
			"FROM tbl_folders " & _
			"LEFT JOIN tbl_folders as folders on tbl_folders.folder_id = folders.sub_to " & _
			"LEFT JOIN tbl_documents on tbl_folders.folder_id = tbl_documents.folder " & _
			"WHERE " & _ 
			"tbl_folders.folder_id = tbl_documents.folder " & _
			"or tbl_folders.folder_id = folders.sub_to " & _
			"GROUP BY tbl_folders.folder_id"
	end if
	sql_GetFoldersCount = sql
End Function

Function sql_GetFolderByID(id)
	dim sql 
	sql = "SELECT * FROM tbl_folders WHERE folder_id = " & id & ""
	sql_GetFolderByID = sql
End Function

Function sql_GetSubFoldersByFolderID(id)
	dim sql
	sql = "SELECT * FROM tbl_folders WHERE sub_to = " & id & " ORDER BY folder"
	sql_GetSubFoldersByFolderID = sql
End Function 

Function sql_GetDocumentByDocumentID(id)
	dim sql
	sql = "SELECT * FROM tbl_documents " & _
		"WHERE id = " & id & ""
	sql_GetDocumentByDocumentID = sql
End Function

Function sql_UpdateDocument(id, submitby,description,folder)
	dim sql 
	sql = "UPDATE tbl_documents set submitby = " & submitby & ", " & _
		"description = '" & description & "', " & _
		"folder = " & folder & " " & _
	"WHERE id = " & id & ""
	sql_UpdateDocument = sql
End Function

Function sql_UpdateFolder(id, folder,sub_to)
	dim sql 
	sql = "UPDATE tbl_folders set folder = '" & folder & "', " & _
		"sub_to = " & sub_to & " " & _
	"WHERE folder_id = " & id & ""
	'Response.Write sql
	sql_UpdateFolder = sql
End Function

Function sql_DeleteFolder(id)
	dim sql
	sql = "DELETE FROM tbl_folders WHERE (folder_id = " & id & ") OR (sub_to = " & id & ") "
	sql_DeleteFolder = sql
End Function

Function sql_GetDocumentsByFolderID(id)
	dim sql
	sql = "SELECT * FROM tbl_documents " & _
		"WHERE folder = " & id & " " & _
		"ORDER BY submitDate DESC, filename"
	sql_GetDocumentsByFolderID = sql
End Function

Function sql_GetAllFolders()
	dim sql
	sql = "SELECT * FROM tbl_folders ORDER BY folder"
	sql_GetAllFolders = sql
End Function

Function sql_GetMainFolders()
	dim sql
	sql = "SELECT * FROM tbl_folders where sub_to=0 ORDER BY folder"
	sql_GetMainFolders = sql
End Function

Function sql_DeleteDocument(id)
	dim sql
	sql = "DELETE FROM tbl_documents WHERE id = " & id & ""
	sql_DeleteDocument = sql
End Function

Function sql_InsertDocument(fileNameShow, Folder, URL, Description, employee_id, eDate) 
	dim sql
	sql = "INSERT INTO tbl_documents (filename, folder, url, description, submitby, submitdate) " & _ 
		"values (" & _
		"'" & fileNameShow & "', " & _
		Folder & ", " & _
		"'" & URL & "', " & _
		"'" & Description & "', " & _
		employee_id & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(eDate) & DB_DATEDELIMITER & ")"
	sql_InsertDocument = sql
End Function

Function sql_InsertFolder(folderName, subTo)
	dim sql
	sql = "INSERT INTO tbl_folders (folder, sub_to) VALUES ('" & folderName & "', " & subTo & ")"
	sql_InsertFolder = sql
End Function

Function sql_GetMyLatestFolder(folderName, subTo)
	dim sql
	sql = "SELECT folder_id FROM tbl_folders WHERE folder = '" & folderName & "' and sub_to = " & subTo & ""
	sql_GetMyLatestFolder = sql
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Employees
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetEmployeesByID(id)
	dim sql
	sql = "Select * from tbl_employees where employee_id = " & id & ""
	sql_GetEmployeesByID = sql
End Function

Function sql_GetEmployeesPermissionsByID(id)
	dim sql
	sql = "Select * from tbl_permissions where employee_id = " & id & ""
	sql_GetEmployeesPermissionsByID = sql
End Function

Function sql_GetEmployeesByLogin(id)
	dim sql
	sql = "SELECT * FROM tbl_employees WHERE (employeelogin = '" & Ucase(id) & "')"
	sql_GetEmployeesByLogin = sql
End Function

Function sql_UpdateEmployeesLastLogin(strDate, login)
	dim sql
	sql = "UPDATE tbl_employees SET lastlogin = " & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & " " & _
	"WHERE employeelogin = '" & Ucase(login) & "'"
	sql_UpdateEmployeesLastLogin = sql
End Function

Function sql_GetActiveEmployees()
	dim sql
	sql = "SELECT * " & _
		"FROM tbl_employees " & _
		"WHERE active = " & PrepareBit(1) & " " & _
		"ORDER BY employeename"
	sql_GetActiveEmployees = sql
End Function

Function sql_GetAllEmployees()
	dim sql
	sql = "SELECT * " & _
		"FROM tbl_employees " & _
		"ORDER BY employeename"
	sql_GetAllEmployees = sql
End Function

Function sql_GetHourlyEmployees()
	dim sql
	sql = "SELECT * FROM tbl_employees WHERE hourly = " & PrepareBit(1) & ""
	sql_GetHourlyEmployees = sql
End Function

Function sql_GetHourlyEmployeesByID(strEmpID)
	dim sql
	sql = "SELECT * from tbl_employees " & _
	"WHERE employee_id = " & strEmpID & " and hourly = " & PrepareBit(1) & ""
	sql_GetHourlyEmployeesByID = sql
End Function

Function sql_GetEmployeeTypes()
	dim sql
	sql = "SELECT * FROM tbl_employeetypes order by employeetype_id"
	sql_GetEmployeeTypes = sql
End Function

Function sql_GetEmployeeTypesByID(id)
	dim sql
	sql = "SELECT * FROM tbl_employeetypes where employeetype_id = " & id & ""
	sql_GetEmployeeTypesByID = sql
End Function

Function sql_InsertEmployeeType(strType, strRate, strCost)
	dim sql
	sql = "INSERT INTO tbl_employeetypes (employeetype, rate, cost) VALUES (" & _
		"'" & strType & "', " & strRate & ", " & strCost & ")"
	sql_InsertEmployeeType = sql
End Function

Function sql_UpdateEmployeeType(strID, strType, strRate, strCost)
	dim sql
	sql = "UPDATE tbl_employeetypes SET " & _
		"employeetype = '" & strType & "', " & _
		"rate = " & strRate & ", " & _
		"cost = " & strCost & " " & _
		"WHERE employeetype_id = " & strID & ""
	sql_UpdateEmployeeType = sql
End Function

Function sql_DeleteEmployeeType(strID)
	dim sql
	sql = "DELETE FROM tbl_employeetypes where employeetype_id = " & strID & ""
	sql_DeleteEmployeeType = sql
End Function

Function sql_GetActiveDepartments()
	dim sql
	sql = "SELECT DISTINCT department FROM tbl_employees " & _
		"WHERE active=" & PrepareBit(1) & " ORDER BY department"
	sql_GetActiveDepartments = sql
End Function

Function sql_GetEmployeesWithTypeByID(id)
	dim sql
	sql = "SELECT * FROM tbl_employees INNER JOIN tbl_employeetypes ON " & _
		"tbl_employees.employeetype_id = tbl_employeetypes.employeetype_id " & _
	"WHERE employee_id = " & id & ""
	sql_GetEmployeesWithTypeByID = sql
End Function

Function sql_GetActiveEmployeesPRGView()
	dim sql
	sql = "SELECT tbl_employees.employeename, tbl_employees.productiongoal, " & _
		"tbl_employees.employee_id, tbl_employees.employeelogin, tbl_employees.department " & _
	"FROM tbl_employees " & _
	"WHERE productiongoal > 0 and active=" & PrepareBit(1) & " " & _
	"ORDER BY department, employeename asc"
	sql_GetActiveEmployeesPRGView = sql
End Function

Function sql_InsertEmployee(empName, empStartDate, empLogin, empTitle, empEmail, empEmailAlt, _
			empCommissionRate, empSalesGoal, empProductionGoal, empPassword, _
			empActive, empHourly, empType, empReportsTo, empDepartment, _
			empBirthDate, empHomePhone, empMobilePhone, empWorkPhone, _
			empWorkPhoneExt, empVoiceMail, empHomeStreet1, empHomeStreet2, _
			empHomeCity, empHomeState, empHomeZip, empIMName, empNewFile)
	dim sql
	sql = "INSERT INTO tbl_employees (employeename, startdate, employeelogin, employeetitle, " & _
			"emailaddress, emailaddressalt, commissionrate, salesgoal, productiongoal, [password], " & _
			"active, hourly, employeetype_id, reportsto, department, bdate, homephone, " & _
			"mobilephone, workphone, workphoneext, voicemail, homestreet1, homestreet2, " & _
			"homecity, homestate, homezip, imname, [image]) VALUES ("
	sql = sql & "'" & empName & "', "
	if empStartDate = "" then
		sql = sql & "NULL, "
	else
		sql = sql & "" & DB_DATEDELIMITER & MediumDate(empStartDate) & DB_DATEDELIMITER & ", "
	end if
	sql = sql & "'" & empLogin & "', '" & empTitle & "', '" & empEmail & "', " & _
		"'" & empEmailAlt & "', " & empCommissionRate & ", " & empSalesGoal & ", " & _
		empProductionGoal & ", '" & empPassword & "', " & empActive & ", " & empHourly & ", " & _
		empType & ", " & empReportsTo & ", '" & empDepartment & "', "
	if empBirthDate = "" then
		sql = sql & "NULL, "
	else
		sql = sql & "" & DB_DATEDELIMITER & MediumDate(empBirthDate) & DB_DATEDELIMITER & ", " 
	end if
	sql = sql & "'" & empHomePhone & "', '" & empMobilePhone & "', '" & empWorkPhone & "', " & _
		"'" & empWorkPhoneExt & "', '" & empVoiceMail & "', '" & empHomeStreet1 & "', " & _
		"'" & empHomeStreet2 & "', '" & empHomeCity & "', '" & empHomeState & "', " & _
		"'" & empHomeZip & "', '" & empIMNAME & "', '" & empNewFile & "')"		 
	sql_InsertEmployee = sql
End Function

Function sql_UpdateEmployee( _
			empID, empName, empStartDate, empLogin, empTitle, empEmail, empEmailAlt, _
			empCommissionRate, empSalesGoal, empProductionGoal, empPassword, _
			empActive, empHourly, empType, empReportsTo, empDepartment, _
			empBirthDate, empHomePhone, empMobilePhone, empWorkPhone, _
			empWorkPhoneExt, empVoiceMail, empHomeStreet1, empHomeStreet2, _
			empHomeCity, empHomeState, empHomeZip, empIMName, empNewFile)
	dim sql
	sql = "UPDATE tbl_employees SET " & _
		"employeename = '" & empName & "', "
	if empStartDate = "" then
		sql = sql & "startdate = NULL, "
	else
		sql = sql & "startdate = " & DB_DATEDELIMITER & MediumDate(empStartDate) & DB_DATEDELIMITER & ", "
	end if
	sql = sql & "employeelogin = '" & empLogin & "', " & _
		"employeetitle = '" & empTitle & "', " & _
		"emailaddress = '" & empEmail & "', " & _
		"emailaddressalt = '" & empEmailAlt & "', " & _
		"commissionrate = " & empCommissionRate & ", " & _
		"salesgoal = " & empSalesGoal & ", " & _
		"productiongoal = " & empProductionGoal & ", " & _
		"[password] = '" & empPassword & "', " & _
		"active = " & empActive & ", " & _
		"hourly = " & empHourly & ", " & _
		"employeetype_id = " & empType & ", " & _
		"reportsto = " & empReportsTo & ", " & _
		"department = '" & empDepartment & "', "
	if empBirthDate = "" then
		sql = sql & "bdate = NULL, "
	else
		sql = sql & "bdate = " & DB_DATEDELIMITER & MediumDate(empBirthDate) & DB_DATEDELIMITER & ", "
	end if
	sql = sql & "homephone = '" & empHomePhone & "', " & _
		"mobilephone = '" & empMobilePhone & "', " & _
		"workphone = '" & empWorkPhone & "', " & _
		"workphoneext = '" & empWorkPhoneExt & "', " & _
		"voicemail = '" & empVoiceMail & "', " & _
		"homestreet1 = '" & empHomeStreet1 & "', " & _
		"homestreet2 = '" & empHomeStreet2 & "', " & _
		"homecity = '" & empHomeCity & "', " & _
		"homestate = '" & empHomeState & "', " & _
		"homezip = '" & empHomeZip & "', " & _
		"imname = '" & empIMName & "', " & _
		"[image] = '" & empNewFile & "'" & _
	"WHERE employee_id = " & empID  & ""
	sql_UpdateEmployee = sql
End Function

Function sql_UpdateEmployeeProfile( _
			strID, empEmail, empEmailAlt, empPassword, _
			empHomePhone, empMobilePhone, empWorkPhone, _
			empWorkPhoneExt, empVoiceMail, empHomeStreet1, empHomeStreet2, _
			empHomeCity, empHomeState, empHomeZip, empIMName, empNewFile)
	dim sql
	sql = "UPDATE tbl_employees SET " & _
		"emailaddress = '" & empEmail & "', " & _
		"emailaddressalt = '" & empEmailAlt & "', " & _
		"[password] = '" & empPassword & "', " & _
		"homephone = '" & empHomePhone & "', " & _
		"mobilephone = '" & empMobilePhone & "', " & _
		"workphone = '" & empWorkPhone & "', " & _
		"workphoneext = '" & empWorkPhoneExt & "', " & _
		"voicemail = '" & empVoiceMail & "', " & _
		"homestreet1 = '" & empHomeStreet1 & "', " & _
		"homestreet2 = '" & empHomeStreet2 & "', " & _
		"homecity = '" & empHomeCity & "', " & _
		"homestate = '" & empHomeState & "', " & _
		"homezip = '" & empHomeZip & "', " & _
		"imname = '" & empIMName & "', " & _
		"[image] = '" & empNewFile & "' " & _
	"WHERE employee_id = " & strID  & ""
	sql_UpdateEmployeeProfile = sql
End Function			
			
Function sql_UpdatePermissions(strEmpID, permAll, permAdmin, permAdminCalendar, _
					permAdminDatabaseSetup, permAdminEmployees, permAdminEmployeesPerms, _
					permAdminFileRepository, permAdminForum, permAdminNews, permAdminResources, _
					permAdminSurveys, permAdminThoughts, permClientsAdd, permClientsEdit, _
					permClientsDelete, permProjectsAdd, permProjectsEdit, permProjectsDelete, _
					permTasksAdd, permTasksEdit, permTasksDelete, permTimecardsAdd, _
					permTimecardsEdit, permTimecardsDelete, permTimesheetsEdit, permForumAdd, _
					permRepositoryAdd, permPTOAdmin)
	dim sql
	sql = sql_GetEmployeesPermissionsByID(strEmpID)
	Call RunSQL(sql, rsPerms)
	if not rsPerms.eof then
		sql = "UPDATE tbl_permissions set " & _
			"permall = " & permAll & ", " & _
			"permadmin = " & permAdmin & ", " & _
			"permadmincalendar = " & permAdminCalendar & ", " & _
			"permadmindatabasesetup = " & permAdminDatabaseSetup & ", " & _
			"permadminemployees = " & permAdminemployees & ", " & _
			"permadminemployeesperms = " & permAdminemployeesperms & ", " & _
			"permadminfilerepository = " & permAdminFileRepository & ", " & _
			"permadminforum = " & permAdminForum & ", " & _
			"permadminnews = " & permAdminNews & ", " & _
			"permadminresources = " & permAdminResources & ", " & _
			"permadminsurveys = " & permAdminSurveys & ", " & _
			"permadminthoughts = " & permAdminThoughts & ", " & _
			"permclientsadd = " & permClientsAdd & ", " & _
			"permclientsedit = " & permClientsEdit & ", " & _
			"permclientsdelete = " & permClientsDelete & ", " & _
			"permprojectsadd = " & permProjectsAdd & ", " & _
			"permprojectsedit = " & permProjectsEdit & ", " & _
			"permprojectsdelete = " & permProjectsDelete & ", " & _
			"permtasksadd = " & permTasksAdd & ", " & _
			"permtasksedit = " & permTasksEdit & ", " & _
			"permtasksdelete = " & permTasksDelete & ", " & _
			"permtimecardsadd = " & permTimecardsAdd & ", " & _
			"permtimecardsedit = " & permTimecardsEdit & ", " & _
			"permtimecardsdelete = " & permTimecardsDelete & ", " & _
			"permtimesheetsedit = " & permTimesheetsEdit & ", " & _
			"permforumadd = " & permForumAdd & ", " & _
			"permrepositoryadd = " & permRepositoryAdd & ", " & _
			"permptoadmin = " & permPTOAdmin & " " & _
		"WHERE employee_id = " & strEmpID & ""
	else
		sql = "INSERT INTO tbl_permissions (employee_id, permall, permadmin, permadmincalendar, " &  _
					"permadmindatabasesetup, permadminemployees, permadminemployeesperms, " & _
					"permadminfilerepository, permadminforum, permadminnews, permadminresources, " & _
					"permadminsurveys, permadminthoughts, permclientsadd, permclientsedit, " & _
					"permclientsdelete, permprojectsadd, permprojectsedit, permprojectsdelete, " & _
					"permtasksadd, permtasksedit, permtasksdelete, permtimecardsadd, " & _
					"permtimecardsedit, permtimecardsdelete, permtimesheetsedit, permforumadd, " & _
					"permrepositoryadd, permptoadmin) VALUES (" & _ 
			strEmpID & ", " & _
			permAll & ", " & _
			permAdmin & ", " & _
			permAdminCalendar & ", " & _
			permAdminDatabaseSetup & ", " & _
			permAdminEmployees & ", " & _
			permAdminEmployeesPerms & ", " & _			
			permAdminFileRepository & ", " & _
			permAdminForum & ", " & _
			permAdminNews & ", " & _
			permAdminResources & ", " & _
			permAdminSurveys & ", " & _
			permAdminThoughts & ", " & _
			permClientsAdd & ", " & _
			permClientsEdit & ", " & _
			permClientsDelete & ", " & _
			permProjectsAdd & ", " & _
			permProjectsEdit & ", " & _
			permProjectsDelete & ", " & _
			permTasksAdd & ", " & _
			permTasksEdit & ", " & _
			permTasksDelete & ", " & _
			permTimecardsAdd & ", " & _
			permTimecardsEdit & ", " & _
			permTimecardsDelete & ", " & _
			permTimesheetsEdit & ", " & _
			permForumAdd & ", " & _
			permRepositoryAdd & ", " & _
			permPTOAdmin & ")"
		end if
	rsPerms.close
	set rsPerms = nothing
	sql_UpdatePermissions = sql
End Function

Function sql_GetPasswordReminder(emailaddress)
	dim sql
	sql = "SELECT employeelogin, [password], active from tbl_employees " & _
		"WHERE emailaddress = '" & emailaddress & "'"
	sql_GetPasswordReminder = sql
End Function 

Function sql_SelectLatestEmployeeID()
	dim sql
	sql = "SELECT MAX(employee_id) as maxid FROM tbl_employees"
	sql_SelectLatestEmployeeID = sql
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Forum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Forum code is kept in the files forum_common.asp and MTS_Forums.asp
Function sql_GetTop3Forums()
	dim sql
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = "SELECT * FROM forums ORDER BY forum_start_date DESC, forum_id DESC LIMIT 3"
	else
		sql = "SELECT top 3 * FROM forums ORDER BY forum_start_date DESC, forum_id DESC"
	end if	
	sql_GetTop3Forums = sql
End Function

Function sql_GetForumMessageCountByForum(forumID)
	dim sql
	sql = "SELECT COUNT(message_id) AS nummessages FROM messages " & _
	"WHERE forum_id = " & forumID & ""
	sql_GetForumMessageCountByForum = sql
End Function

Function sql_InsertForum(forumName, forumDesc, startDate, strGroup, strSection)
	dim sql
	sql = "INSERT INTO forums (forum_name, forum_description, forum_start_date, forum_grouping, forum_section) " & _
	"VALUES (" & _
	"'" & forumName & "', " & _
	"'" & forumDesc & "', " & _
	"" & DB_DATEDELIMITER & MediumDate(startDate) & DB_DATEDELIMITER & ", " & _
	"'" & strGroup & "', " & strSection & ")"
	sql_InsertForum = sql
End Function

Function sql_GetMyLatestForum(forumName, forumDesc, startDate)
	dim sql
	sql = "SELECT forum_id from forums " & _
	"WHERE forum_name = '" & forumName & "' and forum_description = '" & forumDesc & "' " & _
		"and forum_start_date = " & DB_DATEDELIMITER & MediumDate(startDate) & DB_DATEDELIMITER & ""		
	 sql_GetMyLatestForum = sql
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Main Page (main.asp - the home page) 
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetLeastUsedThought()
	dim sql
	sql = "SELECT * FROM tbl_thoughts ORDER BY shows ASC"
	sql_GetLeastUsedThought = sql
End Function

Function sql_IncrementThoughtByID(strThoughtID) 
	dim sql
	sql = "UPDATE tbl_thoughts SET shows = shows + 1 WHERE id = " & strThoughtID & ""
	sql_IncrementThoughtByID = sql
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Production Report Grid
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetTeamMonthlyProductionGoal()
	dim sql
	sql = "SELECT SUM(tbl_employees.productiongoal) AS sumofproductiongoal FROM tbl_employees"
	sql_GetTeamMonthlyProductionGoal = sql
End Function

Function sql_GetTeamMonthlyBillableHours(strDate)
	dim sql
	sql = "SELECT Sum(timeamount) AS workthismonth " & _
	"FROM tbl_timecards "
	if DB_BOOLPOSTGRES then
		sql = sql & "WHERE date_part('month', dateworked) = date_part('month', timestamp '" & MediumDate(strDate) & "') " & _
		"AND date_part('year', dateworked) = date_part('year', timestamp '" & MediumDate(strDate) & "') "
	else
		sql = sql & "WHERE Month(dateworked) = " & month(strDate) & " " & _
		"AND Year(dateworked) = " & year(strDate) & " "	
	end if
	if gsHostNonBillable then
		sql = sql & "AND (Not (tbl_timecards.client_id)=1)"
	end if
	sql_GetTeamMonthlyBillableHours = sql
End Function

Function sql_GetTeamWeeklyBillableHours(strMon, strSun)
	dim sql
	sql = "SELECT Sum(timeamount) AS workthisweek " & _
	"FROM tbl_timecards " & _
	"WHERE dateworked >= " & DB_DATEDELIMITER & MediumDate(strMon) & DB_DATEDELIMITER & " and " & _
		"dateworked <= " & DB_DATEDELIMITER & MediumDate(strSun) & DB_DATEDELIMITER & " "
	if gsHostNonBillable then
		sql = sql &	"AND Not(tbl_timecards.client_id=1)"
	end if
	sql_GetTeamWeeklyBillableHours = sql
End Function

Function sql_GetTeamDailyBillableHours(strDate)
	dim sql
	sql= "SELECT SUM(tbl_timecards.timeamount) AS worktoday " & _
	"FROM tbl_timecards " & _
	"WHERE dateworked = " & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & " "
	if gsHostNonBillable then
		sql = sql & "AND Not(tbl_timecards.client_id=1)"	
	end if
	sql_GetTeamDailyBillableHours = sql
End Function

Function sql_GetTeamWeekendBillableHours(strMonday)
	dim sql
	sql = "SELECT SUM(timeamount) AS worktoday " & _
	"FROM tbl_timecards " & _
	"WHERE (dateworked = " & DB_DATEDELIMITER &  MediumDate(dateAdd("d",-1,strMonday)) & DB_DATEDELIMITER & " OR " & _
		"dateworked = " & DB_DATEDELIMITER & MediumDate(dateAdd("d",-2,strMonday)) & DB_DATEDELIMITER & " OR " & _
		"dateworked = " & DB_DATEDELIMITER & MediumDate(dateAdd("d",-3,strMonday)) & DB_DATEDELIMITER & ") " 
	if gsHostNonBillable then
		sql = sql & "AND (NOT (tbl_timecards.client_id=1))"
	end if
	sql_GetTeamWeekendBillableHours = sql
End Function

Function sql_GetEmployeeMonthlyProductionGoal(strEmpID)
	dim sql
	sql = "SELECT productiongoal FROM tbl_employees " & _
	"WHERE tbl_employees.employee_id = " & strEmpID & ""
	sql_GetEmployeeMonthlyProductionGoal = sql
End Function

Function sql_GetEmployeeMonthlyBillableHours(strEmpID, strDate)
	dim sql
	sql = "SELECT Sum(timeamount) AS workthismonth FROM tbl_timecards "
	if DB_BOOLPOSTGRES then
		sql = sql & "WHERE date_part('month', dateworked) = date_part('month', timestamp '" & MediumDate(strDate) & "') " & _
		"And date_part('year', dateworked) = date_part('year', timestamp '" & MediumDate(strDate) & "') "
	else
		sql = sql & "WHERE Month(dateworked) = " & month(strDate) & " " & _
		"And Year(dateworked) = " & year(strDate) & " "
	end if
	sql = sql & "and employee_id = " & strEmpID & " " 
	if gsHostNonBillable then
		sql = sql & "AND client_id<>1"	
	end if
	sql_GetEmployeeMonthlyBillableHours = sql
End Function

Function sql_GetEmployeeDailyBillableHours(strEmpID, strDate)
	dim sql
	sql = "SELECT Sum(timeamount) AS worktoday " & _
	"FROM tbl_timecards "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "WHERE dateworked = " & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & " " & _
		"AND employee_id = " & strEmpID & " "	
	else
		sql = sql & "WHERE dateworked = " & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & " " & _
		"AND employee_id = " & strEmpID & " "
	end if
	if gsHostNonBillable then
		sql = sql & "AND client_id<>1 "
	end if
	sql_GetEmployeeDailyBillableHours = sql
End Function

Function sql_GetEmployeeWeekendBillableHours(employeeID, datToday)
	dim sql
	'CHANGED WorkThisWeek to WorkToday to solve error i recieved on a monday
	sql = "SELECT Sum([timeamount]) AS worktoday " & _
	"FROM tbl_timecards " & _
	"WHERE tbl_timecards.employee_id = " & employeeID & " and " & _
		"(dateworked = " & DB_DATEDELIMITER &  MediumDate(dateAdd("d",-1,datToday)) & DB_DATEDELIMITER & " OR " & _
		"dateworked = " & DB_DATEDELIMITER & MediumDate(dateAdd("d",-2,datToday)) & DB_DATEDELIMITER & " OR " & _
		"dateworked = " & DB_DATEDELIMITER & MediumDate(dateAdd("d",-3,datToday)) & DB_DATEDELIMITER & ") " 
	if gsHostNonBillable then
		sql = sql & "AND (NOT (tbl_timecards.client_id=1))"
	end if
	sql_GetEmployeeWeekendBillableHours = sql
End Function

Function sql_GetEmployeeWeeklyBillableHours(employeeID, ThisMon, ThisSun)
	dim sql
	sql = "SELECT SUM(timeamount) AS workthisweek " & _
	"FROM tbl_timecards " & _
	"WHERE tbl_timecards.employee_id = " & employeeID & " AND " & _
		"dateworked >= " & DB_DATEDELIMITER & MediumDate(ThisMon) & DB_DATEDELIMITER & " " & _
		"and dateworked <= " & DB_DATEDELIMITER & MediumDate(ThisSun) & DB_DATEDELIMITER & " "
	if gsHostNonBillable then
		sql = sql & "AND (Not(tbl_timecards.client_id=1))"
	end if
	sql_GetEmployeeWeeklyBillableHours = sql
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PTO - ** added in Version 2.7.2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetPTOByDate(eDate)
	dim sql
	sql = "SELECT tbl_employees.employeename, tbl_pto.* " & _
		"FROM tbl_pto, tbl_employees " & _
		"WHERE ((Start_Date = " & DB_DATEDELIMITER & MediumDate(eDate) & DB_DATEDELIMITER & ") or " & _
			 "(End_Date = " & DB_DATEDELIMITER & MediumDate(eDate) & DB_DATEDELIMITER & ") or " & _
			 "(Start_Date <= " & DB_DATEDELIMITER & MediumDate(eDate) & DB_DATEDELIMITER & " and " & _
			 "End_Date >= " & DB_DATEDELIMITER & MediumDate(eDate) & DB_DATEDELIMITER & ")) and " & _
			 "tbl_pto.approved = " & PrepareBit(1) & " and tbl_pto.employee_id = tbl_employees.employee_id"
	sql_GetPTOByDate = sql
End Function

Function sql_GetPTOByDateSpan(navDate, navDateLast)
	dim sql
	sql = "SELECT tbl_employees.employeename, tbl_pto.* " & _
		"FROM tbl_pto, tbl_employees " & _
		"WHERE ((Start_Date >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " and " & _
			"Start_Date <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & ") or " & _
			"(End_Date >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " and " & _
			"End_Date <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & ") or " & _
			"(Start_Date >= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " and " & _
			"End_Date <= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & ") or " & _
			"(Start_Date <= " & DB_DATEDELIMITER & MediumDate(navDate) & DB_DATEDELIMITER & " and " & _
			"End_Date >= " & DB_DATEDELIMITER & MediumDate(navDateLast) & DB_DATEDELIMITER & ")) " & _
			"and tbl_pto.approved = " & PrepareBit(1) & " and tbl_pto.employee_id = tbl_employees.employee_id"
	sql_GetPTOByDateSpan = sql
End Function

Function sql_GetEmployeePTOByDateRange(date_query)
	dim sql
	sql = "SELECT employeename, tbl_pto.* " & _
		"FROM tbl_employees, tbl_pto " & _
		"WHERE tbl_employees.active = " & PrepareBit(1) & " and " & _
			"tbl_employees.employee_id = tbl_pto.employee_id" & date_query & " " & _
		"ORDER by employeename, start_date"
	sql_GetEmployeePTOByDateRange = sql
End Function

Function sql_GetEmployeesPTO()
	dim sql
	sql= "SELECT employeename, tbl_pto.* " & _
		"FROM tbl_employees, tbl_pto " & _
		"WHERE tbl_employees.employee_id = tbl_pto.employee_id and " & _
			"tbl_employees.active = " & PrepareBit(1) & " " & _
		"ORDER by employeename, start_date"
	sql_GetEmployeesPTO = sql
End Function

Function sql_GetEmployeesPTONotCompleted()
	dim sql
	sql = "SELECT employeename, tbl_pto.* " & _
		"FROM tbl_employees, tbl_pto " & _
		"WHERE tbl_employees.employee_id = tbl_pto.employee_id and " & _
			"tbl_employees.active = " & PrepareBit(1) & " and updated=0 " & _
		"ORDER by employeename, start_date"
	sql_GetEmployeesPTONotCompleted = sql
End Function

Function sql_GetEmployeePTOByID(eID)
	dim sql
	sql = "SELECT employeename, tbl_pto.* " & _
		"FROM tbl_employees, tbl_pto " & _
		"WHERE tbl_employees.employee_id = tbl_pto.employee_id and " & _
			"tbl_employees.employee_id = " & eID & " " & _
		"ORDER by start_date"
	sql_GetEmployeePTOByID = sql
End Function

Function sql_GetPTOByID(pid)
	dim sql 
	sql = "SELECT * from tbl_pto where pto_id = " & pid & ""
	sql_GetPTOByID = sql
End Function

Function sql_UpdatePTO(Employee_ID, Start_Date, Start_Time, End_Date, _
	End_Time, Total_Hours, Paid, Balance, Reason, Approved, Excused, Updated, pto_ID)
	dim sql
	sql = "UPDATE tbl_pto SET " & _
			"Employee_ID = " & Employee_ID & "," & _
			"Start_Date = " & DB_DATEDELIMITER & Start_Date & DB_DATEDELIMITER & "," & _
			"Start_Time = " & DB_DATEDELIMITER & Start_Time & DB_DATEDELIMITER & "," & _
			"End_Date = " & DB_DATEDELIMITER & End_Date & DB_DATEDELIMITER & "," & _
			"End_Time = " & DB_DATEDELIMITER & End_Time & DB_DATEDELIMITER & "," & _
			"Total_Hours = " & Total_Hours & "," & _
			"Paid = " & PrepareBit(Paid) & "," & _
			"Balance = " & Balance & "," & _
			"Reason = '" & Reason & "'," & _
			"Approved = " & PrepareBit(Approved) & ", " & _
			"Excused = " & PrepareBit(Excused) & ", " & _
			"Updated = " & PrepareBit(Updated) & " " & _
		"WHERE pto_ID = " & pto_ID & ""
	sql_UpdatePTO = sql
End Function

Function sql_InsertPTO(employee_ID, Start_Date, Start_Time, End_Date, _
		End_Time, Total_Hours, Paid, Balance, Updated, Reason)
	dim sql
	sql = "INSERT INTO tbl_pto (employee_id, start_date, start_time, end_date, " & _
		"end_time, total_hours, paid, balance, updated, reason) VALUES(" & _
		employee_ID & ", " & _
		"" & DB_DATEDELIMITER & Start_Date & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & Start_Time & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & End_Date & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & End_Time & DB_DATEDELIMITER & ", " & _
		Total_Hours & ", " & _
		PrepareBit(Paid) & ", " & _
		Balance & ", " & _
		PrepareBit(updated) & ", " & _
		"'" & Reason & "')"
	sql_InsertPTO = sql
End Function

Function sql_GetLatestPTO()
	dim sql
	sql = "SELECT MAX(pto_id) AS max_pto FROM tbl_pto"
	sql_GetLatestPTO = sql
End Function
	
	
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'News
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetAllNews()
	dim sql 
	sql = "SELECT * FROM tbl_news ORDER BY news_date DESC, news_heading"
	sql_GetAllNews = sql
End Function

Function sql_GetActiveNews()
	dim sql 
	sql = "SELECT id, news_date, news_heading, news_url, news_source, news_abstract " & _
		"FROM tbl_news " & _
		"WHERE news_live = " & PrepareBit(1) & " and news_deleted = 0 " & _
		"ORDER BY news_date DESC, news_heading"
	sql_GetActiveNews = sql
End Function

Function sql_GetNewsByID(id)
	dim sql
	sql = "SELECT * " & _
		"FROM tbl_news " & _
		"WHERE id = " & id & ""
	sql_GetNewsByID = sql
End Function

Function sql_GetNewsByDateSpanAdmin(strStartDate, strEndDate)
	dim sql
	sql = "SELECT * from tbl_news where news_date >= " & DB_DATEDELIMITER & MediumDate(strStartDate) & DB_DATEDELIMITER & " and " & _
		"news_date <= " & DB_DATEDELIMITER & MediumDate(strEndDate) & DB_DATEDELIMITER & " and " & _
		"news_deleted = 0 " & _
		"ORDER BY news_date desc, news_heading"
	sql_GetNewsByDateSpanAdmin = sql
End Function

Function sql_GetTop3News()
	dim sql
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = "SELECT id, news_date, news_heading, news_URL, news_source, news_abstract " & _
		"FROM tbl_news WHERE news_deleted = 0 and news_live = " & PrepareBit(1) & " ORDER BY news_date desc LIMIT 3"
	else
		sql = "SELECT TOP 3 id, news_date, news_heading, news_URL, news_source, news_abstract " & _
		"FROM tbl_news WHERE news_deleted = 0 and news_live = " & PrepareBit(1) & " ORDER BY news_date desc"
	end if
	sql_GetTop3News = sql
End Function

Function sql_InsertNewsEvent(eventDateEntered, eventDate, eventHeading, eventURL, _
			eventSource, eventAbstract, eventContent, eventLive, eventDeleted, _
			eventEnteredBy)
	dim sql
	sql = "INSERT INTO tbl_news (news_dateentered, news_date, news_heading, news_url, " & _
		"news_source, news_abstract, news_content, news_live, news_deleted, " & _
		"news_enteredby) VALUES(" & _
		"" & DB_DATEDELIMITER & MediumDate(eventDateEntered) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(eventDate) & DB_DATEDELIMITER & ", " & _
		"'" & eventHeading & "', '" & eventURL & "', '" & eventSource & "', " & _
		"'" & eventAbstract & "', '" & eventContent & "', " & eventLive & ", " & _
		eventDeleted & ", " & eventEnteredBy & ")"
	sql_InsertNewsEvent = sql
End Function 	

Function sql_UpdateNewsEvent(eventDate, eventHeading, eventURL, eventSource, _
			eventAbstract, eventContent, eventLive, strEventID)
	dim sql
	sql = "UPDATE tbl_news SET " & _
		"news_date = " & DB_DATEDELIMITER & MediumDate(eventDate) & DB_DATEDELIMITER & ", " & _
		"news_heading = '" & eventHeading & "', " & _
		"news_url = '" & eventURL & "', " & _
		"news_source = '" & eventSource & "', " & _
		"news_abstract = '" & eventAbstract & "', " & _
		"news_content = '" & eventContent & "', " & _
		"news_live = " & eventLive & " " & _
	"WHERE id = " & strEventID & ""
	sql_UpdateNewsEvent = sql
End Function   

Function sql_DeleteNewsEvent(id)
	Dim sql
	sql = "UPDATE tbl_news set news_deleted = " & PrepareBit(1) & " " & _
		"WHERE id = " & id & ""
	sql_DeleteNewsEvent = sql
End Function  		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Projects
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetAllProjects()
	dim sql
	sql = "SELECT * FROM tbl_projects order by project_id"
	sql_GetAllProjects = sql
End Function

Function sql_GetProjectTypes()
	dim sql
	sql = "SELECT * FROM tbl_projecttypes ORDER BY projecttype_id"
	sql_GetProjectTypes = sql
End Function

Function sql_GetProjectTypeByID(id)
	dim sql
	sql = "SELECT * FROM tbl_projecttypes where projecttype_id = " & id & " ORDER BY projecttype_id"
	sql_GetProjectTypeByID = sql
End Function

Function sql_InsertProject(client_id, ProjectName, workorder_number, accountexec_id, status, _
	start_date, launch_date, projecttype_id, active, comments, color)
	dim sql
	sql = "INSERT into tbl_projects (client_id, description, workorder_number, accountexec_id, " & _
		"status, start_date, launch_date, projecttype_id, active, comments, color) VALUES (" & _
		client_id & "," & _
		"'" & ProjectName & "'," & _
		"'" & workorder_number & "'," & _
		accountexec_id & "," & _
		status & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(start_date) & DB_DATEDELIMITER & "," & _
		"" & DB_DATEDELIMITER & MediumDate(launch_date) & DB_DATEDELIMITER & "," & _
		projecttype_id & "," & _
		active & ", " & _
		"'" & comments & "'," & _
		"'" & color & "')"
	sql_InsertProject = sql
End Function

Function sql_InsertProjectSQLServer(client_id, ProjectName, workorder_number, accountexec_id, status, _
	start_date, launch_date, projecttype_id, active, comments, color)
	sql = "SET NOCOUNT ON; " 
	sql = sql & "INSERT into tbl_projects (client_id, description, workorder_number, accountexec_id, " & _
		"status, start_date, launch_date, projecttype_id, active, comments, color) VALUES (" & _
		client_id & "," & _
		"'" & ProjectName & "'," & _
		"'" & workorder_number & "'," & _
		accountexec_id & "," & _
		status & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(start_date) & DB_DATEDELIMITER & "," & _
		"" & DB_DATEDELIMITER & MediumDate(launch_date) & DB_DATEDELIMITER & "," & _
		projecttype_id & "," & _
		active & ", " & _
		"'" & comments & "'," & _
		"'" & color & "'); "
	sql = sql & "SELECT identityinsert=@@identity; SET NOCOUNT OFF;"
	sql_InsertProjectSQLServer = sql
End Function

Function sql_SelectLatestProjectID()
	dim sql
	sql = "SELECT MAX(project_id) as maxid FROM tbl_projects"
	sql_SelectLatestProjectID = sql
End Function

Function sql_SelectLargestWorkOrderNum() 
	dim sql 
	sql = "SELECT MAX(WorkOrder_Number) as maxnumber FROM tbl_projects" 
	sql_SelectLargestWorkOrderNum = sql 
End Function

Function sql_DeleteProjectBudgetsByID(id)
	dim sql
	sql = "DELETE FROM tbl_projects_budget WHERE project_id = " & id & ""
	sql_DeleteProjectBudgetsByID = sql
End Function

Function sql_InsertProjectsBudget(qID, empTypeID, empID, qRate, qHours) 
	dim sql
	sql = "INSERT INTO tbl_projects_budget (project_id, employeetype_id, employee_id, rate, hours) " & _
		" VALUES (" & _
		qID & ", " & _
		empTypeID & ", " & _
		empID & ", " & _
		DecimalCommaToPeriod(qRate) & ", " & _
		DecimalCommaToPeriod(qHours) & ")"
	sql_InsertProjectsBudget = sql
End Function

Function sql_GetProjectsByID(id)
	dim sql
	sql = "SELECT * FROM tbl_projects WHERE project_id = " & id & ""	
	sql_GetProjectsByID = sql
End Function

Function sql_GetProjectBudgetsByID(id)
	dim sql
	sql = "Select * from tbl_projects_budget where project_id = " & id & " order by employeeType_ID"
	sql_GetProjectBudgetsByID = sql
End Function

Function sql_UpdateProjects(client_id, description, workorder_number, accountexec_id, start_date, _
	launch_date, projecttype_id, active, comments, color, project_id)
	dim sql
	sql = "UPDATE tbl_projects SET " & _
		"client_id = " & client_id & "," & _
		"description = '" & description & "'," & _
		"workorder_number = '" & workorder_number & "'," & _
		"accountexec_id = " & accountexec_id & "," & _
		"start_date = " & DB_DATEDELIMITER & MediumDate(start_date) & DB_DATEDELIMITER & "," & _
		"launch_date = " & DB_DATEDELIMITER & MediumDate(launch_date) & DB_DATEDELIMITER & "," & _
		"projecttype_id = " & projecttype_id & "," & _
		"active = " & active & "," & _
		"comments = " & "'" & comments & "'," & _
		"color = " & "'" & color & "' " & _
		"where project_id = " & project_id & ""
	sql_UpdateProjects = sql
End Function

Function sql_GetProjectHoursByEmpType(Project,EmpType)
	dim sql
	sql = "SELECT SUM(timeamount) as hoursused " & _
		"FROM tbl_projects, tbl_timecards, tbl_timecardtypes, tbl_employees " & _
		"WHERE tbl_projects.project_id = " & Project & " and " & _
			"tbl_projects.project_id = tbl_timecards.project_id and " & _
			"tbl_timecards.employee_id = tbl_employees.employee_id and " & _
			"tbl_timecards.timecardtype_id = tbl_timecardtypes.timecardtype_id and " & _
			"tbl_timecards.[non-billable] = 0 and " & _
			"tbl_timecardtypes.employeetype_id = " & EmpType & ""
	sql_GetProjectHoursByEmpType = sql
End Function

Function sql_GetProjectHoursByPhase(Project, Phase)
	dim sql
	sql = "SELECT SUM(timeamount) as hoursused " & _
		"FROM tbl_projects, tbl_timecards " & _
		"WHERE tbl_projects.project_id = " & Project & " and " & _
		"tbl_projects.project_id = tbl_timecards.project_id and " & _
		"tbl_timecards.[non-billable] = 0 and tbl_timecards.projectphaseid = " & Phase & ""
	sql_GetProjectHoursByPhase = sql
End Function

Function sql_GetProjectPhases()
	dim sql
	sql = "SELECT * FROM tbl_projectphases ORDER BY projectphaseid"
	sql_GetProjectPhases = sql
End Function

Function sql_GetProjectPhasesByID(id)
	dim sql
	sql = "SELECT * FROM tbl_projectphases where projectphaseid = " & id & ""
	sql_GetProjectPhasesByID = sql
End Function

Function sql_InsertProjectPhase(pName)
	dim sql
	sql = "INSERT INTO tbl_projectphases (projectphasename) values ('" & pName & "')"
	sql_InsertProjectPhase = sql
End Function

Function sql_UpdateProjectPhase(pID, pName)
	dim sql
	sql = "UPDATE tbl_projectphases set projectphasename = '" & pName & "' " & _
		"WHERE projectphaseid = " & pID & ""
	sql_UpdateProjectPhase = sql
End Function

Function sql_DeleteProjectPhase(pID)
	dim sql
	sql = "DELETE FROM tbl_projectphases where projectphaseid = " & pID & ""
	sql_DeleteProjectPhase = sql
End Function

Function sql_GetProjectView(typeProject, boolAllTypes, active)
	dim sql
	sql = "SELECT tbl_clients.client_id as cl_id, * " & _
		"FROM tbl_projects, tbl_clients, tbl_projecttypes " & _
		"WHERE tbl_projects.client_id = tbl_clients.client_id and " & _
			"tbl_projects.projecttype_id = tbl_projecttypes.projecttype_id "
	activeString = "and tbl_projects.active=" & PrepareBit(1) & " "
	typeString = "and tbl_projecttypes.projecttypedescription = '" & typeProject & "' "  
	if not boolAllTypes then
		sql = sql & typeString
	end if
	if (active = "Active") then
		sql = sql & activeString
	end if
	sql = sql & "ORDER BY tbl_projects.projecttype_id, tbl_projects.launch_date, " & _
		"tbl_clients.client_name, tbl_projects.description"
	sql_GetProjectView = sql
End Function

Function sql_GetProjectViewByID(project_id)
	dim sql 
	sql = "SELECT tbl_clients.client_id as cl_id, * " & _
		"FROM tbl_projects, tbl_clients, tbl_projecttypes " & _
		"WHERE tbl_projects.client_id = tbl_clients.client_id and " & _
			"tbl_projects.projecttype_id = tbl_projecttypes.projecttype_id and " & _
			 "project_id = " & project_id & ""
	sql_GetProjectViewByID = sql
End Function

Function sql_GetActiveProjectsWithClients()
	dim sql
	sql = "SELECT tbl_clients.client_id as cl_id, tbl_clients.client_name, " & _
		"tbl_projects.client_id, tbl_projects.project_id, tbl_projects.description " & _
		"FROM tbl_clients " & _
		"INNER JOIN tbl_projects ON tbl_clients.client_id = tbl_projects.client_id " & _
		"WHERE tbl_projects.active=" & PrepareBit(1) & " and " & _
			"tbl_clients.active = " & PrepareBit(1) & " " & _
		"ORDER BY tbl_clients.client_name, tbl_projects.description"
	sql_GetActiveProjectsWithClients = sql
End Function

Function sql_GetAllProjectsWithClients()
	dim sql
	sql = "SELECT tbl_clients.client_name, tbl_projects.client_id, tbl_projects.project_id, tbl_projects.description " & _
		"FROM tbl_clients " & _
		"INNER JOIN tbl_projects ON tbl_clients.client_id = tbl_projects.client_id " & _
		"ORDER BY tbl_clients.client_name, tbl_projects.description"
	sql_GetAllProjectsWithClients = sql
End Function

Function sql_UpdateProjectForumAndFolder(project_id, forum_id, folder_id)
	dim sql 
	sql = "UPDATE tbl_projects SET forum_id = " & forum_id & ", folder_id = " & folder_id & " " & _
		"WHERE project_id = " & project_id & ""
	sql_UpdateProjectForumAndFolder = sql
End Function 

Function sql_InsertProjectQuote(project_id, dateentered, project_name, _
		client_name, account_rep, work_order, start_date, end_date, _
		project_type, comments, total_hours, total_price)
	dim sql
	sql = "Insert into tbl_projectquotes (" & _
		"project_id, dateentered, project_name, client_name, account_rep, work_order, " & _
		"start_date, end_date, project_type, comments, total_hours, total_price) values(" & _
	    project_id & ", " & _
	    "" & DB_DATEDELIMITER & MediumDate(dateentered) & DB_DATEDELIMITER & ", " & _ 
		"'" & project_name & "', " & _
		"'" & client_name & "', " & _
		"'" & account_rep & "', " & _
		"'" & work_order & "', " & _
		"" & DB_DATEDELIMITER & MediumDate(start_date) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(end_date) & DB_DATEDELIMITER & ", " & _
		"'" & project_type & "', " & _
		"'" & comments & "', " & _
		DecimalCommaToPeriod(total_hours) & ", " & _
		DecimalCommaToPeriod(total_price) & ")"
	sql_InsertProjectQuote = sql
End Function

Function sql_GetLatestProjectQuote()
	dim sql
	sql = "SELECT MAX(projectquote_id) as maxid FROM tbl_projectquotes"
	sql_GetLatestProjectQuote = sql
End Function

Function sql_DeleteProjectQuoteItems(id)
	dim sql
	sql = "Delete from tbl_projectquotesitems where projectquote_id = " & id & ""
	sql_DeleteProjectQuoteItems = sql
End Function

Function sql_DeleteProjectQuote(id)
	dim sql
	sql = "Delete from tbl_projectquotes where projectquote_id = " & id & ""
	sql_DeleteProjectQuote = sql
End Function

Function sql_InsertProjectQuoteItems(pqid, project_id, item_name, item_rate, item_hours, item_total)
	dim sql
	sql = "Insert into tbl_projectquotesitems (projectquote_id, project_id, " & _
		"itemname, itemrate, itemhours, itemtotal) values (" & _
		pqid & ", " & _
		project_id & ", " & _
		"'" & item_name & "', " & _
		DecimalCommaToPeriod(item_rate) & ", " & _
		DecimalCommaToPeriod(item_hours) & ", " & _
		DecimalCommaToPeriod(item_total) & ")"
	sql_InsertProjectQuoteItems = sql
End Function

Function sql_GetProjectQuoteByID(pqid)
	dim sql
	sql = "SELECT * from tbl_projectquotes where projectquote_id = " & pqid & ""
	sql_GetProjectQuoteByID = sql
End Function

Function sql_GetProjectQuotesByProjectID(pid)
	dim sql
	sql = "Select projectquote_id, dateentered, total_price from tbl_projectquotes where project_id = " & pid & ""
	sql_GetProjectQuotesByProjectID = sql
End Function

Function sql_GetProjectQuoteItemsByProjectQuoteID(pqid)
	dim sql
	sql = "select * from tbl_projectquotesitems where projectquote_id = " & pqid & " order by projectquoteitems_id"
	sql_GetProjectQuoteItemsByProjectQuoteID = sql
End Function

'## for shoeman's material quotes
Function sql_GetProjectQuoteMaterialsByID(pqid)
	dim sql
	sql = "SELECT * from tbl_projectquotesmaterials where projectquote_id = " & pqid & ""
	sql_GetProjectQuoteMaterialsByID = sql
End Function

Function sql_GetProjectQuotesMaterialsByProjectID(pid)
	dim sql
	sql = "Select projectquote_id, dateentered, total_price from tbl_projectquotesmaterials where project_id = " & pid & ""
	sql_GetProjectQuotesMaterialsByProjectID = sql
End Function

Function sql_GetProjectQuoteMaterialsItemsByProjectQuoteID(pqid)
	dim sql
	sql = "select * from tbl_projectquotesmaterialsitems where projectquote_id = " & pqid & " order by projectquotematerials_id"
	sql_GetProjectQuoteMaterialsItemsByProjectQuoteID = sql
End Function

Function sql_DeleteProjectQuoteMaterialsItems(id)
	dim sql
	sql = "Delete from tbl_projectquotesmaterialsitems where projectquote_id = " & id & ""
	sql_DeleteProjectQuoteMaterialsItems = sql
End Function

Function sql_DeleteProjectQuoteMaterials(id)
	dim sql
	sql = "Delete from tbl_projectquotesmaterials where projectquote_id = " & id & ""
	sql_DeleteProjectQuoteMaterials = sql
End Function

Function sql_InsertProjectQuoteMaterials(project_id, dateentered, project_name, _
		client_name, account_rep, work_order, tax_rate, discount_rate, legal, _
		start_date, end_date, project_type, comments, pre_discount_total, _
		discount_total, subtotal, tax_total, total_price)
	dim sql
	sql = "Insert into tbl_projectquotesmaterials (" & _
		"project_id, dateentered, project_name, client_name, account_rep, work_order, " & _
		"tax_rate, discount_rate, legal, start_date, end_date, project_type, " & _
		"comments, pre_discount_total, discount_total, subtotal, tax_total, " & _
		"total_price) values(" & _
	    project_id & ", " & _
	    "" & DB_DATEDELIMITER & MediumDate(dateentered) & DB_DATEDELIMITER & ", " & _ 
		"'" & project_name & "', " & _
		"'" & client_name & "', " & _
		"'" & account_rep & "', " & _
		"'" & work_order & "', " & _
		tax_rate & ", " & _
		discount_rate & ", " & _
		"'" & legal & "', " & _
		"" & DB_DATEDELIMITER & MediumDate(start_date) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(end_date) & DB_DATEDELIMITER & ", " & _
		"'" & project_type & "', " & _
		"'" & comments & "', " & _
		DecimalCommaToPeriod(pre_discount_total) & ", " & _
		DecimalCommaToPeriod(discount_total) & ", " & _
		DecimalCommaToPeriod(subtotal) & ", " & _
		DecimalCommaToPeriod(tax_total) & ", " & _
		DecimalCommaToPeriod(total_price) & ")"
	sql_InsertProjectQuoteMaterials = sql
End Function

Function sql_InsertProjectQuoteMaterialsItems(pqid, project_id, itemname, itemcomments, itemunitsprice, itemunits, itemtaxable,item_total)
	dim sql
	sql = "Insert into tbl_projectquotesmaterialsitems (projectquote_id, project_id, " & _
		"itemname, itemcomments, itemunitprice, itemunits, itemtaxable, itemtotal) values (" & _
		pqid & ", " & _
		project_id & ", " & _
		"'" & itemname & "', " & _
		"'" & itemcomments & "', " & _
		DecimalCommaToPeriod(itemunitsprice) & ", " & _
		DecimalCommaToPeriod(itemunits) & ", " & _
		PrepareBit(itemtaxable) & ", " & _
		DecimalCommaToPeriod(item_total) & ")"
	sql_InsertProjectQuoteMaterialsItems = sql
End Function

Function sql_GetLatestProjectQuoteMaterials()
	dim sql
	sql = "SELECT MAX(projectquote_id) as maxid FROM tbl_projectquotesmaterials"
	sql_GetLatestProjectQuoteMaterials = sql
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Resources
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetTop20Resources()
	dim sql
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = "SELECT * FROM tbl_resources " & _
			"WHERE resources_live = " & PrepareBit(1) & " and resources_deleted = 0 " & _
			"ORDER BY resources_title limit 20"
	else
		sql = "SELECT TOP 20 * FROM tbl_resources " & _
			"WHERE resources_live = " & PrepareBit(1) & " and resources_deleted = 0 " & _
			"ORDER BY resources_title"
	end if
	sql_GetTop20Resources = sql
End Function

Function sql_GetResources()
	dim sql
	sql = "SELECT * FROM tbl_resources " & _
			"WHERE resources_live = " & PrepareBit(1) & " and resources_deleted = 0 " & _
			"ORDER BY resources_title"
	sql_GetResources = sql
End Function

Function sql_GetAllResources()
	dim sql
	sql = "SELECT * FROM tbl_resources " & _
			"WHERE resources_deleted = 0 " & _
			"ORDER BY resources_title"
	sql_GetAllResources = sql
End Function

Function sql_InsertResource(eventDate, eventHeading, eventURL, eventAbstract, _
		eventLive, eventDeleted, eventEnteredBy)
	dim sql
	sql = "INSERT INTO tbl_resources (resources_date, resources_title, " & _
		"resources_URL, resources_caption, resources_live, resources_deleted, " & _
		"resources_enteredby) VALUES (" & _
		"" & DB_DATEDELIMITER & MediumDate(eventDate) & DB_DATEDELIMITER & ", " & _
		"'" & eventHeading & "', " & _
		"'" & eventURL & "', " & _
		"'" & eventAbstract & "', " & _
		PrepareBit(eventLive) & ", " & _
		PrepareBit(eventDeleted) & ", " & _
		eventEnteredBy & ")"
	sql_InsertResource = sql
End Function

Function sql_GetResourcesByID(strID)
	dim sql
	sql = "SELECT * FROM tbl_resources where resources_id = " & strID & ""
	sql_GetResourcesByID = sql
End Function

Function sql_UpdateResource(eventHeading, eventURL, eventAbstract, eventLive, strID)
	dim sql
	sql = "UPDATE tbl_resources set " & _
		"resources_title = '" & eventHeading & "', " & _
		"resources_url = '" & eventURL & "', " & _
		"resources_caption = '" & eventAbstract & "', " & _
		"resources_live = " & PrepareBit(eventLive) & " " & _
	"WHERE resources_id = " & strID & ""
	sql_UpdateResource = sql
End Function

Function sql_DeleteResource(strID)
	dim sql
	sql = "UPDATE tbl_resources SET resources_deleted = " & prepareBit(1) & " " & _
		"where resources_id = " & strID & ""
	sql_DeleteResource = sql
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Survey
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_CheckVoterIPByPollID(strIP, PollQuestion)
	dim sql
	sql = "SELECT * FROM tbl_pollresultsip " & _
		"WHERE fdpollresultsip = '" & strIP & "' AND fdpollid = " & PollQuestion & ""
	sql_CheckVoterIPByPollID = sql
End Function

Function sql_GetCurrentSurvey()
	dim sql
	if DB_BOOLMYSQL then
		sql = "SELECT * "
	else
		sql = "SELECT tbl_poll.fdpollid as fdpollid, * "
	end if
	sql = sql & "FROM tbl_pollcurrent, tbl_poll " & _
		"WHERE rtrim(tbl_pollcurrent.fdpollid) = rtrim(tbl_poll.fdpollid)"
	sql_GetCurrentSurvey = sql
End Function

Function sql_GetAllSurveys()
	dim sql
	sql = "SELECT * FROM tbl_poll ORDER BY fdpolltopic"
	sql_GetAllSurveys = sql
End Function

Function sql_GetCurrentSurveyRecords()
	dim sql
	sql = "SELECT * FROM tbl_pollcurrent"
	sql_GetCurrentSurveyRecords = sql
End Function

Function sql_GetAnswersByQuestionID(question_id)
	dim sql
	sql = "SELECT * FROM tbl_pollanswer " & _
		"WHERE fdpollquestionid = " & question_id & " " & _
		"ORDER BY fdpollanswerid"
	sql_GetAnswersByQuestionID = sql
End Function

Function sql_GetPollByID(Poll)
	dim sql
	sql = "SELECT * FROM tbl_poll WHERE fdpollid = " & Poll & ""
	sql_GetPollByID = sql
End Function

Function sql_GetQuestionsByPollID(polltopic)
	dim sql
	sql = "SELECT * FROM tbl_pollquestion " & _
		"WHERE fdpollid = " & polltopic & " " & _
		"ORDER BY fdpollquestionid"
	sql_GetQuestionsByPollID = sql
End Function

Function sql_GetAnswerByID(answer)
	dim sql
	sql = "SELECT * FROM tbl_pollanswer WHERE fdpollanswerid = " & answer & ""
	sql_GetAnswerByID = sql
End Function

Function sql_GetQuestionByID(question)
	dim sql
	sql = "SELECT * FROM tbl_pollquestion WHERE fdpollquestionid = " & question & ""
	sql_GetQuestionByID = sql
End Function

Function sql_GetQuestionByPollIDQuestionValue(poll, question)
	dim sql
	sql = "SELECT * from tbl_pollquestion " & _
		"WHERE fdpollid = " & poll & " and fdpollquestion = '" & question & "' " & _
		"ORDER by fdpollquestionid desc"
	sql_GetQuestionByPollIDQuestionValue = sql
End Function

Function sql_GetSurveyResultsByPollID(poll)
	dim sql
	sql = "SELECT * FROM tbl_pollresultsip WHERE fdpollid = " & poll & " " & _
		"ORDER BY fdpollresultsipdate"
	sql_GetSurveyResultsByPollID = sql
End Function

Function sql_GetSurveyResultsByID(entry)
	dim sql
	sql = "SELECT * FROM tbl_pollresultsip WHERE fdpollresultsipid = " & entry & ""
	sql_GetSurveyResultsByID = sql
End Function

Function sql_InsertCurrentSurvey(poll)
	dim sql
	sql = "INSERT INTO tbl_pollcurrent (fdpollid) VALUES (" & poll & ")"
	sql_InsertCurrentSurvey = sql
End Function

Function sql_InsertSurvey(poll)
	dim sql
	sql = "INSERT INTO tbl_poll (fdpolltopic) VALUES ('" & poll & "')"
	sql_InsertSurvey = sql
End Function

Function sql_InsertPollResults(polltopic, strIP, strDate, PollResult, recpt)
	dim sql
	sql = "INSERT INTO tbl_pollresultsip " & _
		"(fdpollid, fdpollresultsip, fdpollresultsipdate, fdpollresults, fdpollresultsemail) " & _
		"VALUES (" & _
		polltopic & ", " & _
		"'" & strIP & "', " & _
		"" & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & ", " & _
		"'" & PollResult & "', " & _
		"'" & recpt & "')"
	sql_InsertPollResults = sql
End Function

Function sql_InsertPollAnswer(question, answer, num)
	dim sql
	sql = "INSERT INTO tbl_pollanswer (fdpollquestionid, fdpollanswer, fdpollanswerresult) " & _
		"VALUES ('" & question & "', '" & answer & "', " & num & ")"
	sql_InsertPollAnswer = sql
End Function

Function sql_InsertQuestion(poll, question , num, int_type)
	dim sql
	sql = "INSERT INTO tbl_pollquestion " & _
		"(fdpollid, fdpollquestion, fdpollquestionresult, fdpollquestiontype) " & _
		"values ('" & _
		poll & "', " & _
		"'" & question & "', " & _
		num & ", " & _
		int_type & ")"
	sql_InsertQuestion = sql
End Function

Function sql_UpdateCurrentSurvey(poll)
	dim sql
	sql = "UPDATE tbl_pollcurrent SET fdpollid = " & poll & " where 1=1"
	sql_UpdateCurrentSurvey = sql
End Function

Function sql_UpdatePollResults(answerID)
	dim sql
	sql = "UPDATE tbl_pollanswer SET fdpollanswerresult = fdpollanswerresult + 1 " & _
		"WHERE fdpollanswerid = " & answerID & ""
	sql_UpdatePollResults = sql
End Function

Function sql_UpdatePollQuestionResults(question_id)
	dim sql
	sql = "UPDATE tbl_pollquestion SET fdpollquestionresult = fdpollquestionresult + 1 " & _
		"WHERE fdpollquestionid = " & question_id & ""
	sql_UpdatePollQuestionResults = sql
End Function

Function sql_UpdateAnswer(answer_text, answer)
	dim sql
	sql = "UPDATE tbl_pollanswer SET " & _
		"fdpollanswer = '" & answer_text & "' " & _
		"WHERE fdpollanswerid = " & answer & ""
	sql_UpdateAnswer = sql
End Function

Function sql_UpdateSurveyQuestion(question_text, int_type, question)
	dim sql
	sql = "UPDATE tbl_pollquestion SET " & _
		"fdpollquestion = '" & question_text & "', " & _
		"fdpollquestiontype = " & int_type & " " & _
		"WHERE fdpollquestionid = " & question & ""
	sql_UpdateSurveyQuestion = sql
End Function

Function sql_DeleteSurveyAnswerByAnswerID(answer)
	dim sql
	sql = "DELETE FROM tbl_pollanswer WHERE fdpollanswerid = " & answer & ""
	sql_DeleteSurveyAnswerByAnswerID = sql
End Function

Function sql_DeleteSurveyAnswerByQuestionID(question)
	dim sql
	sql = "DELETE FROM tbl_pollanswer WHERE fdpollquestionid = " & question & ""
	sql_DeleteSurveyAnswerByQuestionID = sql
End Function

Function sql_DeleteSurveyQuestionByQuestionID(question)
	dim sql
	sql = "DELETE FROM tbl_pollquestion WHERE fdpollquestionid = " & question & ""
	sql_DeleteSurveyQuestionByQuestionID = sql
End Function

Function sql_DeleteSurveyQuestionByPollID(poll)
	dim sql
	sql = "DELETE FROM tbl_pollquestion WHERE fdpollid = " & poll & ""
	sql_DeleteSurveyQuestionByPollID = sql
End Function

Function sql_DeleteSurveyResultsByPollID(poll)
	dim sql
	sql = "DELETE FROM tbl_pollresultsip WHERE fdpollid = " & poll & ""
	sql_DeleteSurveyResultsByPollID = sql
End Function

Function sql_DeleteSurvey(poll)
	dim sql
	sql = "DELETE FROM tbl_poll WHERE fdpollid = " & poll & ""
	sql_DeleteSurvey = sql
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Tasks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetAllTasks()
	dim sql
	sql = "SELECT * FROM tbl_tasks order by task_id"
	sql_GetAllTasks = sql
End Function

Function sql_GetTaskHoursByEmployee(ID)
	dim sql
	sql = "SELECT SUM (estimatedhours) AS curhours " & _
		"FROM tbl_tasks " & _
		"WHERE tbl_tasks.assignedto = " & ID & " and "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "(tbl_tasks.done=0 and tbl_tasks.showit=" & PrepareBit(1) & ")"	
	else
		sql = sql & "(tbl_tasks.done=0 and tbl_tasks.[show]=" & PrepareBit(1) & ")"
	end if
	sql_GetTaskHoursByEmployee = sql
End function

Function sql_GetTaskHoursCompleteByEmployee(ID)
	dim sql
	sql = "SELECT sum(tbl_timecards.timeamount) as timecharged " & _
		"FROM tbl_timecards, tbl_tasks " & _
		"WHERE tbl_timecards.employee_id = " & ID & " and " & _
			"tbl_timecards.task_id = tbl_tasks.task_id and "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "(tbl_tasks.done = 0 and tbl_tasks.showit=" & PrepareBit(1) & ")"
	else
		sql = sql & "(tbl_tasks.done = 0 and tbl_tasks.[show]=" & PrepareBit(1) & ")"
	end if
	sql_GetTaskHoursCompleteByEmployee = sql
End Function 

Function sql_DeleteTasks(id)
	dim sql
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = "UPDATE tbl_tasks SET showit=0 WHERE task_id=" & id & ""
	else
		sql = "UPDATE tbl_tasks SET [show]=0 WHERE task_id=" & id & ""
	end if
	sql_DeleteTasks = sql
End Function

Function sql_UndeleteTasks(id)
	dim sql
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = "UPDATE tbl_tasks SET showit=1 WHERE task_id=" & id & ""
	else
		sql = "UPDATE tbl_tasks SET [show]=1 WHERE task_id=" & id & ""
	end if 
	sql_UndeleteTasks = sql
End Function

Function sql_CompleteTasks(id, action)
	dim sql
	sql = "UPDATE tbl_tasks SET done = " & action & " WHERE task_id = " & id & ""
	sql_CompleteTasks = sql
End Function

Function sql_GetTasksByID(id)
	dim sql
	sql = "SELECT * FROM tbl_tasks WHERE task_id=" & id & ""
	sql_GetTasksByID = sql
End Function

Function sql_UpdateTasks(client_ID, project_ID, description, priority, orderedBy, assignedTo, _
	estimatedHours, dateCreated, dateDue, taskDone, id)
	dim sql
	sql = "UPDATE tbl_tasks SET " & _
		"client_id = " & client_ID & ", " & _
		"project_id = '" & project_ID & "', " & _
		"description = '" & description & "', " & _
		"priority = " & priority & ", " & _
		"orderedby = '" & orderedBy & "', " & _
		"assignedto = '" & assignedTo & "', " & _
		"estimatedhours = " & DecimalCommaToPeriod(estimatedHours) & ", " & _
		"datecreated = " & DB_DATEDELIMITER & MediumDate(dateCreated) & DB_DATEDELIMITER & ", " & _
		"datedue = " & DB_DATEDELIMITER & MediumDate(dateDue) & DB_DATEDELIMITER & ", " & _
		"done = " & taskDone & " " & _
		"WHERE task_id = " & id & ""
	sql_UpdateTasks = sql
End Function

Function sql_InsertTask(description, priority, orderedBy, assignedTo, estimatedHours, _
	client_ID, project_ID, dateCreated, timeCreated, done, show, DateDue)
	dim sql
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = "INSERT into tbl_tasks (description, priority, orderedby, assignedto, estimatedhours, " & _
		"client_id, project_id, datecreated, timecreated, done, showit, datedue) VALUES ("
	else
		sql = "INSERT into tbl_tasks (description, priority, orderedby, assignedto, estimatedhours, " & _
		"client_id, project_id, datecreated, timecreated, done, [show], datedue) VALUES ("	
	end if		
	sql = sql & "'" & description & "', " & _
		priority & ", " & _
		orderedBy & ", " & _
		assignedTo & ", " & _
		DecimalCommaToPeriod(estimatedHours) & ", " & _
		client_ID & ", " & _
		project_ID & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(dateCreated) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & timeCreated & DB_DATEDELIMITER & ", " & _
		done & ", " & _
		show & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(dateDue) & DB_DATEDELIMITER & ")"
	sql_InsertTask = sql 
End Function

Function sql_GetTaskDoneEmailInfoByProjectID(id)
	dim sql
	sql = "SELECT tbl_employees.employeename AS worker, " & _
			"tbl_employees.emailaddress AS worker_email, " & _
			"tbl_employeesb.employeename AS producer, " & _
			"tbl_employeesb.emailaddress, tbl_clients.client_name, " & _
			"tbl_tasks.description, tbl_tasks.datecreated, " & _
			"tbl_tasks.estimatedhours, tbl_tasks.done, tbl_tasks.task_id " & _
		"FROM (tbl_clients INNER JOIN " & _
			"(tbl_tasks INNER JOIN tbl_employees ON tbl_tasks.assignedto = tbl_employees.employee_id) " & _
			"ON tbl_clients.client_id = tbl_tasks.client_id) " & _
			"INNER JOIN tbl_employees AS tbl_employeesb " & _
			"ON tbl_tasks.orderedby = tbl_employeesb.employee_id " & _
		"WHERE tbl_tasks.task_id = " & id & ""
	sql_GetTaskDoneEmailInfoByProjectID = sql
End Function

Function sql_GetTaskViewAll(active)
	dim sql
	sql = "SELECT tbl_employees.employeename as employeename_1, tbl_employees_1.employeename, * " & _
	"FROM tbl_clients INNER JOIN " & _
		"((tbl_tasks INNER JOIN tbl_employees ON tbl_tasks.assignedto = tbl_employees.employee_id) " & _
		"INNER JOIN tbl_employees AS tbl_employees_1 " & _
		"ON tbl_tasks.orderedby = tbl_employees_1.employee_id) " & _
		"ON tbl_clients.client_id = tbl_tasks.client_id "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "WHERE tbl_tasks.showit=" & PrepareBit(active) & " "
	else
		sql = sql & "WHERE tbl_tasks.[show]=" & PrepareBit(active) & " "
	end if
	sql = sql & "ORDER BY tbl_clients.client_name, tbl_tasks.priority desc, tbl_tasks.datedue"
	sql_GetTaskViewAll = sql
End Function

Function sql_GetTaskViewByClientID(active, id)
	dim sql
	sql = "SELECT tbl_employees.employeename as employeename_1, tbl_employees_1.employeename, * " & _
	"FROM tbl_clients INNER JOIN " & _
		"((tbl_tasks INNER JOIN tbl_employees ON tbl_tasks.assignedto = tbl_employees.employee_id) " & _
		"INNER JOIN tbl_employees as tbl_employees_1 " & _
		"ON tbl_tasks.orderedby = tbl_employees_1.employee_id) " & _
		"ON tbl_clients.client_id = tbl_tasks.client_id "
	sql = sql & "WHERE tbl_tasks.client_id = " & id & " and " 
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "tbl_tasks.showit=" & PrepareBit(active) & " "
	else
		sql = sql & "tbl_tasks.[show]=" & PrepareBit(active) & " "
	end if
	sql = sql & "ORDER BY tbl_tasks.priority desc, tbl_tasks.datedue"
	sql_GetTaskViewByClientID = sql
End Function

Function sql_GetTaskViewByProjectID(active, id)
	dim sql
	sql = "SELECT tbl_tasks.description as desc1, tbl_employees.employeename as employeename_1, " & _
		"tbl_employees_1.employeename, * " & _
	"FROM tbl_projects INNER JOIN " & _
		"((tbl_tasks INNER JOIN tbl_employees ON tbl_tasks.assignedto = tbl_employees.employee_id) " & _
		"INNER JOIN tbl_employees as tbl_employees_1 " & _
		"ON tbl_tasks.orderedby = tbl_employees_1.employee_id) " & _
		"ON tbl_projects.project_id = tbl_tasks.project_id " & _
	"WHERE tbl_tasks.project_id = " & id & " and " 
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "tbl_tasks.showit=" & PrepareBit(active) & " "
	else
		sql = sql & "tbl_tasks.[show]=" & PrepareBit(active) & " "	
	end if
	sql = sql & "ORDER BY tbl_tasks.priority desc, tbl_tasks.datedue"
	sql_GetTaskViewByProjectID = sql
End Function

Function sql_GetTaskViewByAssignor(active)
	dim sql
	sql = "SELECT tbl_employees.employeename AS employeename_1, tbl_employees_1.employeename, * " & _
	"FROM tbl_clients INNER JOIN " & _
		"((tbl_tasks INNER JOIN tbl_employees ON tbl_tasks.assignedto = tbl_employees.employee_id) " & _
		"INNER JOIN tbl_employees AS tbl_employees_1 " & _
		"ON tbl_tasks.orderedby = tbl_employees_1.employee_ID) " & _
		"ON tbl_clients.client_id = tbl_tasks.client_id "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "WHERE tbl_tasks.showit=" & PrepareBit(active) & " "
	else
		sql = sql & "WHERE tbl_tasks.[show]=" & PrepareBit(active) & " "
	end if
	sql = sql & "ORDER BY tbl_employees_1.employeename, tbl_tasks.priority desc, tbl_tasks.datedue, tbl_clients.client_name"
	sql_GetTaskViewByAssignor = sql
End Function

Function sql_GetTaskViewByEmployees(active)
	dim sql
	sql = "SELECT tbl_employees.employeename AS employeename_1, tbl_employees_1.employeename, * " & _
		"FROM tbl_clients INNER JOIN " & _
		"((tbl_tasks INNER JOIN tbl_employees ON tbl_tasks.assignedto = tbl_employees.employee_id) " & _
		"INNER JOIN tbl_employees AS tbl_employees_1 " & _
		"ON tbl_tasks.orderedby = tbl_employees_1.employee_id) " & _
		"ON tbl_clients.client_id = tbl_tasks.client_id "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "WHERE tbl_tasks.showit=" & PrepareBit(active) & " "
	else
		sql = sql & "WHERE tbl_tasks.[show]=" & PrepareBit(active) & " "
	end if
	sql = sql & "ORDER BY tbl_employees.employeename, tbl_tasks.priority desc, tbl_tasks.datedue, tbl_clients.client_name"
	sql_GetTaskViewByEmployees = sql
End Function

Function sql_GetTaskViewByEmployeeID(active, id)
	dim sql
	sql = "SELECT tbl_employees.employeename AS employeenameassignedto, " & _
		"tbl_employees_1.employeename as employeenameorderedby, " & _
		"tbl_employees_1.emailaddress as emailaddressorderedby , " & _
		"tbl_projects.description as projectname, tbl_projects.project_id as project, " & _
		"tbl_clients.client_id as client, tbl_projecttypes.projecttypedescription, * " & _
	"FROM ((tbl_projects INNER JOIN ((tbl_tasks INNER JOIN tbl_employees ON " & _
		"tbl_tasks.assignedto = tbl_employees.employee_id) INNER JOIN " & _
		"tbl_employees as tbl_employees_1 ON tbl_tasks.orderedby = tbl_employees_1.employee_id) " & _
		"ON tbl_projects.project_id = tbl_tasks.project_id) INNER JOIN " & _
		"tbl_clients ON tbl_projects.client_id = tbl_clients.client_id) INNER JOIN " & _
		"tbl_projecttypes ON tbl_projects.projecttype_id = tbl_projecttypes.projecttype_id " & _
	"WHERE tbl_clients.client_id=tbl_projects.client_id AND " & _
		"tbl_tasks.project_id=tbl_projects.project_id AND " 
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "tbl_tasks.showit=" & PrepareBit(active) & " AND "
	else
		sql = sql & "tbl_tasks.[show]=" & PrepareBit(active) & " AND "
	end if
	sql = sql & "tbl_tasks.assignedto=" & id & " " & _
	"ORDER BY tbl_tasks.done asc, tbl_tasks.priority DESC, tbl_tasks.datedue, tbl_clients.client_name"
	sql_GetTaskViewByEmployeeID = sql
End Function

Function sql_GetTaskViewByAssignorID(active, id)
	dim sql
	sql = "SELECT tbl_employees.employeename AS employeenameassignedto, " & _
		"tbl_employees_1.employeename as employeenameorderedby, " & _
		"tbl_employees_1.emailaddress as emailaddressorderedby , " & _
		"tbl_projects.description as projectname, tbl_projects.project_id AS project, " & _
		"tbl_clients.client_id as client, tbl_projecttypes.projecttypedescription, * " & _
	"FROM ((tbl_projects INNER JOIN ((tbl_tasks INNER JOIN tbl_employees ON " & _
		"tbl_tasks.assignedto = tbl_employees.employee_id) INNER JOIN " & _
		"tbl_employees as tbl_employees_1 ON tbl_tasks.orderedby = tbl_employees_1.employee_id) " & _
		"ON tbl_projects.project_id = tbl_tasks.project_id) INNER JOIN " & _
		"tbl_clients ON tbl_projects.client_id = tbl_clients.client_id) INNER JOIN " & _
		"tbl_projecttypes ON tbl_projects.projecttype_id = tbl_projecttypes.projecttype_id " & _
	"WHERE tbl_clients.client_id=tbl_projects.client_id AND " & _
		"tbl_tasks.project_id=tbl_projects.project_id AND " 
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "tbl_tasks.showit=" & PrepareBit(active) & " AND "
	else
		sql = sql & "tbl_tasks.[show]=" & PrepareBit(active) & " AND "
	end if	
	sql = sql & "tbl_tasks.orderedby=" & id & " " & _
	"ORDER BY tbl_tasks.done asc, tbl_tasks.priority DESC, tbl_tasks.datedue, tbl_clients.client_name"
	sql_GetTaskViewByAssignorID = sql
End Function

Function sql_GetTaskListAndCount(active)
	dim sql
	sql = "SELECT Count(tbl_tasks.task_id) AS countofid, tbl_clients.client_id, " & _
		"tbl_clients.client_name " & _
	"FROM tbl_tasks INNER JOIN tbl_clients " & _
		"ON tbl_tasks.client_id = tbl_clients.client_id "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "WHERE tbl_tasks.showit=" & PrepareBit(active) & " "
	else
		sql = sql & "WHERE tbl_tasks.[show]=" & PrepareBit(active) & " "
	end if
	sql = sql & "GROUP BY tbl_clients.client_id, tbl_clients.client_name " & _
	"ORDER BY Count(tbl_tasks.task_id) DESC , tbl_clients.client_name ASC"
	sql_GetTaskListAndCount = sql
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Timecards
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetAllTimecards()
	dim sql
	sql = "SELECT * FROM tbl_timecards order by timecard_id"
	sql_GetAllTimecards = sql
End Function

Function sql_GetActiveTimecardTypes()
	dim sql
	sql = "SELECT * FROM tbl_timecardtypes WHERE active = " & PrepareBit(1) & " " & _
	"ORDER BY sortorder, timecardtypedescription"
	sql_GetActiveTimecardTypes = sql
End Function

Function sql_GetTimecardTypes()
	dim sql
	sql = "SELECT * FROM tbl_timecardtypes ORDER BY sortorder, timecardtypedescription"
	sql_GetTimecardTypes = sql
End Function

Function sql_GetTimecardTypesByID(strID)
	dim sql
	sql = "SELECT * FROM tbl_timecardtypes where timecardtype_id = " & strID & ""
	sql_GetTimecardTypesByID = sql
End Function

Function sql_InsertTimeCardType(strType, strActive, strEmpType)
	dim sql
	sql = "INSERT INTO tbl_timecardtypes (timecardtypedescription, active, sortorder, employeetype_id) " & _
		"VALUES ('" & strType & "', " & strActive & ", 15, " & strEmpType & ")"
	sql_InsertTimeCardType = sql
End Function

Function sql_UpdateTimeCardType(strID, strType, strActive, strEmpType)
	dim sql
	sql = "UPDATE tbl_timecardtypes SET " & _
		"timecardtypedescription = '" & strType & "', " & _
		"active = " & strActive & ", " & _
		"employeetype_id = " & strEmpType & " " & _
		"WHERE timecardtype_id = " & strID & ""
	sql_UpdateTimeCardType = sql
End Function

Function sql_DeleteTimecardType(strID)
	dim sql
	sql = "DELETE FROM tbl_timecardtypes where timecardtype_id = " & strID & ""
	sql_DeleteTimecardType = sql
End Function

Function sql_GetTimecardHoursTodayByEmployeeID(id)
	sql = "SELECT SUM(tbl_timecards.timeamount) as timecharged, " & _
		"COUNT(tbl_timecards.timeamount) as numtimecards " & _
	"FROM tbl_timecards " & _
	"WHERE tbl_timecards.employee_ID = " & id & " and " & _
		"dateworked = " & DB_DATEDELIMITER & MediumDate(date()) & DB_DATEDELIMITER & ""
	sql_GetTimecardHoursTodayByEmployeeID = sql
End Function

Function sql_GetTimecardHoursByTaskID(id)
	dim sql
	sql = "SELECT SUM(tbl_timecards.timeamount) as timecharged " & _
	"FROM tbl_timecards WHERE tbl_timecards.task_id = " & id & ""
	sql_GetTimecardHoursByTaskID = sql
End Function

Function sql_GetActiveTaskProjectClientList()
	dim sql
	sql = "SELECT tbl_tasks.description as taskdesc, tbl_clients.client_id as client_id, " & _
		"tbl_projects.project_id as project_id, * " & _
	"FROM tbl_tasks, tbl_clients, tbl_projects " & _
	"WHERE tbl_clients.client_id = tbl_projects.client_id and " & _
		"tbl_projects.project_id = tbl_tasks.project_id and " & _
		"tbl_clients.client_id = tbl_tasks.client_id and tbl_clients.active = " & PrepareBit(1) & " and " & _
		"tbl_projects.active = " & PrepareBit(1) & " "
	if DB_BOOLPOSTGRES or DB_BOOLMYSQL then
		sql = sql & "and tbl_tasks.showit = " & PrepareBit(1) & " "
	else
		sql = sql & "and tbl_tasks.[show] = " & PrepareBit(1) & " "	
	end if
	sql = sql & "ORDER BY tbl_clients.client_name, tbl_projects.description"
	sql_GetActiveTaskProjectClientList = sql
End Function

Function sql_InsertTimecard(Client_ID, Project_ID, Task_ID, projectPhaseID, strEmployeeID, _
	TimeCardType_ID, TimeAmount, WorkDescription, dateEntered, dateWorked, dateLastEdited, _
	Employee_ID, lngTimeCardRate, lngTimeCardCost, reconciled, NonBillable)
	dim sql
	sql = "Insert INTO tbl_timecards (client_id, project_id, task_id, projectphaseid, " & _
		"employee_id, timecardtype_id, timeamount, workdescription, dateentered, dateworked, " & _
		"datelastedited, lasteditedby, rate, cost, reconciled, [non-billable]) Values(" & _
	Client_ID & ", " & _
	Project_ID & ", " & _
	Task_ID & ", " & _
	projectPhaseID & ", " & _
	strEmployeeID & ", " & _
	TimeCardType_ID & ", " & _
	DecimalCommaToPeriod(TimeAmount) & ", " & _
	"'" & WorkDescription & "', " & _
	"" & DB_DATEDELIMITER & MediumDate(dateEntered) & DB_DATEDELIMITER & ", " & _
	"" & DB_DATEDELIMITER & MediumDate(dateWorked) & DB_DATEDELIMITER & ", " & _
	"" & DB_DATEDELIMITER & MediumDate(dateLastEdited) & DB_DATEDELIMITER & ", " & _
	Employee_ID & ", " & _
	DecimalCommaToPeriod(lngTimeCardRate) & ", " & _
	DecimalCommaToPeriod(lngTimeCardCost) & ", " & _
	reconciled & ", " & _
	NonBillable & ")"
	sql_InsertTimecard = sql
End Function

Function sql_DeleteTimecard(id)
	dim sql
	sql = "DELETE FROM tbl_timecards WHERE timecard_id = " & ID & ""
	sql_DeleteTimecard = sql
End Function

Function sql_GetTimecardsByID(id)
	dim sql
	sql = "SELECT * FROM tbl_timecards WHERE timecard_id = " & id & ""
	sql_GetTimecardsByID = sql
End Function

Function sql_UpdateTimecards(Client_ID, Project_ID, Task_ID, Employee_ID, TimeCardType_ID, _
	TimeAmount, WorkDescription, DateWorked, DateLastEdited, LastEditedBy, lngTimeCardRate, _
	lngTimeCardCost, projectPhaseID, NonBillable, Reconciled, timecard_id)
	dim sql
	sql = "UPDATE tbl_timecards SET " & _
		"client_id = " & Client_ID & "," & _
		"project_id = " & Project_ID & "," & _
		"task_id = " & Task_ID & "," & _
		"employee_id = " & Employee_ID & "," & _
		"timecardtype_id = " & TimeCardType_ID & "," & _
		"timeamount = " & DecimalCommaToPeriod(TimeAmount) & "," & _
		"workdescription = '" & WorkDescription & "'," & _
		"dateworked = " & DB_DATEDELIMITER & MediumDate(DateWorked) & DB_DATEDELIMITER & "," & _
		"datelastedited = " & DB_DATEDELIMITER & MediumDate(DateLastEdited) & DB_DATEDELIMITER & "," & _
		"lasteditedby = " & LastEditedBy & "," & _
		"rate = " & DecimalCommaToPeriod(lngTimeCardRate) & "," & _
		"cost = " & DecimalCommaToPeriod(lngTimeCardCost) & "," & _
		"projectphaseid = " & projectPhaseID & "," & _
		"[non-billable] = " & NonBillable & ", " & _
		"[reconciled] = " & Reconciled & " " & _
		"WHERE timecard_id = " & timecard_id & ""
	sql_UpdateTimecards = sql
End Function

Function sql_GetTimecardViewByEmployeeID(id, dateToStart)
	dim sql
	sql = "SELECT tbl_timecards.timecard_id, tbl_timecards.dateworked, " & _
		"tbl_clients.client_name, tbl_projects.description, tbl_timecards.workdescription, " & _
		"tbl_timecards.timeamount " & _
	"FROM tbl_projects RIGHT JOIN (tbl_clients INNER JOIN tbl_timecards ON " & _
		"tbl_clients.client_id = tbl_timecards.client_id) ON " & _
		"tbl_projects.project_id = tbl_timecards.project_id " & _
	"WHERE tbl_timecards.dateworked > " & DB_DATEDELIMITER & MediumDate(dateToStart) & DB_DATEDELIMITER & " AND " & _
		"employee_id = " & id & " ORDER BY tbl_timecards.dateworked DESC"
	sql_GetTimecardViewByEmployeeID = sql
End Function		
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Timesheets
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function sql_GetAllTimesheets()
	dim sql
	sql = "SELECT * FROM tbl_timesheets order by id"
	sql_GetAllTimesheets = sql
End Function

Function sql_GetTimesheetsByID(id)
	dim sql
	sql = "SELECT * FROM tbl_timesheets WHERE id = " & id & ""
	sql_GetTimesheetsByID = sql
End Function

Function sql_GetTimesheetsByDate(strDate, id)
	dim sql
	sql = "SELECT * FROM tbl_timesheets " & _
	"WHERE employee_id = " & id & " and " & _
		"datehourslogged = " & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & ""
	sql_GetTimesheetsByDate = sql
End Function

Function sql_InsertTimesheets(employee_id, strDate, strTime, LoggedIn)
	dim sql
	sql = "INSERT INTO tbl_timesheets (employee_id,datehourslogged,in1,loggedin) VALUES (" & _
		employee_id & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & strTime & DB_DATEDELIMITER & ", " & _
		LoggedIn & ")"
	sql_InsertTimesheets = sql
End Function

Function sql_UpdateTimesheetsNow(in2, out1, out2, employee_id) 
	dim sql	
	sql = "UPDATE tbl_timesheets SET "
	if len("***" & out1) < 4 then
		sql = sql & "out1 = " & DB_DATEDELIMITER & Time() & DB_DATEDELIMITER & ","
		sql = sql & "loggedin = 0"
		session("loginAction") = "Logout"
	elseif len("***" & in2) < 4 then
		sql = sql & "in2 = " & DB_DATEDELIMITER & Time() & DB_DATEDELIMITER & ","
		sql = sql & "loggedin = 1"
		session("loginAction") = "Login"
	elseif len("***" & out2) < 4 then
		sql = sql & "out2 = " & DB_DATEDELIMITER & Time() & DB_DATEDELIMITER & ","
		sql = sql & "loggedin = 0"
		session("loginAction") = "Logout"
	elseif len("***" & out2) > 4 then
		sql = "0"
	end if
	if sql <> "0" then
		sql = sql & " WHERE employee_id = " & employee_id & " and " & _
			"datehourslogged = " & DB_DATEDELIMITER & MediumDate(Date()) & DB_DATEDELIMITER & ""
	end if
	sql_UpdateTimesheetsNow = sql
End Function

Function sql_UpdateTimesheets(datehourslogged, in1, in2, out1, out2, id)
	dim sql
	sql = "UPDATE tbl_timesheets SET " & _
		"datehourslogged = " & DB_DATEDELIMITER & MediumDate(datehourslogged) & DB_DATEDELIMITER & ", "
	if isDate(in1) then
		sql = sql & "in1 = " & DB_DATEDELIMITER & in1 & DB_DATEDELIMITER & ","
	else
		sql = sql & "in1 = NULL, "
	end if
	if isDate(in2) then
		sql = sql & "in2 = " & DB_DATEDELIMITER & in2 & DB_DATEDELIMITER & ","
	else
		sql = sql & "in2 = NULL, "
	end if
	if isDate(out1) then
		sql = sql & "out1 = " & DB_DATEDELIMITER & out1 & DB_DATEDELIMITER & ","
	else
		sql = sql & "out1 = NULL, "
	end if
	if isDate(out2) then
		sql = sql & "out2 = " & DB_DATEDELIMITER & out2 & DB_DATEDELIMITER & " "
	else
		sql = sql & "out2 = NULL "
	end if
	sql = sql & " where id = " & id & ""
	sql_UpdateTimesheets = sql
End Function

Function sql_GetPayPeriods()
	dim sql
	sql = "SELECT * FROM tbl_payperiods ORDER BY enddate desc"
	sql_GetPayPeriods = sql
End Function

Function sql_GetPayPeriodsByID(strID)
	dim sql
	sql = "SELECT * FROM tbl_payperiods WHERE id = " & strID & ""
	sql_GetPayPeriodsByID = sql
End Function

Function sql_GetCurrentPayPeriod()
	dim sql
	sql = "SELECT * FROM tbl_payperiods " & _
	"WHERE startdate <= " & DB_DATEDELIMITER & MediumDate(Date()) & DB_DATEDELIMITER & " and " & _
		"enddate >= " & DB_DATEDELIMITER & MediumDate(Date()) & DB_DATEDELIMITER & ""
	sql_GetCurrentPayPeriod = sql
End Function

Function sql_InsertPayPeriod(pStartDate, pEndDate)
	dim sql
	sql = "INSERT INTO tbl_payperiods (startdate, enddate) values (" & _
		"" & DB_DATEDELIMITER & MediumDate(pStartDate) & DB_DATEDELIMITER & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(pEndDate) & DB_DATEDELIMITER & ")"
	sql_InsertPayPeriod = sql
End Function

Function sql_UpdatePayPeriod(pID, pStartDate, pEndDate)
	dim sql
	sql = "UPDATE tbl_payperiods set startdate = " & DB_DATEDELIMITER & MediumDate(pStartDate) & DB_DATEDELIMITER & ", " & _
		"enddate = " & DB_DATEDELIMITER & MediumDate(pEndDate) & DB_DATEDELIMITER & " " & _
		"WHERE id = " & pID & ""
	sql_UpdatePayPeriod = sql
End Function

Function sql_DeletePayPeriod(pID)
	dim sql
	sql = "DELETE FROM tbl_payperiods where id = " & pID & ""
	sql_DeletePayPeriod = sql
End Function

Function sql_GetPayPeriodsByEndDate(datLastDayOfPreviousPeriod)
	dim sql
	sql = "SELECT * FROM tbl_payperiods " & _
	"WHERE enddate = " & DB_DATEDELIMITER & MediumDate(datLastDayOfPreviousPeriod) & DB_DATEDELIMITER & ""
	sql_GetPayPeriodsByEndDate = sql
End Function

Function sql_GetPayPeriodsByDate(dat)
	dim sql
	sql = "SELECT * FROM tbl_payperiods " & _
	"WHERE enddate >= " & DB_DATEDELIMITER & MediumDate(dat) & DB_DATEDELIMITER & " " & _
		"AND startdate <= " & DB_DATEDELIMITER & MediumDate(dat) & DB_DATEDELIMITER & ""
	sql_GetPayPeriodsByDate = sql
End Function

Function sql_GetPayPeriodsBetweenStartDateEndDate(dat,datEnd)
	dim sql
	sql = "SELECT * FROM tbl_payperiods " & _
	"WHERE startdate >= " & DB_DATEDELIMITER & MediumDate(dat) & DB_DATEDELIMITER & " " & _
		"AND enddate <= " & DB_DATEDELIMITER & MediumDate(datEnd) & DB_DATEDELIMITER & " " & _
	"ORDER BY startdate desc "
	sql_GetPayPeriodsBetweenStartDateEndDate = sql
End Function

Function sql_GetFirstPayPeriod()
	dim sql
	sql = "SELECT * FROM tbl_payperiods ORDER BY startdate asc"
	sql_GetFirstPayPeriod = sql
End Function

Function sql_GetTimesheetsByEmployeeIDEndDate(empID, datLastDayOfCurrentPeriod)
	dim sql
	sql = "SELECT * FROM tbl_timesheets " & _
	"WHERE employee_id = " & empID & " and " & _
		"datehourslogged <= " & DB_DATEDELIMITER & MediumDate(datLastDayOfCurrentPeriod) & DB_DATEDELIMITER & " " & _
	"ORDER by datehourslogged asc"
	sql_GetTimeSheetsByEmployeeIDEndDate = sql
End Function

Function sql_GetTimesheetsByEmployeeIDStartDate(empID, datFirstDayOfCurrentPeriod)
	dim sql
	sql = "SELECT * FROM tbl_timesheets " & _
	"WHERE employee_id = " & empID & " and " & _
		"datehourslogged >= " & DB_DATEDELIMITER & MediumDate(datFirstDayOfCurrentPeriod) & DB_DATEDELIMITER & " " & _
	"ORDER by datehourslogged asc"
	sql_GetTimeSheetsByEmployeeIDStartDate = sql
End Function

Function sql_GetTimesheetsByEmployeeIDStartDateEndDate(empID, datFirstDayOFPreviousPeriod, _
	datLastDayOfPreviousPeriod)
	dim sql
	sql = "SELECT * FROM tbl_timesheets " & _
	"WHERE employee_id = " & empID & " and " & _
		"datehourslogged >= " & DB_DATEDELIMITER & MediumDate(datFirstDayOfPreviousPeriod) & DB_DATEDELIMITER & " and " & _	
		"datehourslogged <= " & DB_DATEDELIMITER & MediumDate(datLastDayOfPreviousPeriod) & DB_DATEDELIMITER & " " & _
	"ORDER by datehourslogged asc"
	sql_GetTimesheetsByEmployeeIDStartDateEndDate = sql
End Function

Function sql_GetFirstTimesheetForEmployeeID(empID)
	dim sql
	sql = "SELECT * FROM tbl_timesheets " & _
	"WHERE employee_id = " & empID & " " & _
	"ORDER BY datehourslogged asc"
	sql_GetFirstTimesheetForEmployeeID = sql
End Function

Function sql_InsertLoginAction(loginAction, employee_id, strDate, strIP)
	dim sql	
	sql = "INSERT INTO tbl_loginaction (act, employee_id, act_date, ip) Values (" & _
		"'" & loginAction & "', " & _
		employee_id & ", " & _
		"" & DB_DATEDELIMITER & MediumDate(strDate) & DB_DATEDELIMITER & ", " & _
		"'" & strIP & "')"
	sql_InsertLoginAction = sql
End Function
%>
