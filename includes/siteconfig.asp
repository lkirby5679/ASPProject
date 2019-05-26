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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'General Site Configuration Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
gsColorHighlight            = "#6699CC"					'light color used throughout the site
gsColorBackground           = "#444444"					'darker color used throughout the site and for text
gsColorAltHighlight         = "#e7d19a"					'alt light color used throughout the site
const gsColorWhite          = "#eeeeee"					'alt color for white used throughout the site, especially in alternating rows
const gsColorBlack          = "#111111"					'alt color used for black throughout the site, black is so harsh sometimes

const gsMetaDescription     = "Transworld interactive Project Management Site. Only current authorized personel have access to this site."						'Meta tag description
const gsMetaKeywords        = "transworld project manager,project,projects,manager,management,transworld interactive,transworld,ti,project software"						'Meta tag keywords

const gsSiteName            = "Transworld Interactive Project Manager"
const gsSiteLogo            = "tilogo.gif"			'stored in the images folder
const gsAdminEmail          = "webmaster@transworldinteractive.net"
const gsMailHost            = "mail.transworldinteractive.net"
const gsSiteURL             = "http://project.transworldinteractive.net/"  'full site url, including any folders
const gsSiteRoot            = "/"						'the place in the site structure where this site is
														'if this site is the root then "/" else if this site
														'is in a subfolder like "project" then you should put "/project/" here
														'make sure to include the trailing slash
const gsMailComponent       = "CDONTS"					'CDONTS is the default component coded into the initial release
														'other valid choices include: CDO, ASPXP, PersistsASPMail, Jmail, and ServerObjectsASPMail
														'other components can be easily added to the /includes/mail.asp file

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Site Functionality Options
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
const gsNavigation  = TRUE    'whether to use the top nav or not
const gsDHTMLNavigation  = TRUE    'whether to use the DHTML of the navigation or not
const gsTimecards  = TRUE    'Use Timecards or not.  If using timecards Clients and Projects must be enabled.
const gsTasks  = TRUE    'Use Tasks or not. If using tasks Clients and Projects must be enabled.
const gsProjects  = TRUE    'Use Projects or not
const gsClients  = TRUE    'Use Clients or not
const gsClientContactLog  = TRUE    'Use Client Contact Log or not.  You must use clients to use this.
const gsTimeSheets  = TRUE    'Use Timesheets for hourly employees or not
const gsProductionReportGrid  = TRUE    'Use the production report grid or not. You must use clients, projects, and timecards to use this.
const gsReports  = TRUE    'Use Reports or not. You must use clients, projects and timecards to use this.
const gsCalendar  = TRUE    'Use Calendar or not
const gsDiscussion  = TRUE    'Use Discussion Forum or not
const gsEmployeeDirectory  = TRUE    'Use Employee Directory or not
const gsFileRepository  = TRUE    'Use File Repository or not
const gsNews  = TRUE    'Use News or not
const gsResources  = TRUE    'Use Resources or not
const gsProjectDashboard  = TRUE    'Use Project Dashboard page or not.
const gsProjectQuotes  = TRUE    'Use Project Quotes or not.
const gsOptions  = TRUE    'Use User defined options or not, they can change colors, etc
const gsPTO  = TRUE    'Use PTO or not. 
const gsInventory           = FALSE          'Use Inventory Module or not. for future use

const gsHostNonBillable  = FALSE    'Is Time charged agains the host company Non-billable?

const gsAutoLogin  = TRUE    'Does the site Auto Login the user if their cookie exists?
const gsTaskEmails  = TRUE    'Send Task Emails
const gsDynWorkOrderNumber  = TRUE    'If True all work order numbers will be generated on the fly 
											'NOTE: the very first work order number must be set by you
const gsCalendarClients  = TRUE    'Show Client Start date in Calendar	
const gsCalendarProjects  = TRUE    'Show Project Start and end dates in calendar 
const gsCalendarTasks  = TRUE    'Show Uncompleted Task due dates in calendar for assignor and assignee 

'this is set automagically by the above parameters
if gsCalendar or gsDiscussion or gsEmployeeDirectory or gsFileRepository or gsNews or gsOptions or gsResources then
	gsOther = TRUE
else
	gsOther = FALSE
end if

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Settings for Home Page
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
const gsHomeMinimize            = TRUE			'whether the feature boxes on the home page are able to be collapsed/expanded
const gsHomeThoughts            = TRUE			'which of the features on the home page do you want turned on?				
const gsHomeProductionReport    = TRUE
const gsHomeNews                = TRUE
const gsHomeResources           = TRUE
const gsHomeCalendar            = TRUE
const gsHomeDiscussion          = TRUE
const gsHomeSurvey              = TRUE
const gsHomeExtNews             = FALSE

const gsHomeExtNewsUS           = FALSE			'these are options for the included news feed
const gsHomeExtNewsWorld        = FALSE			'these options are what categories you would
const gsHomeExtNewsTech         = TRUE			'like displayed
const gsHomeExtNewsWeb          = TRUE
const gsHomeExtNewsBusinesss    = FALSE
const gsHomeExtNewsScience      = FALSE
const gsHomeExtNewsSports       = FALSE
const gsHomeExtNewsGaming       = FALSE
const gsHomeExtNewsLinux        = TRUE
const gsHomeExtNewsLinuxGaming  = FALSE
const gsHomeExtNewsLinuxSoftware = FALSE
const gsHomeExtNewsMac          = FALSE
const gsHomeExtNewsMacGaming    = FALSE
const gsHomeExtNewsMacSoftware  = FALSE

const gsHomeExtNewsShowTitle    = "yes"
const gsHomeExtNewsShowTime     = "no"
const gsHomeExtNewsShowIndent   = "no"
const gsHomeExtNewsShowNumber   = "3"
const gsHomeExtNewsShowNewWindow = "yes"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DATABASE CONFIGURATION
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'if using sql server please fill in the folling. if not comment them out.
'DB_DATASOURCE = "ServerName" 'Data Source or Server name
'DB_INITIALCATALOG = "intranet" 'initial database to log into
'DB_USERNAME = "intranet"
'DB_PASSWORD = "password"
'DB_DATEDELIMITER			= "'"
'DB_BOOLACCESS				= FALSE			'this is used for bit/yes/no data fields to convert them to negative for access
'This will be created on its own from above parameters
'DB_CONNECTIONSTRING = "Provider=SQLOLEDB;Data Source=" & DB_DATASOURCE & ";Initial Catalog=" & DB_INITIALCATALOG & ";Network Library=DBMSSOCN;User ID=" & DB_USERNAME & ";Password=" & DB_PASSWORD & ";"

'!!!DO NOT EDIT THE LINE BELOW!!!! - It is used as a reference point for scripts that parse and edit this file.
'if using access please fill in the following.  if not comment them out.
'!!!DO NOT EDIT THE LINE ABOVE!!!!
DB_DATASOURCENAME			= "tiproject.mdb" 'name of mdb file - should be stored in the data folder outside of the htdocs folder

' DB_DATASOURCE				= Server.MapPath(gsSiteRoot & "data/" & DB_DATASOURCENAME & "") connection string one

DB_DATASOURCE				= "C:\Inetpub\wwwroot\transworldproject\db\" & DB_DATASOURCENAME & "" ' connection string two

DB_DATEDELIMITER			= "#"
DB_BOOLACCESS				= TRUE			'this is used for bit/yes/no data fields to convert them to negative for access
DB_BOOLMYSQL				= FALSE
DB_BOOLPOSTGRES				= FALSE
'This will be created on its own from the above parameters
'DB_CONNECTIONSTRING		= "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_DATASOURCE & "" goes with connection string one

DB_CONNECTIONSTRING = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & DB_DATASOURCE & ";"  ' goes with connection string two

'Alt Access connection string: DB_CONNECTIONSTRING = "DSN=intranet"   'DSN must first be created

'if using PostgreSQL please fill in the following.  if not comment them out and set DB_BOOLPOSTGRES to FALSE
'DB_CONNECTIONSTRING		= "PostgreSQLIntranet"	'this uses a DSN previously created for the PostgrSQL database
'DB_BOOLACCESS				= FALSE
'DB_BOOLPOSTGRES			= TRUE
'DB_BOOLMYSQL				= FALSE
'DB_DATEDELIMITER			= "'"

'if using MySQL please fill in the following.  if not comment them out.
'DB_CONNECTIONSTRING = "intranet_MySQL"	'this uses a DSN previously created for the MySQL database
'DB_BOOLMYSQL				= TRUE
'DB_BOOLPOSTGRES				= FALSE
'DB_BOOLACCESS				= FALSE
'DB_DATEDELIMITER			= "'"
%>

<!-- #include file="languages.asp"-->
