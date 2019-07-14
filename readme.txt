
Welcome to Transworld Interactive Projects.

This errata file is included with the Transworld Interactive open source distribution and will be augmented over time.

Download and/or use of this software is implied consent to the Transworld Interactive Open Source License.


INSTALLATION INSTRUCTIONS:
Unzip this distribution zip into the empty root of your website or into the subfolder that this intranet will occupy.

Edit the /includes/SiteConfig.asp file with your favorite text editor.  You must at least set the constant gsSiteRoot for the site to minimally work.

Default login and password for Intranet Open Source is test/pass.  Changing the default login and password is highly recommended.

Once you have the Intranet up and running update the client listing (id #1) to reflect your company's information.  Do not delete client #1's record, just replace your information with the existing information.

View the Administrator's Guide for detailed installation and operating instructions.

WHAT's NEW

version 2.6.1 - 9/18/2002
- bug fix on page /projects/project-quote.asp - fixed where the proper client name is not showing up, changed lines 361 and 362 from <%=rsClientList("Client_Name")%> to <%=GetClientName(rsProject("Client_id"))%> and added the function GetClientName() to the bottom of the page.
- added site configuration section to the admin
	- added folder /admin/config
	- added files /admin/config/default.asp and selcolor.htm
	- updated file /admin/default.asp to have the link to the site config page 
	- used the permDatabaseSetup property as the property for editing the site configuration rather than added a new permissions parameter.
	- the site config admin uses the file system object to actually make the updates to the SiteConfig.asp file - the file system object must be enabled to use this feature and the SiteConfig.asp must have proper permissions assigned to it to allow IUSR_Machinename to edit the file.
	- made updates to the Languages.asp file - added a new section for admin/site configuration - it is commented as such in the file.
	- when updating from a previous version of intranet open source to this version you may have to go in and hand edit your siteconfig.asp file a bit. there must be at least one space between the variable name and the equals sign for the site config admin to work. previous versions of the siteconfig.asp file may have only had tabs between each configuration variable and its equals sign.  you will know this is the problem with your config file if when you go to edit a variable in the admin and it says it was updated but it actually was not.

version 2.6.0 - 8/12/2002
- bug fix on line 183 of the file /includes/MTS_forums.asp, changed the less than signs to be less than or equal to.  This fix the problem with not showing messages under the forum entries.
- bug fix on line 378 of the file /projects/dashboard.asp, changed the path for the URL, it was wrong before
- made change to line 148 of file /logoncheck.asp - added a new conditional to check to make sure the session variable is not null and the length of the variable is not 4 or less before doing a "left" on it.
- updated /includes/connection_open.asp to include error trapping functionality and to print error to screen.
- commented out the link to the database setup wizard from the admin.  It needs to be rebuilt to take advantage of PostgresSQL and MySQL configuration lines in the siteconfig.asp file.
- updated SiteConfig.asp - added database config lines for MySQL database.  Full MySQL compatibility is not in place yet - some of the sql queries need to be changed.
- added external news url capability to news
	- added 2 fields to tbl_news: news_url text 255, news_source text 255
	- updated file /includes/news.asp to take advantage of new fields
	- updated file /news/default.asp to take advantage of new fields
	- update functions in file SiteSQL.asp that get top 3 news items for home page, add news and edit news
	- added a term in the languages.asp file for the news "source" field.
- add resources - an additional option for the home page
	- added options gsHomeResources and gsResources to SiteConfig.asp
	- updated database: added new table tbl_resources
	- updated database: added new field to tbl_permissions: permAdminResources Yes/No default 0
	- updated siteSQL.asp file with new section of functions for resources
	- updated siteSQL.asp file, function sql_UpdatePermissions to use new permission permAdminResources
	- updated /global.asa added line to set permAdminResources default permission to 0
	- added Resources section to the languages.asp file
	- created new /includes/resources.asp file that is an include to the home page
	- update /main.asp to add new resources section, added resources under the tools section in the nav links at the bottom
	- added folder /resources/
	- added file /resources/default.asp
	- updated file /includes/navspans.asp - added nav option under the tools section for resources.  also commented out "options" since that has not been implemented yet.
	- updated file /admin/default.asp - added link to resources
	- updated file /admin/default.asp (employee admin pak) - to add new permAdminResources in permissions types
	- updated file /admin/hnd_edit.asp to make use of new permissions permAdminResources
	- updated file /logoncheck.asp to account for new permissions permAdminResources
	- added folder /admin/resources/
	- added file /admin/resources/default.asp
- beginning with this version all updates to languages.asp file will be identified	


version 2.5.1 - 7/23/2002
- added additional permissions item for permission to update employee permissions also created application variables that hold default employee permissions used when adding an employee
	- added global.asa - creates application variables that hold default employee permissions - we will be making more use of this file as time goes by - we will be putting site configuration info in here - should help performance a bit.
	- default.asp - created onload function where focus is set to the login box
	- logoncheck.asp - setup new permissions session variable, session("permAdminEmployeesPerms")
	- updated languages.asp, added two new items under the admin/employees section
	- sitesql.asp - added a new function: sql_SelectLatestEmployeeID in the employees section and updated the function sql_UpdatePermissions to use the new field permAdminEmployeesPerms
	- added a field to the table tbl_permissions: permAdminEmployeesPerms, Y/N, default:0
- also made some updates to the Employee Admin Pak to use the new permissions variable and default permissions set in the global.asa


version 2.5.0 - 7/18/2002
- changed this readme.txt file to go from newest changes to oldest
- added landing pages for the main task links, a default.asp page in each of the folders with links to the functionality
	- added /tools/ folder
	- added default.asp page to each of the sections: /timecards, /tasks, /projects, /clients, /reports, /tools
	- added default.asp page in /timesheets/ that redirects to /timesheets/view.asp
	- changed the top nav item for each of the above section from just a label to a link to the default page
- made the DHTML part of the navigation separate from the top navigation
	- added a new site variable to the siteconfig page, gsDHTMLNavigation
	- made updates to js.asp to exclude the dhtml javascript functions if they are not needed
	- made updates to nav.asp to exclude the dhtml javascript function calls if they are not needed
	- made updates to navSpans.asp to exclude the page content if it is not needed
- changed the variable <%=dictLanguage("AccountRep")%> to be <%=dictLanguage("Account_Rep")%> on line 184 of /projects/project-status.asp
- added database viewing tools and export utilities to admin page
	- added the following items to languages.asp under the Admin Text section: Database_Export_Utils, Table_View, Excel, CDTF, SQL_Query, No_records_exist, Error_in_SQL
	- added the following functions to the SiteSQL.asp file: sql_GetAllProjects, sql_GetAllTasks, sql_GetAllTimecards, sql_GetAllTimesheets, sql_GetAllNews, sql_GetAllEvents
	- updated the /admin/default.asp page to have links to the different view and export pages
	- added the folder /admin/exports
	- added the following pages to the /admin/exports folder: calendar.asp, clients.asp, employees.asp, news.asp, projects.asp, tasks.asp, timecards.asp, and variable.asp
- added ability to report by reconciled timecards to "Work Log Internal" and "Work Log for Client"
	- made updates to the /reports/ client_work_log.asp, client_work_log_for_client.asp, client_work_log_results.asp, and client_work_log_results_for_client.asp pages


version 2.4.0 - 7/10/2002
- added functionality to save, view, print, and delete quotes for a project.
	- added page: /projects/project-quote.asp
	- made updates to the following pages: 
		/projects/dashboard.asp - added section to create and view project quotes
		/projeces/project-edit.asp - added items to create and view project quotes
		/includes/languages.asp - added a few items relating to project quotes
		/includes/sitesql.asp - added new functions for project quotes
		/includes/style.asp - added style to not display form buttons when page is printed
		/includes/popup_page_close.asp - made the "Close this window" link not displayed when page is printed by adding a class to the <p> that surrounds it.
		/includes/siteconfig.asp - added new variable gsProjectQuotes
	- added the following database tables: 
		tbl_projectquotes
		tbl_projectquotesitems
- added function DecimalCommaToPeriod() to the connection_open.asp file
- used the function several places in the SiteSQL.asp file where input numbers could be decimals, particularly insert/edit timecards, insert/edit tasks, and insert project estimates
- updated the pages client_work_log_results.asp and client_work_log_results_for_client.asp to use a dynamic count for employees rather than the hard-coded 50 and 150.
_ added functionality for work orders to be autogenerated by autoincrementing
	- added new function to SiteSQL.asp page, sql_SelectLargestWorkOrderNum() 
    - added new constant to SiteConfig.asp page, gsDynWorkOrderNumber
    - added new section of code to the page, /projects/project-add.asp

    
version 2.3.4 - 7/5/2002
- updated line 1803 in sitesql.asp page.  removed medium date from around variable lasteditedby.


version 2.3.3 - 7/3/2002
- updated bug on line 334 of /includes/MTS_forums.asp to put MediumDate function around the variable dTimeStamp
- fixed problem with js.asp where if you set an option (like timecards) to not be used it created a javascript error.


version 2.3.2 - 7/2/2002.
- made autologin of returning users via cookies an option in the siteConfig.asp file.  updated /default.asp and /logoncheck.asp to use that option.
- added password reminder feature.  added page /forgotpassword.asp.  updated pages: /default.asp, /includes/languages.asp and /includes/sitesql.asp


version 2.3.1 - 7/1/2002.
- fixed misspelling of column name in siteSQL.asp file


version 2.3 - 7/1/2002.
- added international date support:
	1. added function MediumDate() to file /includes/connection_open.asp, removed function ShortDateToLonger()
	2. updated SiteSQL.asp file by adding MediumDate function to all dates used in database queries
	3. made updates to the following pages that involved fixing date formatting that was causing errors or unexpected results when using international date formats:
		/					popupcalendar.asp
		/admin/forum/		admin.asp, admin_forum_display.asp
		/calendar/			default.asp, event.asp
		/clients/			addmessage.asp, client-add.asp, client-edit.asp, quickadd.asp, view_log.asp
		/employees/			default.asp, directory.asp, directory_download.asp
		/forum/				default.asp, search.asp
		/includes/			calendar.asp
		/includes/			connection_open.asp, forum_common.asp, main_page_open.asp, mts_forums.asp, news.asp, siteconfig.asp, sitesql.asp, style.asp
		/news/				default.asp, event.asp
		/projects/			dashboard.asp, project-add.asp, project-edit.asp
		/reports/			client_work_log_results.asp, client_work_log_results_for_client.asp, production_report_grid.asp, production_report_grid_download.asp
		/repository/		default.asp
		/tasks/				edit.asp, new.asp
		/timecards/			timecard-edit.asp, timecard.asp
		/timesheets/		edit.asp, view2.asp, view3.asp
- fixed problems with viewing/editing timesheets, made changes to: /timesheets/edit.asp, view2.asp, view3.asp
- began adding support for PostgreSQL by adding a boolean variable to the siteconfig.asp page and using that variable in preparing sql statments in the SiteSQL.asp file.  PostgreSQL support is not completed.  SQL syntax and recordset return values need to be validated.


version 2.2 - 6/20/2002.
- added multilanguage support:
    1. added language.asp file that basically is assigns all the text in the site to variables.
    2. updated nearly all the pages by removing the text and replacing it with the variables.
    3. included the language.asp file into the bottom of the SiteConfig.asp file.
    4. changing the language is as simple as editing the language.asp file or downloading the language.asp file for the language you want from the Transworld Interactive site and replacing the default language.asp file with it.
    5. as soon as additional language files are created or contributed they will be available for download from the free downloads page of the Transworld Interactive site.
- fixed problems with /timesheets/view3.asp page where previous pay periods would not be displayed if there were no entries for a more recent one.


version 2.1.1 - 6/13/2002.
- updated /timesheets/view.asp and /timesheets/view2.asp.  page was erroring out when there was no current pay period set in the database.


version 2.1 - 6/11/2002
- added project dashboard page: /projects/dashboard.asp
- updated siteSQL.asp page with new statements
- updated navspans.asp to show new project dashboard nav item
- added project dashboard variable to the config file
- updated forum_common.asp to use links with the directory name rather than just the page name.  this is so the functions could be used in the project dashboard page.
- commented out the cache headers ("Pragma", "no-cache" & "cache-control", "no-store") in the main_page_header.asp file because I was getting tired of going back to pages and getting the page has expired message.  You can uncomment these if you like.


version 2.0.2 - 6/9/2002
- added constant gsTaskEmails to the SiteConfig.asp page that can be set to allow/deny the sending of emails when tasks are marked created/completed/deleted.
- Fixed bug on line 85 of /tasks/view_employee.asp


version 2.0.1 - 6/4/2002
- Fixed bug where news items marked as not live still show up on the home page.
- Fixed bug in sql functions for adding and editing employees where a field name, password, is a reserved word and was not surrounded with brackets.
- Added the newest SiteSQL.asp file with all the newest functions.


version 2.0 - first commercial release - 5/15/2002
