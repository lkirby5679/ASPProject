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
<!--#include file="includes/SiteConfig.asp"-->

<%
'check to see if the database is setup
On Error Resume Next
strConn = DB_CONNECTIONSTRING	
set conn=server.createobject("adodb.connection")
conn.Open strConn
If conn.errors.count <> 0 then
'	Response.Redirect gsSiteRoot & "admin/dbsetup.asp"
end if
conn.close
set conn = nothing
On Error GoTo 0
	
strOption = Request.QueryString("opt")

if gsAutoLogin then
	''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Enable cookie based logon if possible
	'''''''''''''''''''''''''''''''''''''''
	if request.Cookies("shark_user_id") <> "" and request.Cookies("shark_password") <> "" and strOption="" then
		response.redirect "logoncheck.asp"
	end if
end if
%>

<!--#include file="includes/style.asp"-->
<!--Transworld Interactive Projects-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">
<HTML>
<HEAD>
<TITLE>Project | Website Project | Transworld interactive Project | Transworld Interactive Project</TITLE>
<META name="resource-type" content="document">
<META name="robots" content="index,follow">
<META name="Revisit-After" content="1 days">
<META name="Author" content="Professor Landon Kirby,Ph.D,D.Sc,DBA">
<META name="keywords" content="project,projects,transworld interactive project,project development,project management,hosting,design">
<META name="description" content="Transworld Interactive for Project Hosting, Project Design, Project management, Datacenter. Transworld Interactive also does Network Intergration.">
<META name="distribution" content="global">
<META name="copyright" content="Transworld Interactive website copyright 2013 by Professor Landon Kirby,Ph.D,D.Sc,DBA">
<meta name="no-email-collection" content="http://www.transworldinteractive.net/nospambots.htm">
<META NAME="page-topic" CONTENT="Transworld Interactive Project Management Website Hosting, Project Management.">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1251">
<LINK href="css.css" type=text/css rel=stylesheet>
<script src="AC_RunActiveContent.js" type="text/javascript"></script>
</HEAD>
<BODY OnLoad="document.form1.userid.focus();" LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0> 
<TABLE BORDER=0 WIDTH=780 HEIGHT=100% CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
<TR VALIGN=TOP>
	<TD>
		<TABLE BORDER=0 WIDTH=100% HEIGHT=100% CELLPADDING=0 CELLSPACING=0>
		<TR>
			<TD>
				<TABLE BORDER=0 WIDTH=100% HEIGHT=100% CELLPADDING=0 CELLSPACING=0>
				<TR HEIGHT=152>
					<TD>
<script type="text/javascript">
AC_FL_RunContent( 
'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','width','238','height','152','src','9ai5business.swf','quality','high','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','9ai5business','menu','false'); //end AC code
</script>
<noscript>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="238" height="152">
                        <param name="movie" value="9ai5business.swf">
                        <param name="quality" value="high"><param name="LOOP" value="false">
                        <embed src="9ai5business.swf" width="238" height="152" loop="false" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash"></embed>
			        </object></noscript></TD>
				</TR>
				<TR><TD><IMG SRC="i/ban1.gif" WIDTH="238" HEIGHT="76" BORDER="0"></TD></TR>
				<TR><TD><IMG SRC="i/ban2.gif" WIDTH="238" HEIGHT="76" BORDER="0"></TD></TR>
				<TR><TD><IMG SRC="i/h_wellcome.gif" WIDTH="238" HEIGHT="43" BORDER="0"></TD></TR>
				<TR HEIGHT=100% BGCOLOR=#B30000>
					<TD STYLE="PADDING:10px 10px 10px 9px;">
						<DIV id=news>
							<P><IMG SRC="i/pic1.jpg" WIDTH="208" HEIGHT="63" BORDER="0"></P>
							<P><B>Transworld Interactive Project Division</B></P>
							<P><b>Is proud to present to you our new Project Management site. Our Project Manager software is available for our clients to purchase or to lease on their websites. Project manager can be installed on our or your server. Project Manager is an open source program if purchased and can be customized to suit any needs that our clients need</b></P>		
						</DIV>
						<DIV ALIGN=RIGHT><A HREF="" STYLE="COLOR:#000000; FONT-SIZE:11px;"><B>Read more<IMG SRC="i/arrow2.gif" WIDTH="15" HEIGHT="11" BORDER="0" ALIGN=absmiddle></B></A></DIV>
					</TD>
				</TR>
				<TR BGCOLOR=#B30000 ALIGN=RIGHT><TD><A HREF=""><IMG SRC="i/ban3.gif" WIDTH="230" HEIGHT="79" BORDER="0"></A></TD></TR>
				<TR BGCOLOR=#B30000 ALIGN=RIGHT><TD><A HREF=""><IMG SRC="i/ban4.gif" WIDTH="230" HEIGHT="79" BORDER="0" VSPACE=10><br>
				  <script language=JavaScript src="http://livesupport.transworldinteractive.net/als.asp"></script>				</A></TD></TR>
				</TABLE></TD>
			<TD WIDTH=100%>
				<TABLE BORDER=0 WIDTH=100% HEIGHT=100% CELLPADDING=0 CELLSPACING=0>
				<TR>
					<TD>
						<TABLE BORDER=0 WIDTH=100% CELLPADDING=0 CELLSPACING=0>
						<TR VALIGN=TOP>
							<TD><IMG SRC="i/head1.jpg" WIDTH="304" HEIGHT="152" BORDER="0"></TD>
							<TD WIDTH=100% STYLE="BACKGROUND:url(i/top_bg.gif) no-repeat 0px 0px;">
								<DIV ALIGN=CENTER><A HREF=""><IMG SRC="i/ico1.gif" WIDTH="26" HEIGHT="24" BORDER="0"></A><A HREF="about.html"><IMG SRC="i/ico2.gif" WIDTH="29" HEIGHT="24" BORDER="0"></A><A HREF=""><IMG SRC="i/ico3.gif" WIDTH="30" HEIGHT="24" BORDER="0"></A></DIV><BR>
								<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 ALIGN=CENTER>
								<TR>
									<TD STYLE="COLOR:#B30000;"><B>
						    <DIV STYLE="FONT-SIZE:12px;">Call US TODAY! (24/7)     </DIV>
									    <div align="center" class="header"> 1-800-824-0426									    </div>
									</B> </TD>
									<TD><a href="mailto:project@transworldinteractive.net"><IMG SRC="i/call_bg.gif" alt="Email Transworld Interactive" WIDTH="66" HEIGHT="63" BORDER="0"></a></TD>
								</TR>
								</TABLE>
								<DIV STYLE="FONT-SIZE:11px; PADDING:5px 0px 0px 35px;"><B>Email:</B> <a href="mailto:project@transworldinteractive.net">Project Support</a></DIV></TD>
						</TR>
						</TABLE></TD>
				</TR>
				<TR>
					<TD>
						<TABLE BORDER=0 WIDTH=100% CELLPADDING=0 CELLSPACING=0>
						<TR>
							<TD><IMG SRC="i/head2.jpg" WIDTH="304" HEIGHT="304" BORDER="0"></TD>
							<TD WIDTH=100% STYLE="BACKGROUND:url(i/m_bg.gif) repeat-x 0px 0px;" ALIGN=RIGHT>
								<TABLE BORDER=0 WIDTH=100% HEIGHT=304 CELLPADDING=0 CELLSPACING=0 STYLE="BACKGROUND:url(i/m_l.gif) no-repeat 0px 0px;">
								<TR ALIGN=RIGHT VALIGN=BOTTOM><TD>
<script type="text/javascript">
AC_FL_RunContent( 
'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','width','227','height','304','src','9ai5business-nav.swf','quality','high','pluginspage','http://www.macromedia.com/go/getflashplayer','movie','9ai5business-nav','menu','false'); //end AC code
</script>
<noscript>								  
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="227" height="304">
                                    <param name="movie" value="9ai5business-nav.swf">
                                    <param name="quality" value="high"><param name="LOOP" value="false">
                                    <embed src="9ai5business-nav.swf" width="227" height="304" loop="false" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash"></embed>
							    </object></noscript>
							</TD>
								</TR>
								</TABLE></TD>
							<TD>
								<TABLE BORDER=0 WIDTH=10 HEIGHT=304 CELLPADDING=0 CELLSPACING=0>
								<TR HEIGHT=152 BGCOLOR=#B30000><TD></TD></TR>
								<TR><TD></TD></TR>
								<TR HEIGHT=76 BGCOLOR=#696969><TD></TD></TR>
								</TABLE></TD>
						</TR>
						</TABLE></TD>
				</TR>
				<TR HEIGHT=100%>
					<TD>
						<TABLE BORDER=0 WIDTH=542 HEIGHT=100% CELLPADDING=0 CELLSPACING=0>
						<TR VALIGN=TOP>
							<TD WIDTH=228 STYLE="BACKGROUND:url(i/c_rt.gif) no-repeat 100% 0px; BACKGROUND-COLOR:#B30000; PADDING:10px">
								<P><IMG SRC="i/h_services.gif" WIDTH="165" HEIGHT="19" BORDER="0"></P>
								<DIV id=news>
									<UL>
										<LI><A HREF="">Project Management</A></LI>
										<LI><a href="http://helpdesk.transworldinteractive.net">Project Help Desk</a></LI>
										<LI><A HREF="">Project Certification</A></LI>
										<LI><A HREF="">Project Hosting</A></LI>
										<LI><A HREF="">Project Applications</A></LI>
                                        <LI><A HREF="">Project Live Support</A></LI>
									</UL>
								<P>Our project management staff will help you in any of these categories listed. Transworld Interactive's outsourcing services provide IT solutions that support organizational needs, enabling you to minimize costs, improve efficiency and create a competitive advantage.</P>
								<DIV ALIGN=RIGHT><A HREF="" STYLE="COLOR:#000000; FONT-SIZE:11px;"><B>Read more<IMG SRC="i/arrow2.gif" WIDTH="15" HEIGHT="11" BORDER="0" ALIGN=absmiddle></B></A></DIV>
								</DIV>
							</TD>
							<TD WIDTH=314 BGCOLOR=#696969 STYLE="PADDING:10px;">
								<IMG SRC="i/h_news.gif" WIDTH="124" HEIGHT="20" BORDER="0" ALIGN=LEFT>
								<DIV ALIGN=RIGHT><A HREF="" STYLE="COLOR:#000000; FONT-SIZE:11px;"><B>ALL NEWS<IMG SRC="i/arrow2.gif" WIDTH="15" HEIGHT="11" BORDER="0" ALIGN=absmiddle></B></A></DIV><BR>
								<DIV id=news>
									<IMG SRC="i/pic2.jpg" WIDTH="95" HEIGHT="75" BORDER="0" ALIGN=LEFT STYLE="MARGIN-RIGHT:10px;"><DIV id=date>25.10.2006</DIV>
									<P><B>Project Manager is now ready to be licensed to our clients on our servers, when it is fully tested and touched up then it will be released to clients on other servers.</B>.</P>
								</DIV>
								<P ALIGN=RIGHT><A HREF="" STYLE="COLOR:#000000; FONT-SIZE:11px;"><B>Read more<IMG SRC="i/arrow2.gif" WIDTH="15" HEIGHT="11" BORDER="0" ALIGN=absmiddle></B></A></P>
								</DIV>
								<P><IMG SRC="i/emp.gif" WIDTH="100%" HEIGHT="3" BORDER="0" STYLE="BACKGROUND:url(i/hr.gif);"></P>
								<DIV id=news>
									<IMG SRC="i/pic3.jpg" WIDTH="95" HEIGHT="75" BORDER="0" ALIGN=LEFT STYLE="MARGIN-RIGHT:10px;"><DIV id=date>05.06.2011</DIV>
									<table border="0" cellpadding="2" cellspacing="2" align="center" width="185">
				<tr bgcolor="<%=gsColorHighlight%>"><td align="center" class="homeheader" colspan="2"><%=dictLanguage("Intranet_Login")%></td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>
				<form action="logoncheck.asp" method="POST" id="form1" name="form1">
				<input type="hidden" name="hname" value="hidden">
				<input type="hidden" name="ccokie" value="false">
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("ID")%>: </td>
					<td><input type="text" name="userid" class="formstyleShort"></td>
				</tr>
				<tr>
					<td valign="top" class="bolddark"><%=dictLanguage("Password")%>: </td>
					<td><input type="password" name="password" class="formstyleShort"></td>
				</tr>
				<tr><td align="center" colspan="2"><input type="submit" value="logon" class="formButton"></td></tr>
				<tr><td colspan="2">&nbsp;</td></tr>
				<tr><td align="center" colspan="2"><a href="forgotpassword.asp"><%=dictLanguage("ForgetPassword")%></a></td></tr>
				</form>
				<tr>
					<td align="center" colspan="2" class="alert">
						<p>
						<%=session("msg")%>
						<%session("msg") = ""%>
						<%If session("msg") = "" then
							session.abandon
						  End If %>
						</p>
					</td>
				</tr>
			</table>
								</DIV>
								<P ALIGN=RIGHT><A HREF="" STYLE="COLOR:#000000; FONT-SIZE:11px;"><B>Read more<IMG SRC="i/arrow2.gif" WIDTH="15" HEIGHT="11" BORDER="0" ALIGN=absmiddle></B></A></P>
								</DIV>
							</TD>
						</TR>
						</TABLE></TD>
				</TR>
				</TABLE></TD>
		</TR>
		</TABLE></TD>
</TR>
<TR HEIGHT=76>
	<TD>
		<TABLE BORDER=0 WIDTH=780 HEIGHT=76 CELLPADDING=0 CELLSPACING=0>
		<TR>
			<TD WIDTH=466 id=foot><A HREF="">Project</A><IMG SRC="i/f_slash.gif" WIDTH="29" HEIGHT="7" BORDER="0"><A HREF="">Service</A><IMG SRC="i/f_slash.gif" WIDTH="29" HEIGHT="7" BORDER="0"><A HREF="">Solutions</A><IMG SRC="i/f_slash.gif" WIDTH="29" HEIGHT="7" BORDER="0"><A HREF="">Support</A><IMG SRC="i/f_slash.gif" WIDTH="29" HEIGHT="7" BORDER="0"><A HREF="contacts.html">Contacts</A></TD>
			<TD WIDTH=314 id=foot2>Copyright © Transworld Interactive MO, 2011<BR>Read <A HREF="">Privacy / Policy</A>  and  <A HREF="">Teams of services</A></TD>
		</TR>
		</TABLE></TD>
</TR>
</TABLE>
</BODY>
</HTML>