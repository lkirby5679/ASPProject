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

<style type="text/css">
<!--
/* this section contains the general styles applied to different tags */
BODY {
	background-color:#FFFFFF; 
	margin-top: 0; 
	margin-right: 0; 
	margin-bottom: 0; 
	margin-left: 0;
	}

A {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 9pt; 
	font-weight: normal;
	color:<%=gsColorBlack%>;
	text-decoration: underline;
	}

A.home {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 9pt; 
	font-weight: normal;
	color:<%=gsColorWhite%>;
	text-decoration: underline;
	}

P, UL, L, TD {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 9pt; 
	font-weight: normal;
	color:<%=gsColorBlack%>;
	text-decoration: none;
	}	

.alert {
	color:#993333;
}

.boldwhite {
	color:<%=gsColorWhite%>;
	font-weight:bold;	
}

.boldwhitecaps {
	font-family: Arial Narrow, Arial, Helvetica, Sans-Serif; 
	color:<%=gsColorWhite%>;
	font-weight:bold;	
}

.bolddark {
	color:<%=gsColorBackground%>;
	font-weight:bold;	
}

.boldalthighlight {
	font-family: Arial Narrow, Arial, Helvetica, Sans-Serif;
	color: <%=gsColorAltHighlight%>;
	font-weight: bold;
}

.small {
	font-size: 8pt;
}

.smallhome {
	font-size: 8pt;
	color: <%=gsColorWhite%>;
}

.more {
	font-size: 8pt;
	color: <%=gsColorAltHighlight%>;
}

.smallalthighlight {
	font-family: Arial Narrow, Arial, Helvetica, Sans-Serif;
	font-size: 8pt;
	color: <%=gsColorAltHighlight%>;
	font-weight: normal;
}

.homeheader {
	font-family: Arial Narrow, Arial, Helvetica, Sans-Serif; 
	font-size: 10pt;
	color: <%=gsColorBackground%>;
	font-weight: bold;
	text-decoration: none;		
}

.header {
	font-family: Arial Narrow, Arial, Helvetica, Sans-Serif; 
	font-size: 12pt;
	color: <%=gsColorBackground%>;
	font-weight: bold;
	text-decoration: none;		
}

.columnheader {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 10pt;
	color: <%=gsColorBackground%>;
	font-weight: bold;
	text-decoration: none;		
}

.nav {
	font-family: Arial Narrow, Arial, Helvetica, Sans-Serif; 
	font-size: 10pt;
	color: <%=gsColorBackground%>;
	font-weight: bold;
	text-decoration: none;		
}

.homealtheader {
	font-family: Arial Narrow, Arial, Helvetica, Sans-Serif; 
	font-size: 10pt;
	color: <%=gsColorHighlight%>;
	font-weight: bold;
	text-decoration: none;		
}

.hnfirstline {
	font-family: Verdana, Arial, Helvetica, Sans-Serif; 
	font-size: 8pt;
	color: <%=gsColorWhite%>;
	font-weight: normal;
	text-decoration: none;	
}

.hnfirstline A {
	font-family: Verdana, Arial, Helvetica, Sans-Serif; 
	font-size: 8pt;
	color: <%=gsColorWhite%>;
	font-weight: normal;
	text-decoration: underline;	
}

.hnheadline A {
	font-family: Verdana, Arial, Helvetica, Sans-Serif; 
	font-size: 8pt;
	color: <%=gsColorWhite%>;
	font-weight: normal;
	text-decoration: underline;	
}

.hnsource {
	font-family: Verdana, Arial, Helvetica, Sans-Serif; 
	font-size: 8pt;
	color: <%=gsColorWhite%>;
	font-weight: normal;
	text-decoration: none;	
}

.homecontent {
	font-family: Verdana, Arial, Helvetica, Sans-Serif; 
	font-size: 9pt;
	color: <%=gsColorWhite%>;
	font-weight: normal;
	text-decoration: none;		
}

.content {
	font-family: Verdana, Arial, Helvetica, Sans-Serif; 
	font-size: 10pt;
	color: <%=gsColorBlack%>;
	font-weight: normal;
	text-decoration: none;		
}

.formbutton {
	font-family: Arial, Helvetica, Sans-serif; 
	font-size:10pt; 
	font-weight:bold; 
	color:<%=gsColorBackground%>; 
	background-color:<%=gsColorHighlight%>;
	}

@media print {
	.formbutton {display: none;}
	.noprint {display: none;}
}

.formstyleXL {
	font-family: Arial, Helvetica, Sans-serif; 
	font-size:8pt; 
	width: 350px;
	font-weight:normal; 
	color:<%=gsColorBlack%>; 
	background-color:<%=gsColorWhite%>;
	}
		
.formstyleLong {
	font-family: Arial, Helvetica, Sans-serif; 
	font-size:8pt; 
	width: 250px;
	font-weight:normal; 
	color:<%=gsColorBlack%>; 
	background-color:<%=gsColorWhite%>;
	}
	
.formstyleMed {
	font-family: Arial, Helvetica, Sans-serif; 
	font-size:8pt; 
	width: 175px;
	font-weight:normal; 
	color:<%=gsColorBlack%>; 
	background-color:<%=gsColorWhite%>;
	}	
	
.formstyleShort {
	font-family: Arial, Helvetica, Sans-serif; 
	font-size:8pt; 
	width: 90px;
	font-weight:normal; 
	color:<%=gsColorBlack%>; 
	background-color:<%=gsColorWhite%>;
	}
	
.formstyleTiny {
	font-family: Arial, Helvetica, Sans-serif; 
	font-size:8pt; 
	width: 40px;
	font-weight:normal; 
	color:<%=gsColorBlack%>; 
	background-color:<%=gsColorWhite%>;
	}

.calendardates {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 8pt; 
	font-weight: normal;
	color:<%=gsColorBackground%>;
	text-decoration: none;
	BORDER: 1px solid #FFFFFF;	
	}	

.calendardateslink {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 8pt; 
	font-weight: normal;
	color:#EEEEEE;
	text-decoration: underlined;	
	}	

.calendardatestoday {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 8pt; 
	font-weight: bold;
	color:<%=gsColorBackground%>;
	text-decoration: none;
	BORDER: 1px solid #FFFFFF;	
	}
	
.calendardatestodaylink {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 8pt; 
	font-weight: bold;
	color:#EEEEEE;
	text-decoration: underlined;
	}	

.calendarevent {
	font-family: Arial, Helvetica, Sans-Serif; 
	font-size: 8pt; 
	font-weight: normal;
	color:<%=gsColorWhite%>;
	text-decoration: none;	
	}	

-->
</style>
