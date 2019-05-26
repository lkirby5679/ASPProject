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
<script language="JavaScript">
<!-- Hide

<% if gsDHTMLNavigation then %>
var navTimer = 5;

function checkNavTimer() {
	//alert(navTimer);
	if (navTimer <= 0) {killAllSubs();}
	else {navTimer = navTimer - 1;}
	setTimeout("checkNavTimer()", 1000);
}

function rollOnNav(dis) {
	if (document.all) {
		document.all.item('nav',dis).style.color='<%=gsColorAltHighlight%>';
	}
	else if (document.getElementById) {
		//document.getElementById('nav',dis).style.color='<%=gsColorAltHighlight%>';
	}
}

function rollOffNav(dis) {
	if (document.all) {
		document.all.item('nav',dis).style.color='<%=gsColorBackground%>';
	}
	else if (document.getElementById) {
		//document.getElementById('nav',dis).style.color='<%=gsColorBackground%>';
	}	
}

function rollOnSubNav(dis) {
	if (document.all) {
		document.all.item('subnav',dis).style.color='<%=gsColorAltHighlight%>';
	}
	else if (document.getElementById) {
		//document.getElementById('subnav',dis).style.color='<%=gsColorAltHighlight%>';
	}
}

function rollOffSubNav(dis) {
	if (document.all) {
		document.all.item('subnav',dis).style.color='<%=gsColorBackground%>';
	}
	else if (document.getElementById) {
		//document.getElementById('subnav',dis).style.color='<%=gsColorBackground%>';
	}
}

function killAllSubs() {
	if (document.all) {
<% if gsTimecards then %>
		document.all.item('navSubTimecards').style.display='none';
<% end if 
   if gsTasks then %>
		document.all.item('navSubTasks').style.display='none';
<% end if 
   if gsProjects then %>
		document.all.item('navSubProjects').style.display='none';
<% end if
   if gsClients then %>
		document.all.item('navSubClients').style.display='none';
<% end if
   if gsReports then %>
		document.all.item('navSubReports').style.display='none';
<% end if
   if gsOther then %>
		document.all.item('navSubOther').style.display='none';
<% end if %>
	}
	else if (document.getElementById) {
<% if gsTimecards then %>	
		document.getElementById('navSubTimecards').style.visibility='hidden';
<% end if
   if gsTasks then %>
		document.getElementById('navSubTasks').style.visibility='hidden';
<% end if
   if gsProjects then %>
		document.getElementById('navSubProjects').style.visibility='hidden';
<% end if 
   if gsClients then %>
		document.getElementById('navSubClients').style.visibility='hidden';
<% end if
   if gsReports then %>
		document.getElementById('navSubReports').style.visibility='hidden';
<% end if
   if gsOther then %>
		document.getElementById('navSubOther').style.visibility='hidden';	
<% end if %>
	}
	else if (document.layers) {
		//document.layers['navSubTimecards'].visibility='hidden';
		//document.layers['navSubTasks'].visibility='hidden';
		//document.layers['navSubProjects'].visibility='hidden';
		//document.layers['navSubClients'].visibility='hidden';
		//document.layers['navSubReports'].visibility='hidden';
		//document.layers['navSubOther'].visibility='hidden';	
	}
	navTimer = 5;
}

function showNavSub(dis,e) {
	strNavSub = "navSub" + dis;
    if (e != '') {
        if (document.all) {
            x = e.x - 20;
            y = e.y + 5;
	    }
        else if (document.getElementById) {
            x = e.pageX - 20;
            y = e.pageY + 5;
        }     
        else if (document.layers) {
            //x = e.pageX - 20;
            //y = e.pageY + 5;        
        } 
    }

	if (document.all) {
		document.all.item(strNavSub).style.posLeft = x;
		document.all.item(strNavSub).style.posTop = y;
		document.all.item(strNavSub).style.display = 'block';
	}
	else if (document.getElementById) {
	    document.getElementById(strNavSub).style.left = x;
	    document.getElementById(strNavSub).style.top = y;
	    document.getElementById(strNavSub).style.display = 'block';
		document.getElementById(strNavSub).style.visibility = 'visible';	
	}
	else if (document.layers) {
	    //document.layers[strNavSub].left = x;
	    //document.layers[strNavSub].top = y;
		//document.layers[strNavSub].display = 'block';
		//document.layers[strNavSub].visibility = 'visible';
	}
}
<% end if %>

/**********************************************************
* Name:		doNothing()
* Description:	empty placeholder
**********************************************************/
function doNothing()
{
	;
}

/**********************************************************
* Name:		txtStartDate_onKeyPress()
* Description:	IE event that only allows valid chars to 
*				be typed in date textboxes
**********************************************************/
function txtDate_onKeypress()
{
	var string="1234567890/";
	if (string.indexOf(unescape('%' + window.event.keyCode.toString(16))) != -1) 
		window.event.returnValue = true;
	else
		window.event.returnValue = false;
}

/**********************************************************
* Name:		txtTime_onKeyPress()
* Description:	IE event that only allows valid chars to 
*				be typed in time textboxes
**********************************************************/
function txtTime_onKeypress()
{
	var string="1234567890:AMP";
	if (string.indexOf(unescape('%' + window.event.keyCode.toString(16))) != -1) 
		window.event.returnValue = true;
	else
		window.event.returnValue = false;
}

/**********************************************************
 * Name:		StartDate_Change()
 * Description:	Callback for popup calendar, changes the 
 *				start date
 **********************************************************/
function Date_Change(sNewValue,sField)
{
	document.strForm[sField].value = sNewValue;
}

function openCalendar(sDate,sCallback,sCallbackField,nTop,nLeft)
{
	var oWindow;
	oWindow = window.open("<%=gsSiteRoot%>popupcalendar.asp?date=" + sDate + "&callback=" + sCallback + "" + "&callbackfield=" + sCallbackField + "","calendar","top=" + nTop + ",left=" + nLeft + ",height=180,width=160,status=no,resizable=no,toolbar=no,menubar=no,scrollbars=no,location=no");
	oWindow.focus();
	return true;
}

function toggleWindow(sWindow)
{
	//when = newDate();
	//when.setDay(when.getDay()+1);
	//when = when.toGMTString();
	if (document.all) {
		if (document.all(sWindow).style.display == 'none') {
			document.all(sWindow).style.display = 'block';
			sWin = sWindow + "_option";
			document.images(sWin).src = '<%=gsSiteRoot%>images/x.gif';
			document.images(sWin).alt = '<%=dictLanguage("Collapse_This_Window")%>';
			document.cookie = sWindow + "=TRUE";
		}
		else {
			document.all(sWindow).style.display = 'none';
			sWin = sWindow + "_option";
			document.images(sWin).src = '<%=gsSiteRoot%>images/maximize.gif';
			document.images(sWin).alt = '<%=dictLanguage("Open_This_Window")%>';
			document.cookie = sWindow + "=FALSE";					
		}
	}
	else if (document.getElementById) {
		if (document.getElementById(sWindow).style.display == 'none') {
			document.getElementById(sWindow).style.display = 'block';
			document.getElementById(sWindow).style.visibility = 'visible';
			sWin = sWindow + "_option";
			document.images[sWin].src = '<%=gsSiteRoot%>images/x.gif';
			document.images[sWin].alt = '<%=dictLanguage("Collapse_This_Window")%>';
			document.cookie = sWindow + "=TRUE";
		}
		else {
			document.getElementById(sWindow).style.display = 'none';
			document.getElementById(sWindow).style.visibility = 'hidden';
			sWin = sWindow + "_option";
			document.images[sWin].src = '<%=gsSiteRoot%>images/maximize.gif';
			document.images[sWin].alt = '<%=dictLanguage("Open_This_Window")%>';
			document.cookie = sWindow + "=FALSE";					
		}	
	}
}

function popup(url) {
	newWindow = window.open(url,'daughter','menubar=yes,toolbar=no,location=no,scrollbars=yes,resizable=no,width=433,height=500,screenX=100,screenY=100,top=100,left=100');
	if (newWindow.opener == null) newWindow.opener = self;
	newWindow.focus();
}

function popupwide(url) {
	newWindow = window.open(url,'daughter','menubar=yes,toolbar=no,location=no,scrollbars=yes,resizable=no,width=600,height=500,screenX=100,screenY=100,top=100,left=100');
	if (newWindow.opener == null) newWindow.opener = self;
	newWindow.focus();
}

function popup2(url) {
	newWindow = window.open(url,'daughter2','menubar=yes,toolbar=no,location=no,scrollbars=yes,resizable=no,width=433,height=500,screenX=200,screenY=200,top=200,left=200');
	if (newWindow.opener == null) newWindow.opener = self;
	newWindow.focus();
}

function popupcalendar(url) {
	newWindow = window.open(url,'daughter','menubar=yes,toolbar=no,location=no,scrollbars=yes,resizable=no,width=700,height=500,screenX=100,screenY=100,top=100,left=100');
	if (newWindow.opener == null) newWindow.opener = self;
	newWindow.focus();
}

//stop hiding-->
</script>
