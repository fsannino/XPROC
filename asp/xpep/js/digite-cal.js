var weekend = [0,6];
//var weekendColor = "#e0e0e0";
var weekendColor = "#639ACE";
//var fontface = "Verdana";
var fontface = "#E7EBEF";
var fontsize = 2;
var gNow = new Date();
var ggWinCal;

var gCanSelPastDate;

isNav = (navigator.appName.indexOf("Netscape") != -1) ? true : false;
isIE = (navigator.appName.indexOf("Microsoft") != -1) ? true : false;

Calendar.Months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
"Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

Calendar.ShortMonths = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
"Jul", "Ago", "Set", "Out", "Nov", "Dez"];

// Non-Leap year Month days..
Calendar.DOMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
// Leap year Month days..
Calendar.lDOMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

function Calendar(p_item, p_WinCal, p_month, p_year, p_format) {
	if ((p_month == null) && (p_year == null))	return;

	if (p_WinCal == null)
		this.gWinCal = ggWinCal;
	else
		this.gWinCal = p_WinCal;
	
	if (p_month == null) {
		this.gMonthName = null;
		this.gMonth = null;
		this.gYearly = true;
	} else {
		this.gMonthName = Calendar.get_month(p_month);
		this.gMonth = new Number(p_month);
		this.gYearly = false;
	}

	this.gYear = p_year;
	this.gFormat = p_format;
	this.gBGColor = "white";
	this.gFGColor = "black";
	this.gTextColor = "black";
	this.gHeaderColor = "black";
	this.gReturnItem = p_item;
}

Calendar.get_month = Calendar_get_month;
Calendar.get_daysofmonth = Calendar_get_daysofmonth;
Calendar.calc_month_year = Calendar_calc_month_year;
Calendar.print = Calendar_print;

function Calendar_get_month(monthNo) {
	return Calendar.Months[monthNo];
}

function Calendar_get_daysofmonth(monthNo, p_year) {
	/* 
	Check for leap year ..
	1.Years evenly divisible by four are normally leap years, except for... 
	2.Years also evenly divisible by 100 are not leap years, except for... 
	3.Years also evenly divisible by 400 are leap years. 
	*/
	if ((p_year % 4) == 0) {
		if ((p_year % 100) == 0 && (p_year % 400) != 0)
			return Calendar.DOMonth[monthNo];
	
		return Calendar.lDOMonth[monthNo];
	} else
		return Calendar.DOMonth[monthNo];
}

function Calendar_calc_month_year(p_Month, p_Year, incr) {
	/* 
	Will return an 1-D array with 1st element being the calculated month 
	and second being the calculated year 
	after applying the month increment/decrement as specified by 'incr' parameter.
	'incr' will normally have 1/-1 to navigate thru the months.
	*/
	var ret_arr = new Array();
	
	if (incr == -1) {
		// B A C K W A R D
		if (p_Month == 0) {
			ret_arr[0] = 11;
			ret_arr[1] = parseInt(p_Year) - 1;
		}
		else {
			ret_arr[0] = parseInt(p_Month) - 1;
			ret_arr[1] = parseInt(p_Year);
		}
	} else if (incr == 1) {
		// F O R W A R D
		if (p_Month == 11) {
			ret_arr[0] = 0;
			ret_arr[1] = parseInt(p_Year) + 1;
		}
		else {
			ret_arr[0] = parseInt(p_Month) + 1;
			ret_arr[1] = parseInt(p_Year);
		}
	}
	
	return ret_arr;
}

function Calendar_print() {
	ggWinCal.print();
}

// This is for compatibility with Navigator 3, we have to create and discard one object before the prototype object exists.
new Calendar();

Calendar.prototype.getMonthlyCalendarCode = function() {
	var vCode = "";
	var vHeader_Code = "";
	var vData_Code = "";
	
	// Begin Table Drawing code here..
	//vCode = vCode + "<TABLE BORDER=1 BGCOLOR=\"" + this.gBGColor + "\">";

  vCode = vCode + "<tr align='center'> ";
  vCode = vCode + " <td> ";
  vCode = vCode + "  <table border='1' cellpadding=2 cellspacing='0'>";
  vCode = vCode + "   <tbody> ";

	vHeader_Code = this.cal_header();
	vData_Code = this.cal_data();
	vCode = vCode + vHeader_Code + vData_Code;

	vCode = vCode + "   </tbody>";
  vCode = vCode + "  </table";
  vCode = vCode + " </td";
  vCode = vCode + "</tr>";

	return vCode;
}

Calendar.prototype.displayMonthHeader = function()
{
  var vCode = "";
  vCode = "<tr align=center class=f>";
  for(var j=0;j<6;j++)
  {
    vCode = vCode + "<td width=16%";
    if(j==this.gMonth)
	// MUDA FONTE DOS MESES
      //vCode = vCode + " bgcolor='#CC9966'";
	  vCode = vCode + " bgcolor='#639ACE'";
      vCode = vCode + "><b><a href=\"" +
             "javascript:window.opener.Build(" +
		         "'" + this.gReturnItem + "', '" + j + "', '" + this.gYear + "', '" + this.gFormat + "'" + ");" +
		   "\"><font color='#FFFFFF' size='1'>" + Calendar.ShortMonths[j] + "</font></a></b></td>";
           //"\"><font color='#639ACE' size='1'>" + Calendar.ShortMonths[j] + "</font></a></b></td>";

}
  vCode = vCode + "</tr>"

    vCode = vCode + "<tr align=center class=f>";
  for(var j=6;j<12;j++)
  {
    vCode = vCode + "<td width=16%";
    if(j==this.gMonth)
      vCode = vCode + " bgcolor='#CC9966'";
      vCode = vCode + "><b><a href=\"" +
             "javascript:window.opener.Build(" +
		         "'" + this.gReturnItem + "', '" + j + "', '" + this.gYear + "', '" + this.gFormat + "'" + ");" +
           "\"><font color='#FFFFFF' size='1'>" + Calendar.ShortMonths[j] + "</font></a></b></td>";
  }
  vCode = vCode + "</tr>"

    return vCode;
}

Calendar.prototype.show = function() {
  var hdrCode = "";
	var vCode = "";

	this.gWinCal.document.open();

	// Setup the page...
	this.wwrite("<html>");
	this.wwrite("<head><title>Calendar</title>");
  this.wwrite("<link rel='stylesheet' href='common/calender.css' type='text/css'>");
	this.wwrite("</head>");

	this.wwrite("<body " +
		"link=\"" + this.gLinkColor + "\" " +
		"vlink=\"" + this.gVLinkColor + "\" " +
		"alink=\"" + this.gALinkColor + "\" " +
		"text=\"" + this.gTextColor + "\""
    +" leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>");

	// Show navigation buttons
	var prevMMYYYY = Calendar.calc_month_year(this.gMonth, this.gYear, -1);
	var prevMM = prevMMYYYY[0];
	var prevYYYY = prevMMYYYY[1];

	var nextMMYYYY = Calendar.calc_month_year(this.gMonth, this.gYear, 1);
	var nextMM = nextMMYYYY[0];
	var nextYYYY = nextMMYYYY[1];

  hdrCode = this.displayMonthHeader();
  this.wwrite("<table width='100%' align='center'>");
  this.wwrite(" <tr align='center'> ");
  this.wwrite("  <td> ");
  
  this.wwrite("   <table border='0' cellspacing='0' cellpadding='0' bgcolor='#E7EBEF'>");
  //this.wwrite("   <table border='0' cellspacing='0' cellpadding='0' bgcolor='#639ACE'>");
  this.wwrite("    <tr> ");
  this.wwrite("     <td> ");
  this.wwrite("      <table border='0' cellspacing='1' cellpadding='0'>");
  //this.wwrite("       <tr align='center' bgcolor='#5F3F1F'> ");
  this.wwrite("       <tr align='center' bgcolor='#639ACE'> ");    
  this.wwrite("        <td> ");
  this.wwrite("         <table width='100%' cellspacing='0'>");
  this.wwrite("          <tbody> ");

  this.wwrite(hdrCode);

  this.wwrite("          </tbody> ");
  this.wwrite("         </table>");
  this.wwrite("        </td>");
  this.wwrite("       </tr>");
  this.wwrite("       <tr align='center' bgcolor='#E7EBEF'> ");
  this.wwrite("        <td> ");
  this.wwrite("         <table align=center>");
  this.wwrite("          <tbody>");

  this.wwrite("           <tr> ");
  this.wwrite("            <td width='29' valign=middle align=left><a href=\""+
           "javascript:window.opener.Build(" +
		       "'" + this.gReturnItem + "', '" + this.gMonth + "', '" + (this.gYear -1) + "', '" + this.gFormat + "'" + ");" +
           "\"><font color='#000000'><b><font face='Arial, Helvetica, sans-serif' size='2'>&lt;&lt;</font></b></font></a></td>");
  this.wwrite("            <td align=center  width='35' valign=top><b><font color='#000000'><font size='2' face='Arial, Helvetica, sans-serif'>"+this.gYear+"</font></font></b></td>");
  this.wwrite("            <td align=right width='29' valign=top><a href=\"" +
           "javascript:window.opener.Build(" +
		       "'" + this.gReturnItem + "', '" + this.gMonth + "', '" + (this.gYear -1+2) + "', '" + this.gFormat + "'" + ");" +
          "\"><font color='#000000'><b><font face='Arial, Helvetica, sans-serif' size='2'>&gt;&gt;</font></b></font></a></td>");
  this.wwrite("           </tr>");

  this.wwrite("          </tbody> ");
  this.wwrite("         </table>");
  this.wwrite("        </td>");
  this.wwrite("       </tr>");

	//this.wwrite("<TABLE WIDTH='100%' BORDER=1 CELLSPACING=0 CELLPADDING=0 BGCOLOR='#e0e0e0'><TR><TD ALIGN=center>");
	//this.wwrite("[<A HREF=\"" +
		//"javascript:window.opener.Build(" +
		//"'" + this.gReturnItem + "', '" + this.gMonth + "', '" + (parseInt(this.gYear)-1) + "', '" + this.gFormat + "'" +
		//");" +
		//"\"><<<\/A>]</TD><TD ALIGN=center>");
	//this.wwrite("[<A HREF=\"" +
		//"javascript:window.opener.Build(" + 
		//"'" + this.gReturnItem + "', '" + prevMM + "', '" + prevYYYY + "', '" + this.gFormat + "'" +
		//");" +
		//"\"><<\/A>]</TD><TD ALIGN=center>");
	//this.wwrite("[<A HREF=\"javascript:window.print();\">Print</A>]</TD><TD ALIGN=center>");
	//this.wwrite("[<A HREF=\"" +
		//"javascript:window.opener.Build(" + 
		//"'" + this.gReturnItem + "', '" + nextMM + "', '" + nextYYYY + "', '" + this.gFormat + "'" +
		//");" +
		//"\">><\/A>]</TD><TD ALIGN=center>");
	//this.wwrite("[<A HREF=\"" +
		//"javascript:window.opener.Build(" + 
		//"'" + this.gReturnItem + "', '" + this.gMonth + "', '" + (parseInt(this.gYear)+1) + "', '" + this.gFormat + "'" +
		//");" +
		//"\">>><\/A>]</TD></TR></TABLE><BR>");

	// Get the complete calendar code for the month..
	vCode = this.getMonthlyCalendarCode();
	this.wwrite(vCode);

  this.wwrite("      </table>");
  this.wwrite("     </td>");
  this.wwrite("    </tr>");
  this.wwrite("   </table>");
  this.wwrite("  </td>");
  this.wwrite(" </tr>");
  this.wwrite("</table>");

	this.wwrite("</body></html>");
	this.gWinCal.document.close();
}


Calendar.prototype.wwrite = function(wtext) {
	this.gWinCal.document.writeln(wtext);
}

Calendar.prototype.wwriteA = function(wtext) {
	this.gWinCal.document.write(wtext);
}

Calendar.prototype.cal_header = function() {
	var vCode = "";

  vCode = vCode + "<tr align=right class=f> ";
  vCode = vCode + " <td width=19 bgcolor='#CCCCCC'><b>Do</b></td>";
  vCode = vCode + " <td width=19 bgcolor='#CCCCCC'><b>Se</b></td>";
  vCode = vCode + " <td width=19 bgcolor='#CCCCCC'><b>Te</b></td>";
  vCode = vCode + " <td width=19 bgcolor='#CCCCCC'><b>Qu</b></td>";
  vCode = vCode + " <td width=19 bgcolor='#CCCCCC'><b>Qu</b></td>";
  vCode = vCode + " <td width=19 bgcolor='#CCCCCC'><b>Se</b></td>";
  vCode = vCode + " <td width=19 bgcolor='#CCCCCC'><b>Sa</b></td>";
  vCode = vCode + "</tr>";
	
	return vCode;
}

Calendar.prototype.cal_data = function() {
	var vDate = new Date();
	vDate.setDate(1);
	vDate.setMonth(this.gMonth);
	vDate.setFullYear(this.gYear);

	var vFirstDay=vDate.getDay();
	var vDay=1;
	var vLastDay=Calendar.get_daysofmonth(this.gMonth, this.gYear);
	var vOnLastDay=0;
	var vCode = "";
  var defBg = "#FFFFFF";
  var selBg = "#EFEFEF";
	var vNowDay = gNow.getDate();
	var vNowMonth = gNow.getMonth();
	var vNowYear = gNow.getFullYear();
  var bgColor=defBg;

	/*
	Get day for the 1st of the requested month/year..
	Place as many blank cells before the 1st day of the month as necessary. 
	*/

	vCode = vCode + "<tr align=right class=f>";

	for (i=0; i<vFirstDay; i++) {
		//vCode = vCode + "<TD WIDTH='14%'" + this.write_weekend_string(i) + "><FONT SIZE='2' FACE='" + fontface + "'> &nbsp; </FONT></TD>";
    vCode = vCode + "<td bgcolor='#FFFFFF'>&nbsp;</td>"
	}

	// Write rest of the 1st week
	for (j=vFirstDay; j<7; j++) {
		//vCode = vCode + "<TD WIDTH='14%'" + this.write_weekend_string(j) + "><FONT SIZE='2' FACE='" + fontface + "'>" + 
			   //this.formatDate(vDay) + 
				//"</FONT></TD>";
    bgColor=defBg;
	  if (vDay == vNowDay && this.gMonth == vNowMonth && this.gYear == vNowYear)
    {
      bgColor=selBg;
    }


    vCode = vCode + "<td bgcolor='"+bgColor+"'" + this.write_weekend_string(j) + ">" +
			   this.formatDate(vDay) + "</td>";
		vDay=vDay + 1;
	}
	vCode = vCode + "</TR>";

	// Write the rest of the weeks
	for (k=2; k<7; k++) {
	  vCode = vCode + "<tr align=right class=f>";

		for (j=0; j<7; j++) {
			//vCode = vCode + "<TD WIDTH='14%'" + this.write_weekend_string(j) + "><FONT SIZE='2' FACE='" + fontface + "'>" + 
				//this.formatDate(vDay) + 
				//"</FONT></TD>";
    bgColor=defBg;
	  if (vDay == vNowDay && this.gMonth == vNowMonth && this.gYear == vNowYear)
    {
      bgColor=selBg;
    }
    vCode = vCode + "<td bgcolor='"+bgColor+"'" + this.write_weekend_string(j) + ">" +
			   this.formatDate(vDay) + "</td>";
			vDay=vDay + 1;

			if (vDay > vLastDay) {
				vOnLastDay = 1;
				break;
			}
		}

		if (j == 6)
			vCode = vCode + "</TR>";
		if (vOnLastDay == 1)
			break;
	}
	
	// Fill up the rest of last week with proper blanks, so that we get proper square blocks
	for (m=1; m<(7-j); m++) {
		if (this.gYearly)
			vCode = vCode + "<td bgcolor='#FFFFFF' " + this.write_weekend_string(j+m) + 
			"> </td>";
		else
			vCode = vCode + "<td bgcolor='#FFFFFF' " + this.write_weekend_string(j+m) + 
			"> " + "&nbsp;" + "</td>";
	}
	
	return vCode;
}



Calendar.prototype.write_weekend_string = function(vday) {
	var i;

	// Return special formatting for the weekend day.
	for (i=0; i<weekend.length; i++) {
		if (vday == weekend[i])
      return "";
			//return (" BGCOLOR=\"" + weekendColor + "\"");
	}
	
	return "";
}

Calendar.prototype.format_data = function(p_day) {
	var vData;
	var vMonth = 1 + this.gMonth;
	vMonth = (vMonth.toString().length < 2) ? "0" + vMonth : vMonth;

	var vMon = Calendar.get_month(this.gMonth).substr(0,3).toUpperCase();
	var vFMon = Calendar.get_month(this.gMonth).toUpperCase();
	var vY4 = new String(this.gYear);
	var vY2 = new String(this.gYear.substr(2,2));
	var vDD = (p_day.toString().length < 2) ? "0" + p_day : p_day;
  this.gFormat = this.gFormat.toUpperCase();
	switch (this.gFormat) {
		case "MM\/DD\/YYYY" :
			vData = vMonth + "\/" + vDD + "\/" + vY4;
			break;
		case "MM\/DD\/YY" :
			vData = vMonth + "\/" + vDD + "\/" + vY2;
			break;
		case "MM-DD-YYYY" :
			vData = vMonth + "-" + vDD + "-" + vY4;
			break;
		case "MM-DD-YY" :
			vData = vMonth + "-" + vDD + "-" + vY2;
			break;

		case "DD\/MMM\/YYYY" :
			vData = vDD + "\/" + vMon + "\/" + vY4;
			break;
		case "DD\/MMM\/YY" :
			vData = vDD + "\/" + vMon + "\/" + vY2;
			break;
		case "DD-MMM-YYYY" :
			vData = vDD + "-" + vMon + "-" + vY4;
			break;
		case "DD-MMM-YY" :
			vData = vDD + "-" + vMon + "-" + vY2;
			break;

		case "MMM-DD-YYYY" :
			vData = vMon + "-" + vDD + "-" + vY4;
			break;
		case "MMM-DD-YY" :
			vData = vMon + "-" + vDD + "-" + vY2;
			break;

		case "DD\/MONTH\/YYYY" :
			vData = vDD + "\/" + vFMon + "\/" + vY4;
			break;
		case "DD\/MONTH\/YY" :
			vData = vDD + "\/" + vFMon + "\/" + vY2;
			break;
		case "DD-MONTH-YYYY" :
			vData = vDD + "-" + vFMon + "-" + vY4;
			break;
		case "DD-MONTH-YY" :
			vData = vDD + "-" + vFMon + "-" + vY2;
			break;

		case "DD\/MM\/YYYY" :
			vData = vDD + "\/" + vMonth + "\/" + vY4;
			break;
		case "DD\/MM\/YY" :
			vData = vDD + "\/" + vMonth + "\/" + vY2;
			break;
		case "DD-MM-YYYY" :
			vData = vDD + "-" + vMonth + "-" + vY4;
			break;
		case "DD-MM-YY" :
			vData = vDD + "-" + vMonth + "-" + vY2;
			break;


		case "YYYY\/MM\/DD" :
			vData = vY4 + "\/" + vMonth + "\/" + vDD;
			break;
		case "YY\/MM\/DD" :
			vData = vY2 + "\/" + vMonth + "\/" + vDD;
			break;
		case "YYYY-MM-DD" :
			vData = vY4 + "-" + vMonth + "-" + vDD;
			break;
		case "YY-MM-DD" :
			vData = vY2 + "-" + vMonth + "-" + vDD;
			break;

		default :
			vData = vMonth + "\/" + vDD + "\/" + vY4;
	}
	return vData;
}

function Build(p_item, p_month, p_year, p_format) {
	var p_WinCal = ggWinCal;
	gCal = new Calendar(p_item, p_WinCal, p_month, p_year, p_format);

	// Customize your Calendar here..
	gCal.gBGColor="white";
	gCal.gVLinkColor="black";
  gCal.gLinkColor="#0000e1";
  gCal.gALinkColor="red";
	gCal.gTextColor="black";
	gCal.gHeaderColor="darkgreen";

	// Choose appropriate show function
  gCal.show();
}

function show_calendar() {
	/* 
		p_month : 0-11 for Jan-Dec; 12 for All Months.
		p_year	: 4-digit year
		p_format: Date format (mm/dd/yyyy, dd/mm/yy, ...)
		p_item	: Return Item.
	*/
	//alert("AQUI");
	gCanSelPastDate = arguments[0];

	p_item = arguments[1];

	if (arguments[2] == null)
		p_format = "mm/dd/yyyy";
	else
		p_format = arguments[2];
	if (isNaN(arguments[3]))
	{
		p_month = new String(gNow.getMonth());
		
	}
	else
		p_month = arguments[3];

	if (arguments[4] == "" || arguments[4] == null || arguments[4] == undefined)
		p_year = new String(gNow.getFullYear().toString());
	else
		p_year = arguments[4];

	vWinCal = window.open("", "Calendar", 
    "resizable=yes,width=200,height=210")
	vWinCal.opener = self;
	ggWinCal = vWinCal;

	Build(p_item, p_month, p_year, p_format);
}

Calendar.prototype.formatDate=function(a_Day){
	var w_Str ='';
	var vNowDay = gNow.getDate();
	var vNowMonth = gNow.getMonth();
	var vNowYear = gNow.getFullYear();
	var w_CurrentDate = new Date(vNowYear,vNowMonth,vNowDay);
	var w_Date= new Date(this.gYear,this.gMonth,a_Day);
	
	if(w_Date >= w_CurrentDate || gCanSelPastDate==true ){
	    var w_DestControl = "self.opener.document." + this.gReturnItem;
	    //Call onchangeHandler if exists.
	    var w_OnChangeHandler = "if(" + w_DestControl + ".onchange != null && " + w_DestControl + ".onchange != 'undefined' ){ "
				    + w_DestControl + ".onchange();}" ;

		w_Str = "<A HREF='#' " + "onClick=\"" + w_DestControl +
			    ".value='" + this.format_data(a_Day) + "';" + w_OnChangeHandler +
			    "window.close();\">" + 
			    this.format_day(a_Day) +  "</A>" ;

	}else{
		w_Str = this.format_day(a_Day);
	}

	return w_Str;
}


Calendar.prototype.format_day = function(vday) {
	var vNowDay = gNow.getDate();
	var vNowMonth = gNow.getMonth();
	var vNowYear = gNow.getFullYear();

	if (vday == vNowDay && this.gMonth == vNowMonth && this.gYear == vNowYear)
		return ("<B>" + vday + "</B>");
	else
		return (vday);
}

