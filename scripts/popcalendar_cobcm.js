	var	PC_fixedX = -1			// x position (-1 if to appear below control)
	var	PC_fixedY = -1			// y position (-1 if to appear below control)
	var PC_startAt = 1			// 0 - suPC_nDay ; 1 - moPC_nDay
	var PC_showWeekNumber = 1	// 0 - don't show; 1 - show
	var PC_showPC_today = 1		// 0 - don't show; 1 - show
	var PC_PC_imgDir = "imagenes/"			// directory for images ... e.g. var PC_PC_imgDir="/PC_img/"

	var PC_gotoString = "Ir a Mes Actual"
	var PC_PC_todayString = "Hoy es"
	var PC_weekString = "Sem"
	var PC_scrollLeftMessage = "Click to scroll to previous month. Hold mouse button to scroll automatically."
	var PC_scrollRightMessage = "Click to scroll to next month. Hold mouse button to scroll automatically."
	var PC_selectMonthMessage = "Click to select a month."
	var PC_PC_selectYearMessage = "Click to select a year."
	var PC_selectDateMessage = "Select [date] as date." // do not replace [date], it will be replaced by date.

	var	PC_crossobj, PC_crossMonthObj, PC_crossYearObj, PC_monthSelected, PC_yearSelected, PC_dateSelected, PC_monthSelected, PC_yearSelected, PC_dateSelected, PC_monthConstructed, PC_yearConstructed, PC_intervalID1, PC_intervalID2, PC_timeoutID1, PC_timeoutID2, PC_PC_ctlToPlaceValue, PC_PC_ctlNow, PC_dateFormat, PC_nStartingYear

	var	PC_bPageLoaded=false
	var	PC_ie=document.all
	var	PC_dom=document.getElementById

	var	PC_ns4=document.layers
	var	PC_today =	new	Date()
	var	PC_dateNow	 = PC_today.getDate()
	var	PC_monthNow = PC_today.getMonth()
	var	PC_yearNow	 = PC_today.getYear()
	var	PC_PC_imgsrc = new Array("drop1.gif","drop2.gif","left1.gif","left2.gif","right1.gif","right2.gif")
	var	PC_img	= new Array()

	var PC_bShow = false;

    /* hides <select> and <applet> objects (for PC_ie only) */
    function PC_hideElement( PC_elmID, PC_overDiv )
    {
      if( PC_ie )
      {
        for( i = 0; i < document.all.tags( PC_elmID ).length; i++ )
        {
          obj = document.all.tags( PC_elmID )[i];
          if( !obj || !obj.offsetParent )
          {
            continue;
          }
      
          // Find the element's offsetTop and offsetLeft relative to the BODY tag.
          objLeft   = obj.offsetLeft;
          objTop    = obj.offsetTop;
          objParent = obj.offsetParent;
          
          while( objParent.tagName.toUpperCase() != "BODY" )
          {
            objLeft  += objParent.offsetLeft;
            objTop   += objParent.offsetTop;
            objParent = objParent.offsetParent;
          }
      
          objHeight = obj.offsetHeight;
          objWidth = obj.offsetWidth;
      
          if(( PC_overDiv.offsetLeft + PC_overDiv.offsetWidth ) <= objLeft );
          else if(( PC_overDiv.offsetTop + PC_overDiv.offsetHeight ) <= objTop );
          else if( PC_overDiv.offsetTop >= ( objTop + objHeight ));
          else if( PC_overDiv.offsetLeft >= ( objLeft + objWidth ));
          else
          {
            obj.style.visibility = "hidden";
          }
        }
      }
    }
     
    /*
    * unhides <select> and <applet> objects (for PC_ie only)
    */
    function PC_showElement( PC_elmID )
    {
      if( PC_ie )
      {
        for( i = 0; i < document.all.tags( PC_elmID ).length; i++ )
        {
          obj = document.all.tags( PC_elmID )[i];
          
          if( !obj || !obj.offsetParent )
          {
            continue;
          }
        
          obj.style.visibility = "";
        }
      }
    }

	function PC_HolidayRec (d, m, y, desc)
	{
		this.d = d
		this.m = m
		this.y = y
		this.desc = desc
	}

	var HolidaysCounter = 0
	var Holidays = new Array()

	function PC_addHoliday (pc_d, pc_m, pc_y, pc_desc)
	{
		Holidays[HolidaysCounter++] = new PC_HolidayRec ( pc_d, pc_m, pc_y, pc_desc )
	}

	if (PC_dom)
	{
		for	(i=0;i<PC_PC_imgsrc.length;i++)
		{
			PC_img[i] = new Image
			PC_img[i].src = PC_PC_imgDir + PC_PC_imgsrc[i]
		}
		document.write ("<div onclick='PC_bShow=true' id='calendar'	style='z-index:+999;position:absolute;visibility:hidden;'><table	width="+((PC_showWeekNumber==1)?250:220)+" style='font-family:arial;font-size:11px;border-width:1;border-style:solid;border-color:#a0a0a0;font-family:arial; font-size:11px}' bgcolor='#ffffff'><tr background='imagenes/calback.jpg'><td background='imagenes/calback.jpg'><table width='"+((PC_showWeekNumber==1)?248:218)+"'><tr><td style='padding:2px;font-family:arial; font-size:11px;'><font color='#FFFFFF'><B><span id='caption'></span></B></font></td><td align=right><a href='javascript:PC_hideCalendar()'><PC_img SRC='"+PC_PC_imgDir+"close.gif' BORDER='0' ALT='Close the Calendar'></a></td></tr></table></td></tr><tr><td style='padding:5px' bgcolor=#ffffff><span id='content'></span></td></tr>")
			
		if (PC_showPC_today==1)
		{
			document.write ("<tr bgcolor=#f0f0f0><td style='padding:5px' align=center><span id='lblPC_today'></span></td></tr>")
		}
			
		document.write ("</table></div><div id='selectMonth' style='z-index:+999;position:absolute;visibility:hidden;'></div><div id='PC_selectYear' style='z-index:+999;position:absolute;visibility:hidden;'></div>");
	}

//---------------------------------------------TRADUCCION EN ESPAÑOL-----------------------------------------------------------------------------
	var	PC_monthName =	new	Array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","SetPC_iembre","Octubre","NovPC_iembre","DicPC_iembre")
	var	PC_monthName2 = new Array("ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC")
	if (PC_startAt==0)
	{
		dayName = new Array	("Lun","Mar","MPC_ie","Jue","VPC_ie","Sab","PC_dom")
	}
	else
	{
		dayName = new Array	("Lun","Mar","MPC_ie","Jue","VPC_ie","Sab","PC_dom")
	}
	var	PC_styleAnchor="text-decoration:none;color:black;"
	var	PC_styleLightBorder="border-style:solid;border-width:1px;border-color:#a0a0a0;"

	function PC_swapImage(srcPC_img, destPC_img){
		if (PC_ie)	{ document.getElementById(srcPC_img).setAttribute("src",PC_PC_imgDir + destPC_img) }
	}

	function PC_init()	{
		if (!PC_ns4)
		{
			if (!PC_ie) { PC_yearNow += 1900	}

			PC_crossobj=(PC_dom)?document.getElementById("calendar").style : PC_ie? document.all.calendar : document.calendar
			PC_hideCalendar()

			PC_crossMonthObj=(PC_dom)?document.getElementById("selectMonth").style : PC_ie? document.all.selectMonth	: document.selectMonth

			PC_crossYearObj=(PC_dom)?document.getElementById("PC_selectYear").style : PC_ie? document.all.PC_selectYear : document.PC_selectYear

			PC_monthConstructed=false;
			PC_yearConstructed=false;

			if (PC_showPC_today==1)
			{
				document.getElementById("lblPC_today").innerHTML =	PC_PC_todayString + " <a onmousemove='window.status=\""+PC_gotoString+"\"' onmouseout='window.status=\"\"' title='"+PC_gotoString+"' style='"+PC_styleAnchor+"' href='javascript:PC_monthSelected=PC_monthNow;PC_yearSelected=PC_yearNow;PC_constructCalendar();'>"+dayName[(PC_today.getDay()-PC_startAt==-1)?6:(PC_today.getDay()-PC_startAt)]+", " + PC_dateNow + " " + PC_monthName[PC_monthNow].substring(0,15)	+ "	" +	PC_yearNow	+ "</a>"
			}

			sHTML1="<span id='spanLeft'	style='border-style:solid;border-width:1;border-color:#ffffff;cursor:pointer' onmouseover='PC_swapImage(\"changeLeft\",\"left2.gif\");this.style.borderColor=\"#88AAFF\";window.status=\""+PC_scrollLeftMessage+"\"' onclick='javascript:PC_decMonth()' onmouseout='clearInterval(PC_intervalID1);PC_swapImage(\"changeLeft\",\"left1.gif\");this.style.borderColor=\"#ffffff\";window.status=\"\"' onmousedown='clearTimeout(PC_timeoutID1);PC_timeoutID1=setTimeout(\"PC_StartPC_decMonth()\",500)'	onmouseup='clearTimeout(PC_timeoutID1);clearInterval(PC_intervalID1)'>&nbsp<PC_img id='changeLeft' SRC='"+PC_PC_imgDir+"left1.gif' width=10 height=11 BORDER=0>&nbsp</span>&nbsp;"
			sHTML1+="<span id='spanRight' style='border-style:solid;border-width:1;border-color:#ffffff;cursor:pointer'	onmouseover='PC_swapImage(\"changeRight\",\"right2.gif\");this.style.borderColor=\"#88AAFF\";window.status=\""+PC_scrollRightMessage+"\"' onmouseout='clearInterval(PC_intervalID1);PC_swapImage(\"changeRight\",\"right1.gif\");this.style.borderColor=\"#ffffff\";window.status=\"\"' onclick='PC_incMonth()' onmousedown='clearTimeout(PC_timeoutID1);PC_timeoutID1=setTimeout(\"PC_StartPC_incMonth()\",500)'	onmouseup='clearTimeout(PC_timeoutID1);clearInterval(PC_intervalID1)'>&nbsp<PC_img id='changeRight' SRC='"+PC_PC_imgDir+"right1.gif'	width=10 height=11 BORDER=0>&nbsp</span>&nbsp"
			sHTML1+="<span id='spaPC_nMonth' style='border-style:solid;border-width:1;border-color:#ffffff;cursor:pointer'	onmouseover='PC_swapImage(\"changeMonth\",\"drop2.gif\");this.style.borderColor=\"#88AAFF\";window.status=\""+PC_selectMonthMessage+"\"' onmouseout='PC_swapImage(\"changeMonth\",\"drop1.gif\");this.style.borderColor=\"#ffffff\";window.status=\"\"' onclick='PC_popUpMonth()'></span>&nbsp;"
			sHTML1+="<span id='spaPC_nYear' style='border-style:solid;border-width:1;border-color:#fffffff;cursor:pointer' onmouseover='PC_swapImage(\"changeYear\",\"drop2.gif\");this.style.borderColor=\"#88AAFF\";window.status=\""+PC_PC_selectYearMessage+"\"'	onmouseout='PC_swapImage(\"changeYear\",\"drop1.gif\");this.style.borderColor=\"#ffffff\";window.status=\"\"'	onclick='PC_popUpYear()'></span>&nbsp;"
			
			document.getElementById("caption").innerHTML  =	sHTML1

			PC_bPageLoaded=true
		}
	}

	function PC_hideCalendar()	{
		PC_crossobj.visibility="hidden"
		if (PC_crossMonthObj != null){PC_crossMonthObj.visibility="hidden"}
		if (PC_crossYearObj !=	null){PC_crossYearObj.visibility="hidden"}

	    PC_showElement( 'SELECT' );
		PC_showElement( 'APPLET' );
	}

	function PC_padZero(num) {
		return (num	< 10)? '0' + num : num ;
	}

	function PC_constructDate(d,m,y)
	{
		sTmp = PC_dateFormat
		sTmp = sTmp.replace	("dd","<e>")
		sTmp = sTmp.replace	("d","<d>")
		sTmp = sTmp.replace	("<e>",PC_padZero(d))
		sTmp = sTmp.replace	("<d>",d)
		sTmp = sTmp.replace	("mmmm","<p>")
		sTmp = sTmp.replace	("mmm","<o>")
		sTmp = sTmp.replace	("mm","<n>")
		sTmp = sTmp.replace	("m","<m>")
		sTmp = sTmp.replace	("<m>",m+1)
		sTmp = sTmp.replace	("<n>",PC_padZero(m+1))
		sTmp = sTmp.replace	("<o>",PC_monthName[m])
		sTmp = sTmp.replace	("<p>",PC_monthName2[m])
		sTmp = sTmp.replace	("yyyy",y)
		return sTmp.replace ("yy",PC_padZero(y%100))
	}

	function PC_closeCalendar() {
		var	sTmp

		PC_hideCalendar();
		PC_PC_ctlToPlaceValue.value =	PC_constructDate(PC_dateSelected,PC_monthSelected,PC_yearSelected)
		if(typeof fin_calendario == 'function') 
		{ 
			fin_calendario(); 
		} 
	}

	/*** Month Pulldown	***/

	function PC_StartPC_decMonth()
	{
		PC_intervalID1=setInterval("PC_decMonth()",80)
	}

	function PC_StartPC_incMonth()
	{
		PC_intervalID1=setInterval("PC_incMonth()",80)
	}

	function PC_incMonth () {
		PC_monthSelected++
		if (PC_monthSelected>11) {
			PC_monthSelected=0
			PC_yearSelected++
		}
		PC_constructCalendar()
	}

	function PC_decMonth () {
		PC_monthSelected--
		if (PC_monthSelected<0) {
			PC_monthSelected=11
			PC_yearSelected--
		}
		PC_constructCalendar()
	}

	function PC_constructMonth() {
		PC_popDowPC_nYear()
		if (!PC_monthConstructed) {
			sHTML =	""
			for	(i=0; i<12;	i++) {
				sName =	PC_monthName[i];
				if (i==PC_monthSelected){
					sName =	"<B>" +	sName +	"</B>"
				}
				sHTML += "<tr><td id='m" + i + "' onmouseover='this.style.backgroundColor=\"#FFCC99\"' onmouseout='this.style.backgroundColor=\"\"' style='cursor:pointer' onclick='PC_monthConstructed=false;PC_monthSelected=" + i + ";PC_constructCalendar();PC_popDowPC_nMonth();event.cancelBubble=true'>&nbsp;" + sName + "&nbsp;</td></tr>"
			}

			document.getElementById("selectMonth").innerHTML = "<table width=70	style='font-family:arial; font-size:11px; border-width:1; border-style:solid; border-color:#a0a0a0;' bgcolor='#FFFFDD' cellspacing=0 onmouseover='clearTimeout(PC_timeoutID1)'	onmouseout='clearTimeout(PC_timeoutID1);PC_timeoutID1=setTimeout(\"PC_popDowPC_nMonth()\",100);event.cancelBubble=true'>" +	sHTML +	"</table>"

			PC_monthConstructed=true
		}
	}

	function PC_popUpMonth() {
		PC_constructMonth()
		PC_crossMonthObj.visibility = (PC_dom||PC_ie)? "visible"	: "show"
		PC_crossMonthObj.left = parseInt(PC_crossobj.left) + 50
		PC_crossMonthObj.top =	parseInt(PC_crossobj.top) + 26

		PC_hideElement( 'SELECT', document.getElementById("selectMonth") );
		PC_hideElement( 'APPLET', document.getElementById("selectMonth") );			
	}

	function PC_popDowPC_nMonth()	{
		PC_crossMonthObj.visibility= "hidden"
	}

	/*** Year Pulldown ***/

	function PC_incYear() {
	}

	function PC_decYear() {
	}

	function PC_selectYear(PC_nYear) {
		PC_yearSelected=parseInt(PC_nYear);
		PC_yearConstructed=false;
		PC_constructCalendar();
		PC_popDowPC_nYear();
	}

	function PC_constructYear() {
		PC_popDowPC_nMonth()
		sHTML =	""
		if (!PC_yearConstructed) {

			sHTML =	""

			PC_nStartingYear =	PC_yearNow;
			
//Dentro de el siguPC_iente ciclo se generan los años, en el caso de que sea introcir fecha de nacimPC_iento restale 30---- i>=(PC_yearNow-30)-----
//CAMBIA el siguPC_iente for por este " for	(i=PC_yearNow; i<=(PC_yearNow+4); i++) { " si quPC_ieres que los años empizen en el actual en adelante		
			for	(i=PC_yearNow; i>=(PC_yearNow-120); i--) {		
				sName =	i;
				if (i==PC_yearSelected){
					sName =	"<B>" +	sName +	"</B>"
				}

				sHTML += "<tr><td id='y" + sName + "' onmouseover='this.style.backgroundColor=\"#FFCC99\"' onmouseout='this.style.backgroundColor=\"\"' style='cursor:pointer' onclick='PC_selectYear("+i+");event.cancelBubble=true'>&nbsp;" + sName + "&nbsp;</td></tr>"
			}

			sHTML += ""

			document.getElementById("PC_selectYear").innerHTML	= "<table width=44 style='font-family:arial; font-size:11px; border-width:1; border-style:solid; border-color:#a0a0a0;'	bgcolor='#FFFFDD' onmouseover='clearTimeout(PC_timeoutID2)' onmouseout='clearTimeout(PC_timeoutID2);PC_timeoutID2=setTimeout(\"PC_popDowPC_nYear()\",100)' cellspacing=0>"	+ sHTML	+ "</table>"

			PC_yearConstructed	= true
		}
	}

	function PC_popDowPC_nYear() {
		clearInterval(PC_intervalID1)
		clearTimeout(PC_timeoutID1)
		clearInterval(PC_intervalID2)
		clearTimeout(PC_timeoutID2)
		PC_crossYearObj.visibility= "hidden"
	}

	function PC_popUpYear() {
		var	leftOffset

		PC_constructYear()
		PC_crossYearObj.visibility	= (PC_dom||PC_ie)? "visible" : "show"
		leftOffset = parseInt(PC_crossobj.left) + document.getElementById("spaPC_nYear").offsetLeft
		if (PC_ie)
		{
			leftOffset += 6
		}
		PC_crossYearObj.left =	leftOffset
		PC_crossYearObj.top = parseInt(PC_crossobj.top) + 26
	}

	/*** calendar ***/
   function PC_WeekNbr(pc_n) {
      // Algorithm used:
      // From Klaus Tondering's Calendar document (The Authority/Guru)
      // hhtp://www.tondering.dk/claus/calendar.html
      // a = (14-month) / 12
      // y = year + 4800 - a
      // m = month + 12a - 3
      // J = day + (153m + 2) / 5 + 365y + y / 4 - y / 100 + y / 400 - 32045
      // d4 = (J + 31741 - (J mod 7)) mod 146097 mod 36524 mod 1461
      // L = d4 / 1460
      // d1 = ((d4 - L) mod 365) + L
      // WeekNumber = d1 / 7 + 1
 
      year = pc_n.getFullYear();
      month = pc_n.getMonth() + 1;
      if (PC_startAt == 0) {
         day = pc_n.getDate() + 1;
      }
      else {
         day = pc_n.getDate();
      }
 
      a = Math.floor((14-month) / 12);
      y = year + 4800 - a;
      m = month + 12 * a - 3;
      b = Math.floor(y/4) - Math.floor(y/100) + Math.floor(y/400);
      J = day + Math.floor((153 * m + 2) / 5) + 365 * y + b - 32045;
      d4 = (((J + 31741 - (J % 7)) % 146097) % 36524) % 1461;
      L = Math.floor(d4 / 1460);
      d1 = ((d4 - L) % 365) + L;
      week = Math.floor(d1/7) + 1;
 
      return week;
   }

	function PC_constructCalendar () {
		var aNumDays = Array (31,0,31,30,31,30,31,31,30,31,30,31)

		var dateMessage
		var	startDate =	new	Date (PC_yearSelected,PC_monthSelected,1)
		var endDate

		if (PC_monthSelected==1)
		{
			endDate	= new Date (PC_yearSelected,PC_monthSelected+1,1);
			endDate	= new Date (endDate	- (24*60*60*1000));
			numDaysIPC_nMonth = endDate.getDate()
		}
		else
		{
			numDaysIPC_nMonth = aNumDays[PC_monthSelected];
		}

		datePointer	= 0
		dayPointer = startDate.getDay() - PC_startAt
		
		if (dayPointer<0)
		{
			dayPointer = 6
		}

		sHTML =	"<table	 border=0 style='font-family:verdana;font-size:10px;'><tr>"

		if (PC_showWeekNumber==1)
		{
			sHTML += "<td width=27><b>" + PC_weekString + "</b></td><td width=1 rowspan=7 bgcolor='#d0d0d0' style='padding:0px'><PC_img src='"+PC_PC_imgDir+"divider.gif' width=1></td>"
		}

		for	(i=0; i<7; i++)	{
			sHTML += "<td width='27' align='right'><B>"+ dayName[i]+"</B></td>"
		}
		sHTML +="</tr><tr>"
		
		if (PC_showWeekNumber==1)
		{
			sHTML += "<td align=right>" + PC_WeekNbr(startDate) + "&nbsp;</td>"
		}

		for	( var i=1; i<=dayPointer;i++ )
		{
			sHTML += "<td>&nbsp;</td>"
		}
	
		for	( datePointer=1; datePointer<=numDaysIPC_nMonth; datePointer++ )
		{
			dayPointer++;
			sHTML += "<td align=right>"
			sStyle=PC_styleAnchor
			if ((datePointer==PC_dateSelected) &&	(PC_monthSelected==PC_monthSelected)	&& (PC_yearSelected==PC_yearSelected))
			{ sStyle+=PC_styleLightBorder }

			sHint = ""
			for (k=0;k<HolidaysCounter;k++)
			{
				if ((parseInt(Holidays[k].d)==datePointer)&&(parseInt(Holidays[k].m)==(PC_monthSelected+1)))
				{
					if ((parseInt(Holidays[k].y)==0)||((parseInt(Holidays[k].y)==PC_yearSelected)&&(parseInt(Holidays[k].y)!=0)))
					{
						sStyle+="background-color:#FFDDDD;"
						sHint+=sHint==""?Holidays[k].desc:"\n"+Holidays[k].desc
					}
				}
			}

			var regexp= /\"/g
			sHint=sHint.replace(regexp,"&quot;")

			dateMessage = "onmousemove='window.status=\""+PC_selectDateMessage.replace("[date]",PC_constructDate(datePointer,PC_monthSelected,PC_yearSelected))+"\"' onmouseout='window.status=\"\"' "

			var f1= new Date(PC_yearSelected,PC_monthSelected,datePointer);
//------------------------------------TERMINAR LA FUNCION EN 0 PARA QUE INICPC_ie EN EL DIA ACTUAL EN LA SIGUPC_ieNTE LINEA----------------------------------------------
			var PC_temp = addToDate(PC_makePC_dateFormat(PC_dateNow, PC_monthNow, PC_yearNow),0);
			var nDia = Number(PC_temp.substr(0, 2)); 
   			var PC_nMes = Number(PC_temp.substr(3, 2)); 
   			var PC_nAno = Number(PC_temp.substr(6, 4)); 
			var fa= new Date(PC_nAno,PC_nMes,nDia);
			
			if ((datePointer==nDia)&&(PC_monthSelected==PC_nMes)&&(PC_yearSelected==PC_nAno))
			//{ sHTML += "<b style='"+sStyle+"'><font color=#ff0000>&nbsp;" + datePointer + "</font>&nbsp;</b>"}
			{ sHTML += "<b><a "+dateMessage+" title=\"" + sHint + "\" style='"+sStyle+"' href='javascript:PC_dateSelected="+datePointer+";PC_closeCalendar();'><font color=#ff0000>&nbsp;" + datePointer + "</font>&nbsp;</a></b>"}
			//else if	(dayPointer % 7 == (PC_startAt * -1)+1)
			//{ sHTML += "<a "+dateMessage+" title=\"" + sHint + "\" style='"+sStyle+"' href='javascript:PC_dateSelected="+datePointer + ";PC_closeCalendar();'>&nbsp;<font color=#909090>" + datePointer + "</font>&nbsp;</a>" }
			else if (PC_FechaMenor(fa,f1)) //((datePointer<PC_dateNow)&&(PC_monthSelected<=PC_monthNow)&&(PC_yearSelected=PC_yearNow))
			{ sHTML += "&nbsp;<font color=#909090>" + datePointer + "</font>&nbsp;" }
			else
			{ sHTML += "<a "+dateMessage+" title=\"" + sHint + "\" style='"+sStyle+"' href='javascript:PC_dateSelected="+datePointer + ";PC_closeCalendar();'>&nbsp;" + datePointer + "&nbsp;</a>" }

			sHTML += ""
			if ((dayPointer+PC_startAt) % 7 == PC_startAt) { 
				sHTML += "</tr><tr>" 
				if ((PC_showWeekNumber==1)&&(datePointer<numDaysIPC_nMonth))
				{
					sHTML += "<td align=right>" + (PC_WeekNbr(new Date(PC_yearSelected,PC_monthSelected,datePointer+1))) + "&nbsp;</td>"
				}
			}
		}

		document.getElementById("content").innerHTML   = sHTML
		document.getElementById("spaPC_nMonth").innerHTML = "&nbsp;" +	PC_monthName[PC_monthSelected] + "&nbsp;<PC_img id='changeMonth' SRC='"+PC_PC_imgDir+"drop1.gif' WIDTH='12' HEIGHT='10' BORDER=0>"
		document.getElementById("spaPC_nYear").innerHTML =	"&nbsp;" + PC_yearSelected	+ "&nbsp;<PC_img id='changeYear' SRC='"+PC_PC_imgDir+"drop1.gif' WIDTH='12' HEIGHT='10' BORDER=0>"
	}

	function popUpCalendar(PC_ctl,	PC_ctl2, format) {
		var	leftpos=0
		var	toppos=0

		if (PC_bPageLoaded)
		{
			if ( PC_crossobj.visibility ==	"hidden" ) {
				PC_PC_ctlToPlaceValue	= PC_ctl2
				PC_dateFormat=format;

				formatChar = " "
				aFormat	= PC_dateFormat.split(formatChar)
				if (aFormat.length<3)
				{
					formatChar = "/"
					aFormat	= PC_dateFormat.split(formatChar)
					if (aFormat.length<3)
					{
						formatChar = "."
						aFormat	= PC_dateFormat.split(formatChar)
						if (aFormat.length<3)
						{
							formatChar = "-"
							aFormat	= PC_dateFormat.split(formatChar)
							if (aFormat.length<3)
							{
								// invalid date	format
								formatChar=""
							}
						}
					}
				}

				tokensChanged =	0
				if ( formatChar	!= "" )
				{
					// use user's date
					aData =	PC_ctl2.value.split(formatChar)

					for	(i=0;i<3;i++)
					{
						if ((aFormat[i]=="d") || (aFormat[i]=="dd"))
						{
							PC_dateSelected = parseInt(aData[i], 10)
							tokensChanged ++
						}
						else if	((aFormat[i]=="m") || (aFormat[i]=="mm"))
						{
							PC_monthSelected =	parseInt(aData[i], 10) - 1
							tokensChanged ++
						}
						else if	(aFormat[i]=="yyyy")
						{
							PC_yearSelected = parseInt(aData[i], 10)
							tokensChanged ++
						}
						else if	(aFormat[i]=="mmm")
						{
							for	(j=0; j<12;	j++)
							{
								if (aData[i]==PC_monthName[j])
								{
									PC_monthSelected=j
									tokensChanged ++
								}
							}
						}
						else if	(aFormat[i]=="mmmm")
						{
							for	(j=0; j<12;	j++)
							{
								if (aData[i]==PC_monthName2[j])
								{
									PC_monthSelected=j
									tokensChanged ++
								}
							}
						}
					}
				}

				if ((tokensChanged!=3)||isNaN(PC_dateSelected)||isNaN(PC_monthSelected)||isNaN(PC_yearSelected))
				{
//------------------------------------TERMINAR LA FUNCION EN 0 PARA QUE INICPC_ie EN EL DIA ACTUAL EN LA SIGUPC_ieNTE LINEA----------------------------------------------
					var PC_temp = addToDate(PC_makePC_dateFormat(PC_dateNow, PC_monthNow, PC_yearNow),0);
					var nDia = Number(PC_temp.substr(0, 2)); 
   					var PC_nMes = Number(PC_temp.substr(3, 2)); 
   					var PC_nAno = Number(PC_temp.substr(6, 4)); 
			
					PC_dateSelected = nDia
					PC_monthSelected =	PC_nMes
					PC_yearSelected = PC_nAno
				}

				PC_dateSelected=PC_dateSelected
				PC_monthSelected=PC_monthSelected
				PC_yearSelected=PC_yearSelected

				aTag = PC_ctl
				do {
					aTag = aTag.offsetParent;
					leftpos	+= aTag.offsetLeft;
					toppos += aTag.offsetTop;
				} while(aTag.tagName!="BODY");

				PC_crossobj.left =	PC_fixedX==-1 ? PC_ctl.offsetLeft	+ leftpos :	PC_fixedX
				PC_crossobj.top = PC_fixedY==-1 ?	PC_ctl.offsetTop +	toppos + PC_ctl.offsetHeight +	2 :	PC_fixedY
				PC_constructCalendar (1, PC_monthSelected, PC_yearSelected);
				PC_crossobj.visibility=(PC_dom||PC_ie)? "visible" : "show"

				PC_hideElement( 'SELECT', document.getElementById("calendar") );
				PC_hideElement( 'APPLET', document.getElementById("calendar") );			

				PC_bShow = true;
			}
			else
			{
				PC_hideCalendar()
				if (PC_PC_ctlNow!=PC_ctl) {popUpCalendar(PC_ctl, PC_ctl2, format)}
			}
			PC_PC_ctlNow = PC_ctl
		}
		
	}

	document.onkeypress = function hidecal1 () { 
		if (event.keyCode==27) 
		{
			PC_hideCalendar()
		}
	}
	document.onclick = function hidecal2 () { 		
		if (!PC_bShow)
		{
			PC_hideCalendar()
		}
		PC_bShow = false
	}

	if(PC_ie)
	{
		PC_init()
	}
	else
	{
		window.onload=PC_init
	}

// ************ ADICIONAR DIAS A UNA FECHA ***********

  var aFiPC_nMes = new Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31); 

  function fiPC_nMes(PC_nMes, PC_nAno){ 
   return aFiPC_nMes[PC_nMes - 1] + (((PC_nMes == 2) && (PC_nAno % 4) == 0)? 1: 0); 
  } 

   function PC_padNmb(nStr, nLen, sChr){ 
    var sRes = String(nStr); 
    for (var i = 0; i < nLen - String(nStr).length; i++) 
     sRes = sChr + sRes; 
    return sRes; 
   } 

   function PC_makePC_dateFormat(PC_nDay, PC_nMonth, PC_nYear){ 
    var sRes; 
    sRes = PC_padNmb(PC_nDay, 2, "0") + "/" + PC_padNmb(PC_nMonth, 2, "0") + "/" + PC_padNmb(PC_nYear, 4, "0"); 
    return sRes; 
   } 
    
  function PC_incDate(PC_sFec0){ 
   var nDia = parseInt(PC_sFec0.substr(0, 2), 10); 
   var PC_nMes = parseInt(PC_sFec0.substr(3, 2), 10); 
   var PC_nAno = parseInt(PC_sFec0.substr(6, 4), 10); 
   nDia += 1; 
   if (nDia > fiPC_nMes(PC_nMes, PC_nAno)){ 
    nDia = 1; 
    PC_nMes += 1; 
    if (PC_nMes == 13){ 
     PC_nMes = 1; 
     PC_nAno += 1; 
    } 
   } 
   return PC_makePC_dateFormat(nDia, PC_nMes, PC_nAno); 
  } 

  function PC_decDate(PC_sFec0){ 
   var nDia = Number(PC_sFec0.substr(0, 2)); 
   var PC_nMes = Number(PC_sFec0.substr(3, 2)); 
   var PC_nAno = Number(PC_sFec0.substr(6, 4)); 
   nDia -= 1; 
   if (nDia == 0){ 
    PC_nMes -= 1; 
    if (PC_nMes == 0){ 
     PC_nMes = 12; 
     PC_nAno -= 1; 
    } 
    nDia = fiPC_nMes(PC_nMes, PC_nAno); 
   } 
   return PC_makePC_dateFormat(nDia, PC_nMes, PC_nAno); 
  } 

  function addToDate(PC_sFec0, PC_sInc){ 
   var nInc = Math.abs(parseInt(PC_sInc)); 
   var sRes = PC_sFec0; 
   if (parseInt(PC_sInc) >= 0) 
    for (var i = 0; i < nInc; i++) sRes = PC_incDate(sRes); 
   else 
    for (var i = 0; i < nInc; i++) sRes = PC_decDate(sRes); 
   return sRes; 
  } 

  function PC_recalcF1(){ 
   with (document.formulario){ 
    fecha1.value = addToDate(fecha0.value, increm.value); 
   } 
  } 

function PC_ValidarFecha()
{
	/*var PC_temp = new Date()
	fechaActual= (PC_temp.getDate() + 1) + "/";
	fechaActual += PC_temp.getMonth() + "/";
    fechaActual += PC_temp.getFullYear();

	if (( addToDate(fechaActual, 2) < document.form1.date1.value )  && (document.form1.date1.value > fechaActual ) ){
		return true
		}
	else {
		return false;
	} */
	return true;
}



//----------------------------Eliminar las siguPC_ientes lineas para activar todos los dias del año-----y retornar un false-----------------------------------------
function PC_FechaMenor(fechaActual, fecha1) {
	/* 
			var msegActual = fechaActual.getTime();
			var msegFecha1 = fecha1.getTime();
			var Diferencia = msegActual - msegFecha1;
			Diferencia /= 86400000;
			if (Diferencia < 0) {
				return false;
			}
			else {
				return true;
			}*/
			return false;
		}
		

