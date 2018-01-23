<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar

	if permisofacultad("repimpagadoglobal.asp") then
		fechainicio=obtener("fechainicio")
		fechafin=obtener("fechafin")
		if fechainicio="" then
			fechainicio=dateadd("m",-1,Date())
			fechafin=Date()
		end if
		filtrobuscador=" and (A.fechagestion>='" & fechainicio & " 00:00:00' and A.fechagestion<='" & fechafin & " 23:59:59') "
		
		sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaExportar' or descripcion='RutaWebExportar'"
		consultar sql,RS
		RS.Filter=" descripcion='RutaFisicaExportar'"
		RutaFisicaExportar=RS.Fields(1)
		RS.Filter=" descripcion='RutaWebExportar'"
		RutaWebExportar=RS.Fields(1)				
		RS.Filter=""
		RS.Close	
		tiempoexport=Now()		
		
sql="select Linea,nro " & _
		"into ##repimpaglobal" & session("codusuario") & " " & _
		"from " & _
		"( " & _
		"select '<xml xmlns:s=" & chr(34) & "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882" & chr(34) & " xmlns:dt=" & chr(34) & "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882" & chr(34) & " xmlns:rs=" & chr(34) & "urn:schemas-microsoft-com:rowset" & chr(34) & " xmlns:z=" & chr(34) & "#RowsetSchema" & chr(34) & ">' as Linea,1 as nro " & _	
		"union select '<s:Schema id=" & chr(34) & "RowsetSchema" & chr(34) & ">' as Linea,2 as nro " & _	
		"union select '<s:ElementType name=" & chr(34) & "row" & chr(34) & " content=" & chr(34) & "eltOnly" & chr(34) & ">' as Linea,3 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "IDFactura" & chr(34) & " rs:number=" & chr(34) & "1" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,4 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "int" & chr(34) & " dt:maxLength=" & chr(34) & "4" & chr(34) & " rs:precision=" & chr(34) & "10" & chr(34) & " rs:fixedlength=" & chr(34) & "true" & chr(34) & "/>' as Linea,5 as nro " & _	
		"union select '</s:AttributeType>' as Linea,6 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "Fecha" & chr(34) & " rs:number=" & chr(34) & "2" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,7 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "dateTime" & chr(34) & " rs:dbtype=" & chr(34) & "variantdate" & chr(34) & " dt:maxLength=" & chr(34) & "16" & chr(34) & " rs:fixedlength=" & chr(34) & "true" & chr(34) & "/>' as Linea,8 as nro " & _	
		"union select '</s:AttributeType>' as Linea,9 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "Cliente" & chr(34) & " rs:number=" & chr(34) & "3" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,10 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,11 as nro " & _	
		"union select '</s:AttributeType>' as Linea,12 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "TipoDoc" & chr(34) & " rs:number=" & chr(34) & "4" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,13 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,14 as nro " & _	
		"union select '</s:AttributeType>' as Linea,15 as nro " & _			
		"union select '<s:AttributeType name=" & chr(34) & "NumFactura" & chr(34) & " rs:number=" & chr(34) & "5" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,16 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,17 as nro " & _	
		"union select '</s:AttributeType>' as Linea,18 as nro " & _		
		"union select '<s:AttributeType name=" & chr(34) & "Articulo" & chr(34) & " rs:number=" & chr(34) & "6" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,19 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,20 as nro " & _	
		"union select '</s:AttributeType>' as Linea,21 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "Unidad" & chr(34) & " rs:number=" & chr(34) & "7" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,22 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,23 as nro " & _	
		"union select '</s:AttributeType>' as Linea,24 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "PrecioSoles" & chr(34) & " rs:number=" & chr(34) & "8" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,25 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "float" & chr(34) & " rs:dbtype=" & chr(34) & "currency" & chr(34) & " dt:maxLength=" & chr(34) & "8" & chr(34) & " rs:precision=" & chr(34) & "53" & chr(34) & " rs:fixedlength=" & chr(34) & "true" & chr(34) & "/>' as Linea,26 as nro " & _	
		"union select '</s:AttributeType>' as Linea,27 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "Cantidad" & chr(34) & " rs:number=" & chr(34) & "9" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,28 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "float" & chr(34) & " rs:dbtype=" & chr(34) & "currency" & chr(34) & " dt:maxLength=" & chr(34) & "8" & chr(34) & " rs:precision=" & chr(34) & "53" & chr(34) & " rs:fixedlength=" & chr(34) & "true" & chr(34) & "/>' as Linea,29 as nro " & _	
		"union select '</s:AttributeType>' as Linea,30 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "TotalSoles" & chr(34) & " rs:number=" & chr(34) & "10" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,31 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "float" & chr(34) & " rs:dbtype=" & chr(34) & "currency" & chr(34) & " dt:maxLength=" & chr(34) & "8" & chr(34) & " rs:precision=" & chr(34) & "53" & chr(34) & " rs:fixedlength=" & chr(34) & "true" & chr(34) & "/>' as Linea,32 as nro " & _	
		"union select '</s:AttributeType>' as Linea,33 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "Estado" & chr(34) & " rs:number=" & chr(34) & "11" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,34 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,35 as nro " & _	
		"union select '</s:AttributeType>' as Linea,36 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "Credito" & chr(34) & " rs:number=" & chr(34) & "12" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,37 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,38 as nro " & _	
		"union select '</s:AttributeType>' as Linea,39 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "FechaVcto" & chr(34) & " rs:number=" & chr(34) & "13" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,40 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "dateTime" & chr(34) & " rs:dbtype=" & chr(34) & "variantdate" & chr(34) & " dt:maxLength=" & chr(34) & "16" & chr(34) & " rs:fixedlength=" & chr(34) & "true" & chr(34) & "/>' as Linea,41 as nro " & _	
		"union select '</s:AttributeType>' as Linea,42 as nro " & _	
		"union select '<s:AttributeType name=" & chr(34) & "Vendedor" & chr(34) & " rs:number=" & chr(34) & "14" & chr(34) & " rs:nullable=" & chr(34) & "true" & chr(34) & " rs:maydefer=" & chr(34) & "true" & chr(34) & " rs:writeunknown=" & chr(34) & "true" & chr(34) & ">' as Linea,43 as nro " & _	
		"union select '<s:datatype dt:type=" & chr(34) & "string" & chr(34) & " dt:maxLength=" & chr(34) & "250" & chr(34) & "/>' as Linea,44 as nro " & _	
		"union select '</s:AttributeType>' as Linea,45 as nro " & _			
		"union select '<s:extends type=" & chr(34) & "rs:rowbase" & chr(34) & "/>' as Linea,46 as nro " & _					
		"union select '</s:ElementType>' as Linea,47 as nro " & _					
		"union select '</s:Schema>' as Linea,48 as nro " & _					
		"union select '<rs:data>' as Linea,49 as nro " & _		
		"union " & _ 
		"select '<z:row' + " & _
		"					' IDFactura=" & chr(34) & "' + rtrim(ltrim(str(IDFactura))) + '" & chr(34) & "' + " & _ 
		"					' Fecha=" & chr(34) & "' + rtrim(ltrim(str(Year(Fecha))))  + '-' + Right('00' + rtrim(ltrim(str(Month(Fecha)))),2) + '-' + Right('00' + rtrim(ltrim(str(Day(Fecha)))),2) + '" & chr(34) & "' + " & _
		"					' Cliente=" & chr(34) & "' + replace(replace(Cliente,'&','&amp;'),'" & chr(34) & "','´´') + '" & chr(34) & "' + " & _
		"					' TipoDoc=" & chr(34) & "'  + TipoDoc + '" & chr(34) & "' + " & _
		"					' NumFactura=" & chr(34) & "'  + NumFactura + '" & chr(34) & "' + " & _ 
		"					' Articulo=" & chr(34) & "' + replace(replace(Articulo,'&','&amp;'),'" & chr(34) & "','´´') + '" & chr(34) & "' + " & _
		"					' Unidad=" & chr(34) & "'  + Unidad + '" & chr(34) & "' + " & _
		"					' PrecioSoles=" & chr(34) & "' + rtrim(ltrim(convert(varchar,Round(PrecioSoles,4)))) + '" & chr(34) & "' + " & _
		"					' Cantidad=" & chr(34) & "' + rtrim(ltrim(convert(varchar,Round(Cantidad,4)))) + '" & chr(34) & "' + " & _
		"					' TotalSoles=" & chr(34) & "' + rtrim(ltrim(convert(varchar,Round(TotalSoles,4)))) + '" & chr(34) & "' + " & _ 
		"					' Estado=" & chr(34) & "' + Estado + '" & chr(34) & "' + " & _ 
		"					' Credito=" & chr(34) & "'  + Credito + '" & chr(34) & "' + " & _ 
		"					' FechaVcto=" & chr(34) & "' + rtrim(ltrim(str(Year(FechaVcto)))) + '-' + Right('00' + rtrim(ltrim(str(Month(FechaVcto)))),2) + '-' + Right('00' + rtrim(ltrim(str(Day(FechaVcto)))),2) + '" & chr(34) & "' + " & _
		"					' Vendedor=" & chr(34) & "' + replace(replace(Vendedor,'&','&amp;'),'" & chr(34) & "','´´') + '" & chr(34) & "' + " & _
		"					'/>' " & _
		"as Linea,nro from " & _
		"( " & _
		"select DiasVencidoReal as IDFactura, FGestion as Fecha,codcen as Cliente,TipoDoc,Contrato as NumFactura,Producto as Articulo,CodOfi as Unidad,SaldoHoySoles as PrecioSoles,SaldoHoy as Cantidad,SaldoHoySoles as TotalSoles,'R' as Estado,'T' as Credito,FechaIncumpl as FechaVcto,Agencia as Vendedor,49 + Row_Number() OVER (order by A.FGestion) as nro " & _
		"from VISTAGLOBALDISTRIBUIDO A " & _ 
		") as Temp1 " & _
		"union select '</rs:data>' as Linea,99999999999999999998 as nro " & _
		"union select '</xml>' as Linea,99999999999999999999 as nro " & _
		") as Temp2 " & _
		"DECLARE @sql NVARCHAR(4000) " & _ 
		"DECLARE @SERVIDOR VARCHAR(1000),@USUARIO VARCHAR(1000),@CLAVE VARCHAR(1000) " & _
		"DECLARE @RUTAARCHIVO VARCHAR(1000) " & _
		"set @RUTAARCHIVO='" & RutaFisicaExportar & "\repimpaglobal" & session("codusuario") & ".xml' " & _
		"set @SERVIDOR='" & conn_server & "' " & _
		"set @USUARIO='" & conn_uid & "' " & _
		"set @CLAVE='" & conn_pwd & "' " & _
		"set @sql='master.dbo.xp_cmdshell ''bcp " & chr(34) & "select Linea from ##repimpaglobal" & session("codusuario") & " order by nro" & chr(34) & " queryout " & chr(34) & "' + @RUTAARCHIVO + '" & chr(34) & " -T -S' + @SERVIDOR  + ' -U' + @USUARIO + ' -P' + @CLAVE + ' -w''' " & _
		"EXEC (@sql) " & _		
		"set @sql='master.dbo.xp_cmdshell ''CACLS " & RutaFisicaExportar & "\repvtasdet" & session("codusuario") & ".xml /e /g Todos:F''' " & _	
		"EXEC (@sql) " & _
		"drop table ##repimpaglobal" & session("codusuario")
		conn.execute sql
		
	''"set @sql='master.dbo.xp_cmdshell ''CACLS " & Server.MapPath ("exportados/repvtasdet" & session("codusuario") & ".xml") & " /e /g USER-E9947CA976\IUSR_USER-E9947CA976:F''' " & _
	%>
		<html>
		<head>
		<script language='javascript' src="scripts/popcalendar.js"></script> 
		<style>
		A {
			FONT-SIZE: 12px; COLOR: #483d8b; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
		}
		A:visited {
			TEXT-DECORATION: none; COLOR: #483d8b;
		}
		A:hover {
			COLOR: #483d8b; FONT-FACE:"Arial"; TEXT-DECORATION: none
		}
		.skin {
		position:absolute;
		top: 0px;
		font-color:#FFFFFF;
		font-size:12px;
		width:78%;
		height:50%;
		border:0px none #667ec5;
		text-align:left;
		font-family:Arial;
		line-height:16px;
		cursor:hand;
		visibility:hidden;
		background:#FFFFFF;
		}
		TABLE
		{
			border-width: 0px;
			border-style: none;
		}
		TH 
		{
			color:#483d8b;
			background: #dcdcdc;
			font-size:12px;
			font-family:Arial;
			cursor:hand;
		}
		TD
		{
			color:#483d8b;
			font-size:12px;
			font-family:Arial;
		}
		</style>
		<!--TR
		{
			background: #FFFFFF;
		}-->		
		</head>
		
		<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula id="frmMain" method=post action="repvtasdet.asp" onmousemove="if(crossobj.visibility=='hidden'){document.getElementById('filavariable').height='40'};">
			<table width=100% height=100% cellpadding=4 cellspacing=0 border=0>	
			<tr>
				<td bgcolor="#dcdcdc" height=40 valign=top id="filavariable"><font size=2 face=Arial color=#483d8b><b>Desde: <input name="fechainicio" readonly id="dateArrival" onclick="document.getElementById('filavariable').height='100%';popUpCalendar(this, formula.dateArrival, 'dd/mm/yyyy');" type=text maxlength=10 size=10 value="<%=fechainicio%>" style="font-size: xx-small;"> Hasta:  <input name="fechafin" readonly id="dateArrival1" onclick="document.getElementById('filavariable').height='100%';popUpCalendar(this, formula.dateArrival1, 'dd/mm/yyyy');" type=text maxlength=10 size=10 value="<%=fechafin%>" style="font-size: xx-small;"></b></font>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a></td>
			</tr>			
			<tr>
				<td bgcolor="#dcdcdc" height=100% valign=top>
					<OBJECT name="DCube1" id="DCube1" style="WIDTH: 100%; HEIGHT:100%;" classid="clsid:6D63F73D-3688-3000-9C0F-00A0C90F29FC" >
						<PARAM NAME="_ExtentX" VALUE="17939">
						<PARAM NAME="_ExtentY" VALUE="9499">
						<PARAM NAME="DataSource" VALUE="">
						<PARAM NAME="RowAlignment" VALUE="0">
						<PARAM NAME="ColAlignment" VALUE="0">
						<PARAM NAME="RowStyle" VALUE="1">
						<PARAM NAME="ColStyle" VALUE="1">
						<PARAM NAME="OutlineIconAlignment" VALUE="1">
						<PARAM NAME="GridColor" VALUE="12632256">
						<PARAM NAME="BackColor" VALUE="16777215">
						<PARAM NAME="DCConnect" VALUE="">
						<PARAM NAME="DCDatabaseName" VALUE="">
						<PARAM NAME="CursorStyle" VALUE="0">
						<PARAM NAME="FieldsBackColor" VALUE="8421504">
						<PARAM NAME="FieldsForeColor" VALUE="16777215">
						<PARAM NAME="HeadingsForeColor" VALUE="0">
						<PARAM NAME="HeadingsBackColor" VALUE="16777215">
						<PARAM NAME="DCRecordSource" VALUE="">
						<PARAM NAME="TotalsBackColor" VALUE="16777215">
						<PARAM NAME="TotalsForeColor" VALUE="0">
						<PARAM NAME="GridStyle" VALUE="1">
						<PARAM NAME="ForeColor" VALUE="0">
						<PARAM NAME="AllowFiltering" VALUE="-1">
						<PARAM NAME="AllowUserPivotFields" VALUE="-1">
						<PARAM NAME="LeftMargin" VALUE="0.75">
						<PARAM NAME="RightMargin" VALUE="0.75">
						<PARAM NAME="TopMargin" VALUE="0.49">
						<PARAM NAME="BottomMargin" VALUE="0.49">
						<PARAM NAME="HeaderMargin" VALUE="0.49">
						<PARAM NAME="FooterMargin" VALUE="0.49">
						<PARAM NAME="FooterCaption" VALUE="- Page &amp;P -">
						<PARAM NAME="HeaderCaption" VALUE="DynamiCube">
						<PARAM NAME="HeaderJustification" VALUE="1">
						<PARAM NAME="FooterJustification" VALUE="1">
						<PARAM NAME="ColPageBreak" VALUE="0">
						<PARAM NAME="RowPageBreak" VALUE="0">
						<PARAM NAME="ColHeadingsOnEveryPage" VALUE="1">
						<PARAM NAME="RowHeadingsOnEveryPage" VALUE="0">
						<PARAM NAME="DCOptions" VALUE="0">
						<PARAM NAME="AutoDataRefresh" VALUE="-1">
						<PARAM NAME="PrinterColumnSpacing" VALUE="0.01">
						<PARAM NAME="DCConnectType" VALUE="0">
						<PARAM NAME="DCQueryTimeOut" VALUE="0">
						<PARAM NAME="SQLYearPart" VALUE='datepart("yyyy",<field>)'>
						<PARAM NAME="SQLQuarterPart" VALUE='datepart("q",<field>)'>
						<PARAM NAME="SQLMonthPart" VALUE='datepart("m",<field>)'>
						<PARAM NAME="SQLWeekPart" VALUE='datepart("ww",<field>)'>
						<PARAM NAME="BorderStyle" VALUE="1">
						<PARAM NAME="AllowSplitters" VALUE="-1">
						<PARAM NAME="QueryByPass" VALUE="0">
						<PARAM NAME="DataPath" VALUE="">
						<PARAM NAME="DataNotAvailableCaption" VALUE="">
						<PARAM NAME="PageFieldsVisible" VALUE="-1">
						<PARAM NAME="CubeBackColor" VALUE="14215660">
						<PARAM NAME="FooterBackColor" VALUE="-1">
						<PARAM NAME="FooterForeColor" VALUE="0">
						<PARAM NAME="HeaderBackColor" VALUE="-1">
						<PARAM NAME="HeaderForeColor" VALUE="0">
						<PARAM NAME="FilteredFieldBackColor" VALUE="-1">
						<PARAM NAME="FilteredFieldForeColor" VALUE="16777215">
						<PARAM NAME="MousePointer" VALUE="0">
						<PARAM NAME="LoadProgressNotifyDelay" VALUE="1000">
						<PARAM NAME="IncludeColorsInPrintout" VALUE="-1">
						<PARAM NAME="AutoDataRefresh" VALUE="-1">
					</OBJECT>		
					&nbsp;
				</td>
			</tr>
			</table>
			<script language="javascript">
				//esta funcion se agrega en el closeCalendar para que llame al Cerrar e Popup
				function fin_calendario()
				{	
					document.formula.submit();
				}
			</script>	
		<script language="JavaScript">
		var xmlDoc;
		function inicio(){
			xmlDoc = new ActiveXObject('Microsoft.XMLDOM');
			xmlDoc.onreadystatechange = checkState;
			xmlDoc.load("exportados/repimpaglobal<%=session("codusuario")%>.xml");
		}
		function checkState(){
			var returnString;
			var readyState;
			
			readyState = xmlDoc.readyState;
			switch(readyState)
			{
				case 1:
					returnString = '(1) Loading xml file...';
					break;
				case 2:
					returnString = '(2) Reading and parsing xml...';
					break;
				case 3:
					returnString = '(3) Xml has been read and parsed...';
					break;
				case 4:
					var errObject = xmlDoc.parseError;
					if(errObject.errorCode != 0) 
						returnString = '(4) Xml loaded with error(s): ' + errObject.Reason + 'Source Text:[' + errObject.SrcText + '].';
					else
					{
						returnString = '(4) Xml file was loaded successfully!';
						FillCube();
					}
					break;
			}		
			
			//divResults.innerHTML = divResults.innerHTML + 'readyState = ' + returnString + '<br/>';
		}		
		
			/*		cadenaXML=cadenaXML & " IDFactura='" & RS.Fields("IDFactura") & "'"
					cadenaXML=cadenaXML & " Fecha='" & RS.Fields("Fecha") & "'"
					cadenaXML=cadenaXML & " Cliente='" & RS.Fields("Cliente") & "'"
					cadenaXML=cadenaXML & " TipoDoc='" & RS.Fields("TipoDoc") & "'"
					cadenaXML=cadenaXML & " NumFactura='" & RS.Fields("NumFactura") & "'"
					cadenaXML=cadenaXML & " Articulo='" & RS.Fields("Articulo") & "'"
					cadenaXML=cadenaXML & " Unidad='" & RS.Fields("Unidad") & "'"
					cadenaXML=cadenaXML & " PrecioSoles='" & RS.Fields("PrecioSoles") & "'"
					cadenaXML=cadenaXML & " Cantidad='" & RS.Fields("Cantidad") & "'"
					cadenaXML=cadenaXML & " TotalSoles='" & RS.Fields("TotalSoles") & "'"
					cadenaXML=cadenaXML & " Estado='" & RS.Fields("Estado") & "'"
					cadenaXML=cadenaXML & " Credito='" & RS.Fields("Credito") & "'"
					cadenaXML=cadenaXML & " FechaVcto='" & RS.Fields("FechaVcto") & "'"			
			*/
					
		function FillCube(){
			var DCCT_UNBOUND=99;
			var DCRow=2;
			var DCColumn=1;
			var DCData=3;
			var DCLibre=4;
			var thisField;
			var DCube1 = document.frmMain.DCube1;
			
			DCube1.DCConnectType = DCCT_UNBOUND;
			DCube1.Fields.DeleteAll();
			thisField=DCube1.Fields.Add('Cliente', 'Cliente', DCRow);
			thisField=DCube1.Fields.Add('NumFactura', 'N° Documento', DCRow).GroupFooterVisible = false;
			thisField=DCube1.Fields.Add('Credito', 'Crédito', DCRow).GroupFooterVisible = false;						
			thisField=DCube1.Fields.Add('Vendedor', 'Vendedor', DCRow);
			thisField=DCube1.Fields.Add('Articulo', 'Artículo', DCRow);
			thisField=DCube1.Fields.Add('TipoDoc', 'Tipo', DCRow);
			thisField=DCube1.Fields.Add('Unidad', 'Unidad', DCRow).GroupFooterVisible = false;
			thisField=DCube1.Fields.Add('PrecioSoles', 'Precio S/.', DCRow).GroupFooterVisible = false;
			thisField=DCube1.Fields.Add('FechaAAAA','Año', DCColumn);
			//thisField=DCube1.Fields.Add('OrdenTrimestre', 'Trimestre', DCColumn);
			thisField=DCube1.Fields.Add('FechaMes', 'Mes', DCColumn);
			thisField=DCube1.Fields.Add('Cantidad', 'Cantidad', DCData);
			thisField=DCube1.Fields.Add('TotalSoles', 'Total S/.', DCData);
			
			DCube1.RefreshData();
			DCube1.AutoDataRefresh = true;
			DCube1.Fields(1).Orientation = DCLibre;		
			DCube1.Fields(2).Orientation = DCLibre;			
			DCube1.Fields(3).Orientation = DCLibre;	
			DCube1.Fields(5).Orientation = DCLibre;
			DCube1.Fields(6).Orientation = DCLibre;
			DCube1.Fields(7).Orientation = DCLibre;
		}
		</script>
		<script for='DCube1' event='FetchData'>
			var nodeList = xmlDoc.getElementsByTagName('z:row');
			if (nodeList.length <= 0){
				alert('Invalid xml file');
				return;
			}						
			var DCube1 = document.frmMain.DCube1;
			for ( var nodeIndex = 0; nodeIndex < nodeList.length ; nodeIndex++ ){
				var reporteNodeMap = nodeList[nodeIndex].attributes;
				var Cliente = reporteNodeMap.getNamedItem('Cliente').nodeValue;
				var NumFactura = reporteNodeMap.getNamedItem('NumFactura').nodeValue;
				var Credito = reporteNodeMap.getNamedItem('Credito').nodeValue;
				var Vendedor = reporteNodeMap.getNamedItem('Vendedor').nodeValue;																
				var Articulo = reporteNodeMap.getNamedItem('Articulo').nodeValue;
				var TipoDoc = reporteNodeMap.getNamedItem('TipoDoc').nodeValue;
				var Unidad = reporteNodeMap.getNamedItem('Unidad').nodeValue;
				var Fecha = reporteNodeMap.getNamedItem('Fecha').nodeValue;
				var yearOfReporte = parseInt(Fecha.substr(0,4),10);
				var monthOfReporte = parseInt(Fecha.substr(5,2),10);
				//var qtrOfReporte = parseInt((monthOfReporte-1)/3,10)+1;
				var PrecioSoles = reporteNodeMap.getNamedItem('PrecioSoles').nodeValue;
				var Cantidad = reporteNodeMap.getNamedItem('Cantidad').nodeValue;
				var TotalSoles = reporteNodeMap.getNamedItem('TotalSoles').nodeValue;
				//DCube1.AddRowEx(GetVBArray(Cliente, yearOfReporte, qtrOfReporte, monthOfReporte, TotalSoles))
				DCube1.AddRowEx(GetVBArray(Cliente,NumFactura,Credito,Vendedor,Articulo,TipoDoc,Unidad,PrecioSoles, yearOfReporte, monthOfReporte, Cantidad,TotalSoles))								
			}			
		</script>
		<script language="vbscript">
			Function GetVBArray(ClienteInfo, NumFacturaInfo,CreditoInfo,VendedorInfo, ArticuloInfo,TipoDocInfo,UnidadInfo, PrecioSolesInfo, AAAAInfo, MesInfo, CantidadInfo, TotalSolesInfo)
				Dim vbArray(11)
				//vbArray(2) = TrimInfo				
				vbArray(0) = ClienteInfo
				vbArray(1) = NumFacturaInfo
				vbArray(2) = CreditoInfo
				vbArray(3) = VendedorInfo
				vbArray(4) = ArticuloInfo
				vbArray(5) = TipoDocInfo
				vbArray(6) = UnidadInfo
				vbArray(7) = PrecioSolesInfo
				vbArray(8) = AAAAInfo
				if MesInfo=1 then vbArray(9) = "01 Ene"
				if MesInfo=2 then vbArray(9) = "02 Feb"
				if MesInfo=3 then vbArray(9) = "03 Mar"
				if MesInfo=4 then vbArray(9) = "04 Abr"
				if MesInfo=5 then vbArray(9) = "05 May"
				if MesInfo=6 then vbArray(9) = "06 Jun"
				if MesInfo=7 then vbArray(9) = "07 Jul"
				if MesInfo=8 then vbArray(9) = "08 Ago"
				if MesInfo=9 then vbArray(9) = "09 Set"
				if MesInfo=10 then vbArray(9) = "10 Oct"
				if MesInfo=11 then vbArray(9) = "11 Nov"
				if MesInfo=12 then vbArray(9) = "12 Dic"				
				vbArray(10) = CantidadInfo
				vbArray(11) = TotalSolesInfo
				GetVBArray = vbArray
			End Function
		</script>
		<script language="JavaScript">
		<!--
		function exportar()
		{
			var fileName = ""
			var includeTotals = true;
			var formatNumbers = true;
			var includeColors = true;
			var merge = true;
	
			 //formula.dialog.Filter = "Excel Document|*.xls";
			 //formula.dialog.filename="";
			 //formula.dialog.ShowSave();
			 //fileName =formula.dialog.filename;
			 fileName="C:/ExportCube.xls";
	
			if (fileName != "")
			{ 
			     formula.DCube1.ExportToExcel(fileName, includeTotals,formatNumbers, includeColors, merge);
			     window.alert("Se exportó correctamente el archivo " + fileName );
			}
		}	
		-->
		</script>
		<input type="hidden" name="actualizarlista" value="">
		<input type="hidden" name="expimp" value="">		
		<input type="hidden" name="pag" value="<%=pag%>">	
		<input type="hidden" name="estado" value="<%=estado%>">					
		</form>
		<script language="javascript">
			inicio();
		</script>		
		</body>
		</html>		
	<%
	else
	%>
	<script language="javascript">
	    alert("Ud. No tiene autorización para este proceso.");
	    window.open("userexpira.asp", "_top");
	</script>
	<%	
	end if
	desconectar
else
%>
<script language="javascript">
    alert("Tiempo Expirado");
    window.open("index.html", "_top");
</script>
<%
end if
%>




