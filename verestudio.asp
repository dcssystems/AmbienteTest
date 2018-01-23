<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repevalmensual.asp") then
	
		
	
		''fechadatos=obtener("fechadatos")
		''buscador=obtener("buscador")
		fechacierre=mid(obtener("fechacierre"),7,4) & mid(obtener("fechacierre"),4,2) & mid(obtener("fechacierre"),1,2)
		''fechagestion=obtener("fechagestion")
		estudio= obtener("estudio")
		''fechagestion=mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)
		expimp=obtener("expimp")
		if expimp="1" then
			sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaExportar' or descripcion='RutaWebExportar'"
			consultar sql,RS
			RS.Filter=" descripcion='RutaFisicaExportar'"
			RutaFisicaExportar=RS.Fields(1)
			RS.Filter=" descripcion='RutaWebExportar'"
			RutaWebExportar=RS.Fields(1)				
			RS.Filter=""
			RS.Close	
			tiempoexport=Now()
		end if
		''if not IsDate(fechagestionini) then
		''    fechagestionini=CStr(minfechagestion)
		''end if
		''if not IsDate(fechagestionfin) then
		''    fechagestionfin=CStr(maxfechagestion)
		''end if	
		
		sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUpload' or descripcion='RutaWebUpload'"
		consultar sql,RS3
		RS3.Filter=" descripcion='RutaFisicaUpload'"
		RutaFisicaUpload=RS3.Fields(1)
		RS3.Filter=" descripcion='RutaWebUpload'"
		RutaWebUpload=RS3.Fields(1)				
		RS3.Filter=""
		RS3.Close
		
		
		sql="select COUNT(A.CONTRATO) AS NUMCASOS, " &_
                 "COUNT(DISTINCT A.CODCENTRAL) AS NUMCLIENTES," &_
                 "SUM(A.JUD) AS IMPASIGNADO," &_
                 "SUM(CASE WHEN A.FECFASE IS NOT NULL THEN A.JUD ELSE 0 END) AS IMPFASE, " &_
                 "SUM(CASE WHEN YEAR(A.FECHAPASEAMORA)=YEAR(GETDATE()) THEN A.JUD ELSE 0 END) AS IMP2015, " &_
                  "SUM(CASE WHEN YEAR(A.FECHAPASEAMORA)=YEAR(GETDATE())-1 THEN A.JUD ELSE 0 END) AS IMP2014, " &_
                  "SUM(CASE WHEN YEAR(A.FECHAPASEAMORA)<=YEAR(GETDATE())-2 THEN A.JUD ELSE 0 END) AS IMP2013 " &_
                  "from FB.CentroCobranzas.dbo.PD_Casos_JUD A " &_
                 "where rtrim(ltrim(IsNull(A.ESTUDIO,'NO IDENTIFICADO')))= '" & estudio & "'  " &_
                 "AND A.FASIGNA IS NOT NULL " &_
                 "and A.FECHADATOS=(select max(fechacierre) from FB.CentroCobranzas.dbo.PD_FECHASCIERRE WHERE FECHACIERRE< '" & fechacierre & "' ) "	
		
		consultar sql,RS
		if not RS.EOF then
			NUMCASOS=RS.Fields("NUMCASOS")
			NUMCLIENTES=RS.Fields("NUMCLIENTES")
			IMPASIGNADO=RS.Fields("IMPASIGNADO")
			IMPFASE=RS.Fields("IMPFASE")
			IMP2015=RS.Fields("IMP2015")
			IMP2014=RS.Fields("IMP2014")
			IMP2013=RS.Fields("IMP2013")
		end if
		RS.Close
		

		
		sql="SELECT	" &_
	            "SUM(CASE WHEN  MODALIDAD IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as VentaTerceros, " &_
	            "SUM(CASE WHEN  MODALIDAD IN ('D') THEN soles + dolares*tipocambio ELSE 0 END) as Efectivo, " &_ 
	            "SUM(CASE WHEN  MODALIDAD IN ('T') THEN soles + dolares*tipocambio ELSE 0 END) as Transaccion,  " &_
	            "SUM(CASE WHEN  MODALIDAD IN ('R') THEN soles + dolares*tipocambio ELSE 0 END) as Refinanciado, " &_ 
	            "SUM(CASE WHEN  MODALIDAD IN ('M') THEN soles + dolares*tipocambio ELSE 0 END) as Mueble,  " &_
	            "SUM(CASE WHEN  MODALIDAD IN ('N') THEN soles + dolares*tipocambio ELSE 0 END) as Inmueble, " &_
	            "SUM(CASE WHEN  MODALIDAD IN ('G') THEN soles + dolares*tipocambio ELSE 0 END) as Dacion, " &_
	            "SUM(CASE WHEN  MODALIDAD IN ('B') THEN soles + dolares*tipocambio ELSE 0 END) as Leasing, " &_
	            "SUM(CASE WHEN  PRODUCTO IN ('Leasing') THEN soles + dolares*tipocambio ELSE 0 END) as LeasingPYME, " &_
	            "SUM(CASE WHEN  PRODUCTO IN ('PRÉSTAMOS COMERCIALES') THEN soles + dolares*tipocambio ELSE 0 END) as PComercialPYME, " &_
	            "SUM(CASE WHEN  PRODUCTO IN ('Tarjeta Comerciales') THEN soles + dolares*tipocambio ELSE 0 END) as TComercialPYME, " &_
	            "SUM(CASE WHEN  PRODUCTO IN ('HIPOTECARIO') THEN soles + dolares*tipocambio ELSE 0 END) as HipotecarioPART, " &_
	            "SUM(CASE WHEN  PRODUCTO IN ('TARJETA CONSUMO') THEN soles + dolares*tipocambio ELSE 0 END) as TConsumoPART, " &_
	            "SUM(CASE WHEN  PRODUCTO IN ('PRÉSTAMOS CONSUMO') THEN soles + dolares*tipocambio ELSE 0 END) as PConsumoPART, " &_
	            "SUM(CASE WHEN  PRODUCTO IN ('PRÉSTAMO VEHICULAR') THEN soles + dolares*tipocambio ELSE 0 END) as PVehicularPART " &_
            "FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A  " &_
                "LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO		= A.CONT_SAE " &_  
                "LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO		= C.SIGLA " &_ 
            "WHERE	IsNull(C.nomb_estu,'NO IDENTIFICADO')='" & estudio & "' and YEAR(A.FECHA_SALIDAS)=year('" & fechacierre & "') " &_ 
            "AND A.FECHA_SALIDAS<='" & fechacierre & "' " &_
            "AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('PARTICULARES', 'PYME') " &_ 
            "AND ltrim(rtrim(A.entidad))			in ('BC','35') " &_	
            "AND	A.MODALIDAD	in ('B','D','G','M','N','R','T','W')"

		''Response.Write sql
		consultar sql,RS2
		if not RS2.EOF then
			VentaTerceros=RS2.Fields("VentaTerceros")
			Efectivo=RS2.Fields("Efectivo")
			Transaccion=RS2.Fields("Transaccion")
			Refinanciado=RS2.Fields("Refinanciado")
			Mueble=RS2.Fields("Mueble")
			Inmueble=RS2.Fields("Inmueble")
			Dacion=RS2.Fields("Dacion")
			Leasing=RS2.Fields("Leasing")
			LeasingPYME=RS2.Fields("LeasingPYME")
	        PComercialPYME=RS2.Fields("PComercialPYME")
	        TComercialPYME=RS2.Fields("TComercialPYME")
	        HipotecarioPART=RS2.Fields("HipotecarioPART")
	        TConsumoPART=RS2.Fields("TConsumoPART")
	        PConsumoPART=RS2.Fields("PConsumoPART")
	        PVehicularPART=RS2.Fields("PVehicularPART")

		end if
		RS2.Close
		

		
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<title>Ver Estudio</title>
		<script language=javascript src="scripts/TablaDinamica.js"></script>
		<link rel="stylesheet" type="text/css" media="all" href="scripts/calendar-blue2.css" title="blue" />
		<script type="text/javascript" src="scripts/calendar.js"></script>
		<script type="text/javascript" src="scripts/calendar-es.js"></script>
		<script type="text/javascript">
		    function selected(cal, date) {
		        cal.sel.value = date;
		        if (cal.dateClicked && (cal.sel.id == "sel1" || cal.sel.id == "sel3"))
		            cal.callCloseHandler();
		    }
		    function closeHandler(cal) {
		        cal.hide();
		        _dynarch_popupCalendar = null;
		    }
		    function showCalendar(id, format, showsTime, showsOtherMonths) {
		        var el = document.getElementById(id);
		        if (_dynarch_popupCalendar != null) {
		            _dynarch_popupCalendar.hide();
		        } else {
		            var cal = new Calendar(1, null, selected, closeHandler);
		            if (typeof showsTime == "string") {
		                cal.showsTime = true;
		                cal.time24 = (showsTime == "24");
		            }
		            if (showsOtherMonths) {
		                cal.showsOtherMonths = true;
		            }
		            _dynarch_popupCalendar = cal;
		            cal.setRange(1900, 2070);
		            cal.create();
		        }
		        _dynarch_popupCalendar.setDateFormat(format);
		        _dynarch_popupCalendar.parseDate(el.value);
		        _dynarch_popupCalendar.sel = el;

		        _dynarch_popupCalendar.showAtElement(el.nextSibling, "Br");
		        return false;
		    }
		</script>
		<script language=javascript>
		var nuevodir;
		var nuevotelf;
		var nuevoemail;
		var nuevogestion;
		var ventanagestion;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(c0)
		{
			ventanagestion=global_popup_IWTSystem(ventanagestion,"verseguimiento.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado3.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codrespgestion=" + c0,"NewGestion","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=300,width=" + (screen.width/2 + 100) + ",left=" + (screen.width/4 - 50) + ",resizable=yes");
		}		
		function agregardir()
		{
			nuevodir=global_popup_IWTSystem(nuevodir,"adicionardireccion.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado3.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newdir","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 + 250) + ",left=" + (screen.width/4) + ",resizable=yes");
		}
		function eliminardir(elimdir)
		{
			if(confirm("¿Está Seguro de Eliminar la Dirección del Cliente?"))
			{						
				document.formula.direccioneliminar.value=elimdir;
				document.formula.submit();
			}			
		}		
		function agregartelf()
		{
			nuevotelf=global_popup_IWTSystem(nuevotelf,"adicionartelefono.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newtelf","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
		}
		function eliminartelf(elimtelf)
		{
			if(confirm("¿Está Seguro de Eliminar el Teléfono del Cliente?"))
			{						
				document.formula.telefonoeliminar.value=elimtelf;
				document.formula.submit();
			}			
		}				
		function agregaremail()
		{
			nuevoemail=global_popup_IWTSystem(nuevoemail,"adicionaremail.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newemail","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
		}		
		function eliminaremail(elimmail)
		{
			if(confirm("¿Está Seguro de Eliminar el E-Mail del Cliente?"))
			{						
				document.formula.emaileliminar.value=elimmail;
				document.formula.submit();
			}			
		}		
		function agregargestion()
		{
			nuevogestion=global_popup_IWTSystem(nuevogestion,"adicionarseguimiento.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","NewGestion","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=300,width=" + (screen.width/2 + 100) + ",left=" + (screen.width/4 - 50) + ",resizable=yes");
		}	
		function actualizar()
		{
			document.formula.actualizarlista.value=1;
			document.formula.submit();
		}	
		function exportar()
		{
			document.formula.expimp.value=1;
			document.formula.submit();
		}
		function imprimir()
		{
			window.open("impusuarios.asp","ImpUsuarios","scrollbars=yes,scrolling=yes,top=0,height=200,width=200,left=0,resizable=yes");
		}					
		function buscar()
		{
			document.formula.pag.value=1;
			document.formula.submit();
		}		
		function filtrar()
		{
			if (filtrardatos[0]==1)
			{
				filtrardatos[0]=0;
				dibujarTabla(0);
			}
			else
			{
				filtrardatos[0]=1;
				dibujarTabla(0);
			}
		}				
		function mostrarpag(pagina)
		{
			document.formula.pag.value=pagina;
			document.formula.submit();
		}
		</script>
		<style>
		A {
			FONT-SIZE: 12px; COLOR: #00529B; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
		}
		A:visited {
			TEXT-DECORATION: none; COLOR: #00529B;
		}
		A:hover {
			COLOR: #00529B; FONT-FACE:"Arial"; TEXT-DECORATION: none
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
			color:#FFFFFF;
			background: #007DC5;
			font-size:12px;
			font-family:Arial;
			cursor:hand;
		}
		
		</style>
		</head>
		
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula method=post>
			<table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td align="center" height="22"><font size=2 face=Arial color="#FFFFFF"><b>Módulo Detalles de Estudio</b></font></td>
			</tr>
			</table>
			<table width=100% cellpadding=2 cellspacing=2 border=0>

			<tr>
				 <td bgcolor="#BEE8FB" width="25%" ><font size=2 face=Arial color=#00529B>&nbsp;N° Casos:</font></td>
				 <td bgcolor="#E9F8FE" width="25%" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=NUMCASOS%></b></font></td>
				 <td bgcolor="#BEE8FB" width="25%" colspan="2"><font size=2 face=Arial color=#00529B>&nbsp;N° Clientes:</font></td>
				 <td bgcolor="#E9F8FE" width="25%" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=NUMCLIENTES%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Importe Asignado:</font></td>
				 <td bgcolor="#E9F8FE" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(IMPASIGNADO,2)%></b></font></td>
				 <td bgcolor="#BEE8FB" colspan="2"><font size=2 face=Arial color=#00529B>&nbsp;Importe con Fase:</font></td>
				 <td bgcolor="#E9F8FE" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(impfase,2)%></b></font></td>
			</tr>
            <tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Importe 2015:</font></td>
				 <td bgcolor="#E9F8FE" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(imp2015,2)%></b></font></td>
				 <td bgcolor="#BEE8FB" colspan="2"><font size=2 face=Arial color=#00529B>&nbsp;Importe 2014:</font></td>
				 <td bgcolor="#E9F8FE" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(imp2014,2)%></b></font></td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Importe 2013:</font></td>
				 <td bgcolor="#E9F8FE" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(imp2013,2)%></b></font></td>
				 <td bgcolor="#BEE8FB" colspan="2"><font size=2 face=Arial color=#00529B>&nbsp;% Fase:</font></td>
				 <td bgcolor="#E9F8FE" colspan="2" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=formatnumber(impfase*100/impasignado,2)%>%</b></font></td>
			</tr>
			
			<tr>
				 <td bgcolor="#BEE8FB" valign=top rowspan="8"><font size=2 face=Arial color=#00529B >&nbsp;Salidas por Modalidad:</font></td>
	            <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Venta a Terceros:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(VentaTerceros,2)%></b></font></td>
		 
			</tr>
			<tr>
				<td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Efectivo:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(Efectivo,2)%></b></font></td>
			</tr>
			<tr>
	            <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Transacción:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(Transaccion,2)%></b></font></td>
	        </tr>
			<tr>
	            <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Refinanciado:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(Refinanciado,2)%></b></font></td>
			</tr>
			<tr>
                <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Mueble:</font></td>
                <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(Mueble,2)%></b></font></td>
            </tr>
			<tr>   
                <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Inmueble:</font></td>
                <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(Inmueble,2)%></b></font></td>
            </tr>
			<tr>
                <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Dación:</font></td>
                <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(Dacion,2)%></b></font></td>
            </tr>
			<tr>
                <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Leasing:</font></td>
                <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(Leasing,2)%></b></font></td>
     			 
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB" valign=top rowspan="7"><font size=2 face=Arial color=#00529B >&nbsp;Salidas por Producto</font></td>
	            <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Leasing:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(LeasingPYME,2)%></b></font></td>
		 
			</tr>
			<tr>
				<td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Préstamo Comercial:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(PComercialPYME,2)%></b></font></td>
			</tr>
			<tr>
	            <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Tarjeta Comercial:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(TComercialPYME,2)%></b></font></td>
	        </tr>
			<tr>
	            <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Hipotecario:</font></td>
	            <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(HipotecarioPART,2)%></b></font></td>
			</tr>
			<tr>
                <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Tarjeta Consumo:</font></td>
                <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(TConsumoPART,2)%></b></font></td>
            </tr>
			<tr>   
                <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Préstamo Consumo:</font></td>
                <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(PConsumoPART,2)%></b></font></td>
            </tr>
			<tr>
                <td bgcolor="#BEE8FB" colspan="3"><font size=2 face=Arial color=#00529B>Préstamo Vehicular:</font></td>
                <td bgcolor="#E9F8FE" colspan="3" align="right"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=FormatNumber(PVehicularPART,2)%></b></font></td>
            </tr>
            

            <% 
            	sql="SELECT distinct IsNull(D.Especialista_jud_part,'NO IDENTIFICADO') AS ESPECIALISTA, " &_
					"SUM(soles + dolares*tipocambio) AS SALIDASPART " &_
					"FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A  " &_
					    "LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO		= A.CONT_SAE " &_  
					    "LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO		= C.SIGLA  " &_
					    "LEFT OUTER JOIN FB.CentroCobranzas.dbo.PD_CENTROS_ESPECIALISTA D ON A.CODOFICINA=D.COD_OFICINA " &_
					"WHERE	IsNull(C.nomb_estu,'NO IDENTIFICADO')='" & estudio &"' and YEAR(A.FECHA_SALIDAS)=year('" & fechacierre & "')  " &_
					"AND A.FECHA_SALIDAS<='" & fechacierre & "' " &_
					"AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('PARTICULARES')  " &_
					"AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
					"AND	A.MODALIDAD	in ('B','D','G','M','N','R','T','W')  " &_
					"GROUP BY IsNull(D.Especialista_jud_part,'NO IDENTIFICADO') order by SALIDASPART DESC " 
					contador=0
                    
					consultar sql,RS
					
					%>
					<tr>
					    <td bgcolor="#BEE8FB" valign=top>
					        <table width=100% cellpadding=0 cellspacing=0 border=0>
					        <tr>
						        <td><font size=2 face=Arial color=#00529B>&nbsp;Salidas por Especialista Particular:</font></td>
					        </tr>
					        </table>					
				         </td>
				         <td colspan="3">
				         <table id="tablaespart" width=100% cellpadding=0 cellspacing=1 border=0>
				    
					<%
					        Do While Not  RS.EOF
						        %>
							        <tr>
                                        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=RS.Fields("Especialista")%></b></font></td>
						                <td bgcolor="#E9F8FE"  align="right"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=FormatNumber(RS.Fields("salidaspart"),2)%></font></td>
					                </tr>
						        <%
        	
						        RS.MoveNext
					        loop
					        RS.Close
					%>
			            </table>
			        </td>
		        </tr>

            

            <% 
            	sql="SELECT distinct IsNull(D.Especialista_jud_pyme,'NO IDENTIFICADO') AS ESPECIALISTA, " &_
					"SUM(soles + dolares*tipocambio) AS SALIDASPYME " &_
					"FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A  " &_
					    "LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO		= A.CONT_SAE " &_  
					    "LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO		= C.SIGLA  " &_
					    "LEFT OUTER JOIN FB.CentroCobranzas.dbo.PD_CENTROS_ESPECIALISTA D ON A.CODOFICINA=D.COD_OFICINA " &_
					"WHERE	IsNull(C.nomb_estu,'NO IDENTIFICADO')='" & estudio &"' and YEAR(A.FECHA_SALIDAS)=year('" & fechacierre & "')  " &_
					"AND A.FECHA_SALIDAS<='" & fechacierre & "' " &_
					"AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('PARTICULARES')  " &_
					"AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
					"AND	A.MODALIDAD	in ('B','D','G','M','N','R','T','W')  " &_
					"GROUP BY IsNull(D.Especialista_jud_pyme,'NO IDENTIFICADO')  order by SALIDASPYME DESC " 
					
					contador=0

					consultar sql,RS
					%>
					<tr>
					    <td bgcolor="#BEE8FB" valign=top>
					        <table width=100% cellpadding=0 cellspacing=0 border=0>
					        <tr>
						        <td><font size=2 face=Arial color=#00529B>&nbsp;Salidas por Especialista Pyme:</font></td>
					        </tr>
					        </table>					
				         </td>
				         <td colspan="3">
				         <table id="Table1" width=100% cellpadding=0 cellspacing=1 border=0>
				    
					<%
					        Do While Not  RS.EOF
						        %>
							        <tr>
                                        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=RS.Fields("Especialista")%></b></font></td>
						                <td bgcolor="#E9F8FE" align="right"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=FormatNumber(RS.Fields("salidaspyme"),2)%></font></td>
					                </tr>
						        <%
        	
						        RS.MoveNext
					        loop
					        RS.Close
					%>
			            </table>
			        </td>
		        </tr>
           


					
			<!--<tr>
			
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Garantías:</font></td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>-->
						<script language=javascript>
						    function visualizargarantia() {
						        var filas = document.getElementById('tablagarantias').rows.length;
						        //if (document.getElementById('imagengarantia').src.length==document.getElementById('imagengarantia').src.replace('mostrar','').length)
						        if (document.getElementById('imagengarantia').title == "Mostrar") {
						            document.getElementById('imagengarantia').title = "Ocultar";
						            document.getElementById('imagengarantia').alt = "Ocultar";
						            document.getElementById('imagengarantia').src = "imagenes/ocultar.png";
						            for (i = 1; i < filas; i++) {
						                document.getElementById('tablagarantias').rows[i].style.display = '';
						            }
						        }
						        else {
						            document.getElementById('imagengarantia').title = "Mostrar";
						            document.getElementById('imagengarantia').alt = "Mostrar";
						            document.getElementById('imagengarantia').src = "imagenes/mostrar.png";
						            for (i = 1; i < filas; i++) {
						                document.getElementById('tablagarantias').rows[i].style.display = 'none';
						            }
						        }
						    }
						</script>
				 		<!--<table id="tablagarantias" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizargarantia();"><img id="imagengarantia" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Garantía Total:</b></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=formatnumber(garantia,2)%></b></font></td>
						</tr>
						<% if garantrr<>0 then %>			 		
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td  width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía RR:</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantrr,2)%></font></td>

							</tr>
						<% end if %>
						<% if garantpref<>0 then %>	
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía Preferida (Hipotecaria):</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantpref,2)%></font></td>
							</tr>
						<% end if %>
						
						<% if garantauto<>0 then %>
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía Autoliquidable:</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantauto,2)%></font></td>
							</tr>
						<% end if %>
						<% if garantcontra<>0 then %>
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía Contraparte:</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantcontra,2)%></font></td>
							</tr>
						<% end if %>
						</table>
				 </td>
			</tr>-->
			
		<%''end if%>
			<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
			<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
			<input type="hidden" name="contrato" value="<%=contrato%>">
		    <input type="hidden" name="expimp" value="">
		    <input type="hidden" name="pag" value="<%=pag%>">	
    	</form>
			
		
		<!--<script type="text/javascript">
		    initTriStateCheckBox('tristateBox1', 'tristateBox1State', true);
		</script>-->
		<%if contador > 0 then%>
		<script language="javascript">
		    inicio();
		</script>
		<%end if%>
		<!--cargando--><script language=javascript>		                   document.getElementById("imgloading").style.display = "none";</script>							
		</body>
		</html>		
		<%
		''Codigo exp excel
		''Si se pide exportar a excel
				if expimp="1" then
					''Para Exportar a Excel
					''Paso Cero eliminar exportación anterior
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & "," & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
					
					''Primero Cabecera en temp1_(user).txt
					consulta_exp="select 'Fecha de Gestión','Contrato','Agencia','Gestion','Comentario','F.Promesa','Direccion/Telefono','Adjunto'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select convert(varchar,A.fhgestionado,103) + ' ' + convert(varchar,A.fhgestionado,108),char(39) + A.contrato,C.RazonSocial,B.Descripcion, replace(replace(A.comentario,char(10),''),char(13),'') as comentario,ISNULL(convert(varchar,A.fechapromesa,103),''),CASE WHEN ltrim(rtrim(A.fono))<>'' THEN rtrim(ltrim(A.tipofono)) + '-' + rtrim(ltrim(A.prefijo)) + '-' + rtrim(ltrim(A.fono)) + (CASE WHEN LEN(RTRIM(ltrim(A.extension)))> 0 and A.extension<>'0000' THEN '-'+ RTRIM(LTRIM(A.extension)) ELSE '' END) ELSE rtrim(ltrim(A.direccion)) + ' - ' + rtrim(ltrim(UBI.distrito)) + ' - ' + rtrim(ltrim(UBI.provincia)) + ' - ' + rtrim(ltrim(UBI.departamento)) END,CASE WHEN A.ficherogestion IS NOT NULL THEN 'Sí' ELSE 'No' END " & _
								 "from CobranzaCM.dbo.RespuestaGestion A inner join CobranzaCM.dbo.Gestion B on A.codgestion=B.codgestion left outer join CobranzaCM.dbo.Agencia C on A.codagencia=C.codagencia left outer join CobranzaCM.dbo.Ubigeo UBI on UBI.coddpto=A.coddpto and UBI.codprov=A.codprov and UBI.coddist=A.coddist " & filtrobuscador & " order by A.fhgestionado desc"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					''response.Write sql
					conn.execute sql
					
					''Tercero borrar UserExport*.xls
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql						
					
					''Cuarto Uno los 2 archivos en temp*.txt
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''copy " & chr(34) & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt" & chr(34) & " + " & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & " /b''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
					
					''Quinto Elimino los 2 archivos en temp*.txt
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt" & chr(34) & "," & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql										
				%>
					<script language="javascript">
						window.open("<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>","_self");
					</script>
				<%					
				end if			
		%>		
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
    window.open("index.html", "sistema");
    window.close();
</script>
<%
end if
%>



