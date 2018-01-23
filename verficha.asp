<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admvisitas.asp") then
		fdcodcen=obtener("fdcodcen")
		
		sql="select *,(select descripcion from Clasificacion where codclasificacion=A.clasificacion) as clasifica from Cliente_FichaVisita A where convert(varchar,A.fechadatos,112) + A.codigocentral='" & fdcodcen & "'"
		''Response.Write sql
		consultar sql,RS1	
		if not RS1.EOF then
			fechadatos=RS1.fields("fechadatos")
			codigocentral=RS1.fields("codigocentral")
			nombres=RS1.fields("nombres")
			tipodocumento=RS1.fields("tipodocumento")
			numdocumento=RS1.fields("numdocumento")
			marca=RS1.fields("marca")
			segmento_riesgo=RS1.fields("segmento_riesgo")
			banca=RS1.fields("banca")
			codterritorio=RS1.fields("codterritorio")
			territorio=RS1.fields("territorio")
			codoficina=RS1.fields("codoficina")
			oficina=RS1.fields("oficina")
			maxdias=RS1.fields("maxdias")
			if MaxDias mod 30 > 0 then
			    tramo=Int(MaxDias/30) + 1
			else
			    tramo=Int(MaxDias/30)
			end if
			clasificacion=RS1.fields("clasificacion")
			clasifica=RS1.fields("clasifica")
			direccion=RS1.fields("direccion")
			referencia1=RS1.fields("referencia1")
			referencia2=RS1.fields("referencia2")
			distrito=RS1.fields("distrito")
			departamento=RS1.fields("departamento")
			provincia=RS1.fields("provincia")
			ciiu=RS1.fields("ciiu")
			descciiu=RS1.fields("descciiu")
		end if
		RS1.Close
		
		
		fechavisita=obtener("fechavisita")
		horavisita=iif(obtener("horavisita")="","00",obtener("horavisita"))
		minutovisita=iif(obtener("minutovisita")="","00",obtener("minutovisita"))
		tipocontacto=obtener("tipocontacto")
		dirsistema=obtener("dirsistema")
		coddireccionnueva=obtener("coddireccionnueva")
		situaciontel1=obtener("situaciontel1")
		situaciontel2=obtener("situaciontel2")
		situaciontel3=obtener("situaciontel3")
		situaciontel4=obtener("situaciontel4")
		situaciontel5=obtener("situaciontel5")
    	fechapdp=obtener("fechapdp")
		actprinc=obtener("actprinc")		
		actividadotros=obtener("actividadotros")
		actividadadicional=obtener("actividadadicional")
		enactividad=obtener("enactividad")
		localalquilado=obtener("localalquilado")
		localfinanciado=obtener("localfinanciado")
		facturacionprincipal=obtener("facturacionprincipal")
		facturacionadicional=obtener("facturacionadicional")
		puertacalle=obtener("puertacalle")
		ofadministrativa=obtener("ofadministrativa")
		casanegocio=obtener("casanegocio")
		laborartesanal=obtener("laborartesanal")
		existencias=obtener("existencias")
		nropersonas=obtener("nropersonas")
		motivoatraso=obtener("motivoatraso")
		otrascausasatraso=obtener("otrascausasatraso")
		afrontapago=obtener("afrontapago")
		otrosafronta=obtener("otrosafronta")
		cuestionacobro=obtener("cuestionacobro")
		nocontacto=obtener("nocontacto")
		comentario=obtener("comentario")
		estado=obtener("estado")
					
        sql="select * from ClienteDiario A where convert(varchar,A.fechadatos,112) + A.codigocentral='" & fdcodcen & "'"		
		consultar sql,RS1	
		if not RS1.EOF then
			tipofono1=RS1.Fields("tipofono1")
			prefijo1=RS1.Fields("prefijo1")
			fono1=RS1.Fields("fono1")
			extension1=RS1.Fields("extension1")
			tipofono2=RS1.Fields("tipofono2")
			prefijo2=RS1.Fields("prefijo2")
			fono2=RS1.Fields("fono2")
			extension2=RS1.Fields("extension2")
			tipofono3=RS1.Fields("tipofono3")
			prefijo3=RS1.Fields("prefijo3")
			fono3=RS1.Fields("fono3")
			extension3=RS1.Fields("extension3")
			tipofono4=RS1.Fields("tipofono4")
			prefijo4=RS1.Fields("prefijo4")
			fono4=RS1.Fields("fono4")
			extension4=RS1.Fields("extension4")
			tipofono5=RS1.Fields("tipofono5")
			prefijo5=RS1.Fields("prefijo5")
			fono5=RS1.Fields("fono5")
			extension5=RS1.Fields("extension5")
			email=RS1.Fields("email")
		end if		
		RS1.Close
		''esta viariable viene si agregue direccion
	    agreguedir=obtener("agreguedir")	
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<title>Ver Ficha</title>
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
		function inicio()
		{
		dibujarTabla(0);
		}
		function agregardir()
		{
			nuevodir=global_popup_IWTSystem(nuevodir,"progvacio.html","Newdir","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 + 250) + ",left=" + (screen.width/4) + ",resizable=yes");
			document.formula.vistapadre1.value=window.name;
			document.formula.paginapadre1.value="verficha.asp";
            document.formula.action="adicionardireccion.asp";
            document.formula.target="Newdir";
            document.formula.submit();
            document.formula.action="verficha.asp";
            document.formula.target="_self";            
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
			nuevotelf=global_popup_IWTSystem(nuevotelf,"progvacio.html","Newtelf","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
			document.formula.vistapadre1.value=window.name;
			document.formula.paginapadre1.value="verficha.asp";
            document.formula.action="adicionartelefono.asp";
            document.formula.target="Newtelf";
            document.formula.submit();
            document.formula.action="verficha.asp";
            document.formula.target="_self";     			
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
				<td align="center" height="22"><font size=2 face=Arial color="#FFFFFF"><b>INFORME DE GESTIÓN PRESENCIAL</b></font></td>
			</tr>
			</table>
			<table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Fecha de Datos:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=fechadatos%></b></font></td>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Estado:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;PENDIENTE</b></font></td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;N° Documento:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=tipodocumento%> - <%=numdocumento%></b></font></td>
                 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Banca:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=banca%></b></font></td>                 
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Cliente:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=codigocentral + " - " + nombres%></b></font></td>			
                 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Territorio:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=codterritorio + " - " + territorio%></b></font></td>				 				 
			</tr>				
			<tr>
                 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Giro:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=descciiu%></b></font></td>				 			
                 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Oficina:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=codoficina + " - " + oficina%></b></font></td>				 
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Días Atraso / Tramo:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=MaxDias%> / TRAMO <%=tramo%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Clasificación:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=clasificacion & " - " & clasifica%></b></font></td>				 
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Dirección:</font></td>
				 <td bgcolor="#E9F8FE" colspan=3><font size=2 face=Arial color=#00529B><b>&nbsp;<%=direccion & referencia1 & referencia2%> / <%=distrito%> / <%=provincia%> / <%=departamento%></b></font></td>
			</tr>															
			</table>
			<table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;Obligaciones del Cliente</b></font></td>
			</tr>
			</table>		
			<script language=javascript>
			function visualizarcuotas(numcont)
			{
				var filas = document.getElementById('tablacontratos').rows.length;
				//if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
				if (document.getElementById('imagencuota' + numcont).title=="Mostrar")
				{
					document.getElementById('imagencuota' + numcont).title="Ocultar";
					document.getElementById('imagencuota' + numcont).alt="Ocultar";
					document.getElementById('imagencuota' + numcont).src="imagenes/ocultar.png";
					for (i = 1; i < filas; i++)
					{
						if(document.getElementById('tablacontratos').rows[i].id==numcont) document.getElementById('tablacontratos').rows[i].style.display = '';
					}
				}
				else
				{
					document.getElementById('imagencuota' + numcont).title="Mostrar";
					document.getElementById('imagencuota' + numcont).alt="Mostrar";
					document.getElementById('imagencuota' + numcont).src="imagenes/mostrar.png";
					for (i = 1; i < filas; i++)
					{
						if(document.getElementById('tablacontratos').rows[i].id==numcont) document.getElementById('tablacontratos').rows[i].style.display = 'none';
					}							
				}
			}
			</script>				
			<table width=100% id="tablacontratos" cellpadding=1 cellspacing=1 border=0>
			<tr bgcolor="#007DC5">
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Cuotas</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>N° Contrato</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Producto</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>DA</b></font></td>						
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Mon</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Total Impagado</b></font></td>						
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Vencido</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Prov.Const.</b></font></td>
			</tr>
			<%
			sql="select *,IsNull((select top 1 fixing from TipoCambio where divisa=A.divisa and tipo='S' and fechadatos<=A.fechadatos),1) as TipoCambio from Contrato_FichaVisita A where convert(varchar,A.fechadatos,112) + A.codigocentral='" & fdcodcen & "' order by diasvencido desc"
			consultar sql,RS1
			Do While not RS1.EOF
				sql="select * from CuotaDiario where contrato='" & RS1.Fields("Contrato") & "' and fechadatos='" & year(RS1.Fields("FechaDatos")) & right("00" & month(RS1.Fields("FechaDatos")),2) & right("00" & day(RS1.Fields("FechaDatos")),2) & "' order by fechavencimiento,divisa"
				consultar sql,RS2
				nrocuotas=0
				fechavencimiento=""
				Do While Not RS2.EOF
					if fechavencimiento<>RS2.Fields("FechaVencimiento") then
						fechavencimiento=RS2.Fields("FechaVencimiento")
						nrocuotas=nrocuotas + 1
					end if
				RS2.MoveNext
				Loop

                ''sql="select "
    
				'',(select count(distinct fechavencimiento) from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos) as NumCuotas,(select top 1 divisa from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos and divisa<>A.divisa) as DivisaDif 
			%>
			<tr bgcolor="#E9F8FE">
                    <td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=nrocuotas%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Contrato")%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("DescProducto")%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("DiasVencido")%></font></td>		
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("divisa")%></font></td>
					<td valign="top" align="right"><font size=2 face=Arial color=#00529B><%=FormatNumber(RS1.Fields("importe"),2)%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Vencido")%></font></td>
					<td valign="top" align="right"><font size=2 face=Arial color=#00529B><%=FormatNumber(RS1.Fields("ProvConst")/RS1.Fields("TipoCambio"),2)%></font></td>
			</tr>
			<%
			RS2.Close
			RS1.MoveNext
			Loop
			RS1.Close
			%>
			</table>
			<table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;Historial de Gestiones</b></font></td>
			</tr>
			</table>	
            <table width=100% id="idgestiones" cellpadding=1 cellspacing=1 border=0>
			<tr bgcolor="#007DC5">
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Fecha</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Hora</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Respuesta</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Comentario</b></font></td>						
			</tr>
			<%
			sql="select * from Gestiones_FichaVisita A where convert(varchar,A.fechadatos,112) + A.codigocentral='" & fdcodcen & "' order by Fproceso desc"
			consultar sql,RS1
			Do While not RS1.EOF
			%>
			<tr bgcolor="#E9F8FE">
                    <td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("FProceso")%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Hora")%></font></td>
					<td valign="top"><font size=2 face=Arial color=#00529B><%=RS1.Fields("trespuesta")%></font></td>
					<td valign="top"><font size=2 face=Arial color=#00529B><%=RS1.Fields("observaciones")%></font></td>		
			</tr>
			<%
			RS1.MoveNext
			Loop
			RS1.Close
			%>
			</table>	
			
			<table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;Datos de la Visita</b></font></td>
			</tr>
			</table>			

			<table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Fecha de Visita:</font></td>
				 <td bgcolor="#E9F8FE"><input name="fechavisita" type=text maxlength=10 id="sel3" readonly value="<%if IsDate(fechavisita) then%><%=fechavisita%><%else%><%=date()%><%end if%>" style="font-size: x-small; width: 80px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel3', '%d/%m/%Y');"><font size=2 face=Arial color="#00529B">&nbsp;Hora:&nbsp;</font>
				                        <select name="horavisita" style="text-align: center;font-size: x-small; width: 40px;">
				                        <%for i=0 to 23%>
				                        <option value="<%=right("00" & i,2)%>" <%if horavisita=right("00" & i,2) then%> selected<%end if%>><%=right("00" & i,2)%></option>
				                        <%next%>
				                        </select>
				                        <font size=2 face=Arial color="#00529B"><b>&nbsp;:&nbsp;</b></font>
				                        <select name="minutovisita" style="text-align: center;font-size: x-small; width: 40px;">
				                        <%for i=0 to 59%>
				                        <option value="<%=right("00" & i,2)%>" <%if minutovisita=right("00" & i,2) then%> selected<%end if%>><%=right("00" & i,2)%></option>
				                        <%next%>
				                        </select>
                 </td>
			</tr>	
			<tr>
                                <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Tipo de Contacto:</font></td>
                                <td bgcolor="#E9F8FE">
                                    <select name="tipocontacto" style="font-size: x-small; width: 250px;">
                                    <option value="">Seleccione Tipo de Contacto</option>
                                               <%
                                               sql="select codtipocontacto,descripcion from tipocontacto where activo=1 order by codtipocontacto"
                                               consultar sql,RS
                                               Do While Not  RS.EOF
                                               %>
                                                       <option value="<%=RS.Fields("codtipocontacto")%>" <% if tipocontacto<>"" then%><%if RS.fields("codtipocontacto")=int(tipocontacto) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
                                               <%
                                               RS.MoveNext
                                               loop
                                               RS.Close
                                               %>        
                                       </select>
                                </td>
            </tr>	
			</table>
			
            <table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;1. Actualización de Base de Datos:</b></font></td>
			</tr>
			</table>			
			
			<table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=25% colspan=2><font size=2 face=Arial color=#00529B>&nbsp;Corresponde al Domicilio del Sistema:</font></td>
				 <td bgcolor="#E9F8FE"><input name="v_dirsistema" type="radio" onclick="document.formula.dirsistema.value='1';" <%if dirsistema="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_dirsistema" type="radio" onclick="document.formula.dirsistema.value='0';" <%if dirsistema="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB" width="23%"><font size=2 face=Arial color=#00529B>&nbsp;Dirección Adicional:</font></td>
				 <td bgcolor="#BEE8FB" align="center" width="2%"><a href="javascript:agregardir();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar Dirección" title="Agregar Dirección"></a></td>
				 <td bgcolor="#E9F8FE">
						<select name="coddireccionnueva" style="font-size: xx-small; width: 94%">
						<option value="">Seleccionar Dirección</option>
						<%
					    sql="select A.coddireccionnueva,A.direccion,B.departamento,B.provincia,B.distrito from DireccionNueva A left outer join Ubigeo B on A.coddpto=B.coddpto and A.codprov=B.codprov and A.coddist=B.coddist where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.coddireccionnueva desc"
					    consultar sql,RS
					    Do While Not  RS.EOF
							direccion=RS.Fields("direccion")
							departamento=RS.Fields("departamento")
							provincia=RS.Fields("provincia")
							distrito=RS.Fields("distrito")
					    %>
						<option value="<%=RS.Fields("coddireccionnueva")%>" <%if obtener("coddireccionnueva")=Cstr(RS.Fields("coddireccionnueva")) or agreguedir<>"" then%> selected<%agreguedir=""%><%end if%>><%=direccion & " - " & distrito & " - " & provincia & " - " & departamento%></option>
					    <%
					    RS.MoveNext
					    loop
					    RS.Close							
						%>
						</select>				 
				 </td>
			</tr>										
			<tr>
				 <td bgcolor="#BEE8FB" width=23%><font size=2 face=Arial color=#00529B>&nbsp;Reporte de Teléfonos:</font></td>
				 <td bgcolor="#BEE8FB" align="center" width="2%"><a href="javascript:agregartelf();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar Teléfono" title="Agregar Teléfono"></a></td>
				 <td bgcolor="#E9F8FE">
				 		<table width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;</a></font></td>
							<td width=38 bgcolor="#BEE8FB" align="center"><font size=2 face=Arial color=#00529B><b>&nbsp;Tipo</b></font></td>
							<td width=40 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Pref</b></font></td>
							<td width=80 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Teléfono</b></font></td>
							<td width=40 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Extensión</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Situación</b></font></td>
						</tr>					 		
						<%if IsNumeric(fono1) then%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension1<>"0000" then%><%=extension1%><%end if%></font></td>
							<td bgcolor="#E9F8FE">
						    <select name="situaciontel1" style="font-size: xx-small; width: 50%">
						        <option value="">Seleccionar Situación</option>
						        <%
					            sql="select codsituacion,descripcion from SituacionTelefono A where A.activo=1 order by codsituacion"
					            consultar sql,RS
					            Do While Not  RS.EOF
					            %>
						        <option value="<%=RS.Fields("codsituacion")%>" <%if situaciontel1<>"" then%><%if RS.fields("codsituacion")=int(situaciontel1) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
					            <%
					            RS.MoveNext
					            loop
					            RS.Close							
						        %>
						    </select>			
						    </td>					
						</tr>	
						<%end if%>
						<%if IsNumeric(fono2) then%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension2<>"0000" then%><%=extension2%><%end if%></font></td>
							<td bgcolor="#E9F8FE">
						    <select name="situaciontel2" style="font-size: xx-small; width: 50%">
						        <option value="">Seleccionar Situación</option>
						        <%
					            sql="select codsituacion,descripcion from SituacionTelefono A where A.activo=1 order by codsituacion"
					            consultar sql,RS
					            Do While Not  RS.EOF
					            %>
						        <option value="<%=RS.Fields("codsituacion")%>" <%if situaciontel2<>"" then%><%if RS.fields("codsituacion")=int(situaciontel2) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
					            <%
					            RS.MoveNext
					            loop
					            RS.Close							
						        %>
						    </select>			
						    </td>								
						</tr>	
						<%end if%>						
						<%if IsNumeric(fono3) then%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension3<>"0000" then%><%=extension3%><%end if%></font></td>
							<td bgcolor="#E9F8FE">
                            <select name="situaciontel3" style="font-size: xx-small; width: 50%">
						        <option value="">Seleccionar Situación</option>
						        <%
					            sql="select codsituacion,descripcion from SituacionTelefono A where A.activo=1 order by codsituacion"
					            consultar sql,RS
					            Do While Not  RS.EOF
					            %>
						        <option value="<%=RS.Fields("codsituacion")%>" <%if situaciontel3<>"" then%><%if RS.fields("codsituacion")=int(situaciontel3) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
					            <%
					            RS.MoveNext
					            loop
					            RS.Close							
						        %>
						    </select>			
						    </td>					
						</tr>	
						<%end if%>
						<%if IsNumeric(fono4) then%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension4<>"0000" then%><%=extension4%><%end if%></font></td>
							<td bgcolor="#E9F8FE">
                            <select name="situaciontel4" style="font-size: xx-small; width: 50%">
						        <option value="">Seleccionar Situación</option>
						        <%
					            sql="select codsituacion,descripcion from SituacionTelefono A where A.activo=1 order by codsituacion"
					            consultar sql,RS
					            Do While Not  RS.EOF
					            %>
						        <option value="<%=RS.Fields("codsituacion")%>" <%if situaciontel4<>"" then%><%if RS.fields("codsituacion")=int(situaciontel4) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
					            <%
					            RS.MoveNext
					            loop
					            RS.Close							
						        %>
						    </select>			
						    </td>					
						</tr>	
						<%end if%>
						<%if IsNumeric(fono5) then%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension5<>"0000" then%><%=extension5%><%end if%></font></td>
							<td bgcolor="#E9F8FE">
                            <select name="situaciontel5" style="font-size: xx-small; width: 50%">
						        <option value="">Seleccionar Situación</option>
						        <%
					            sql="select codsituacion,descripcion from SituacionTelefono A where A.activo=1 order by codsituacion"
					            consultar sql,RS
					            Do While Not  RS.EOF
					            %>
						        <option value="<%=RS.Fields("codsituacion")%>" <%if situaciontel5<>"" then%><%if RS.fields("codsituacion")=int(situaciontel5) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
					            <%
					            RS.MoveNext
					            loop
					            RS.Close							
						        %>
						    </select>								
						    </td>
						</tr>	
						<%end if%>																		
						<%
						sql="select codtelefononuevo,codtipotelefono,prefijo,fono,extension,codsituacion from TelefonoNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codtelefononuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:eliminartelf('<%=RS.Fields("codtelefononuevo")%>');"><img src="imagenes/eliminar.png" border=0 alt="Eliminar Teléfono" title="Eliminar Teléfono"></a></font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("codtipotelefono")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("prefijo")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("fono")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("extension")%></font></td>
                            <td bgcolor="#E9F8FE">
                            <select name="situaciontelnuevo<%=RS.fields("codtelefononuevo")%>" style="font-size: xx-small; width: 50%">
						        <option value="">Seleccionar Situación</option>
						        <%
					            sql="select codsituacion,descripcion from SituacionTelefono A where A.activo=1 order by codsituacion"
					            consultar sql,RS1
					            Do While Not  RS1.EOF
					            %>
						        <option value="<%=RS1.Fields("codsituacion")%>" <%if obtener("situaciontelnuevo" & RS.fields("codtelefononuevo"))<>"" then%><%if int(obtener("situaciontelnuevo" & RS.fields("codtelefononuevo")))=RS1.fields("codsituacion") then%> selected<%end if%><%else%><%if RS1.fields("codsituacion")=RS.fields("codsituacion") then%> selected<%end if%><%end if%>><%=RS1.Fields("Descripcion")%></option>
					            <%
					            RS1.MoveNext
					            loop
					            RS1.Close							
						        %>
						    </select>								
						    </td>							
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>							
						</table>				 
				 </td>
			</tr>
            </table>

            <table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;2. Actividad Principal:</b></font></td>
			</tr>
			</table>			
							
            <table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Seleccione la actividad principal:</font></td>
				 <td bgcolor="#E9F8FE">
				        <input name="v_actprinc" type="radio" onclick="document.formula.actprinc.value='Manufactura';" <%if actprinc="Manufactura" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Manufactura&nbsp;</font><BR>
				        <input name="v_actprinc" type="radio" onclick="document.formula.actprinc.value='Comercio';" <%if actprinc="Comercio" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Comercio&nbsp;</font><BR>
				        <input name="v_actprinc" type="radio" onclick="document.formula.actprinc.value='Transporte';" <%if actprinc="Transporte" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Transporte&nbsp;</font><BR>
				        <input name="v_actprinc" type="radio" onclick="document.formula.actprinc.value='Servicios';" <%if actprinc="Servicios" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Servicios&nbsp;</font><BR>
				        <input name="v_actprinc" type="radio" onclick="document.formula.actprinc.value='Construcción';" <%if actprinc="Construcción" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Construcción&nbsp;</font><BR>
				        <input name="v_actprinc" type="radio" onclick="document.formula.actprinc.value='Extractivas';" <%if actprinc="Extractivas" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Extractivas&nbsp;</font><BR>
				        <input name="v_actprinc" type="radio" onclick="document.formula.actprinc.value='Otros';" <%if actprinc="Otros" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Otros&nbsp;(Detallar)&nbsp;:&nbsp;</font><input name="actividadotros" type=text maxlength=250 value="<%=obtener("actividadotros")%>" style="font-size: x-small; width: 220px;"><BR>
                 </td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Detallar actividad adicional:</font></td>
				 <td bgcolor="#E9F8FE"><input name="actividadadicional" type=text maxlength=250 value="<%=obtener("actividadadicional")%>" style="font-size: x-small; width: 344px;"></td>
			</tr>		
			</table>					
			
			
            <table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;3. Características del Negocio:</b></font></td>
			</tr>
			</table>		
            <table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;En&nbsp;Actividad:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_enactividad" type="radio" onclick="document.formula.enactividad.value='1';" <%if enactividad="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_enactividad" type="radio" onclick="document.formula.enactividad.value='0';" <%if enactividad="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;Local&nbsp;Alquilado:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_localalquilado" type="radio" onclick="document.formula.localalquilado.value='1';" <%if localalquilado="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_localalquilado" type="radio" onclick="document.formula.localalquilado.value='0';" <%if localalquilado="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;Local&nbsp;Financiado:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_localfinanciado" type="radio" onclick="document.formula.localfinanciado.value='1';" <%if localfinanciado="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_localfinanciado" type="radio" onclick="document.formula.localfinanciado.value='0';" <%if localfinanciado="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Facturación&nbsp;Principal&nbsp;S/.:</font></td>
				 <td bgcolor="#E9F8FE"><input name="facturacionprincipal" type=text maxlength=250 value="<%=obtener("facturacionprincipal")%>" style="text-align: right; font-size: x-small; width: 100px;"></td>				 
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;Puerta&nbsp;a&nbsp;Calle:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_puertacalle" type="radio" onclick="document.formula.puertacalle.value='1';" <%if puertacalle="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_puertacalle" type="radio" onclick="document.formula.puertacalle.value='0';" <%if puertacalle="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;Oficina&nbsp;Administrativa:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_ofadministrativa" type="radio" onclick="document.formula.ofadministrativa.value='1';" <%if ofadministrativa="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_ofadministrativa" type="radio" onclick="document.formula.ofadministrativa.value='0';" <%if ofadministrativa="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;Casa&nbsp;/&nbsp;Negocio:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_casanegocio" type="radio" onclick="document.formula.casanegocio.value='1';" <%if casanegocio="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_casanegocio" type="radio" onclick="document.formula.casanegocio.value='0';" <%if casanegocio="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Facturación&nbsp;Adicional&nbsp;S/.:</font></td>
				 <td bgcolor="#E9F8FE"><input name="facturacionadicional" type=text maxlength=250 value="<%=obtener("facturacionadicional")%>" style="text-align: right; font-size: x-small; width: 100px;"></td>				 
			</tr>	
            <tr>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;Labor&nbsp;Artesanal:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_laborartesanal" type="radio" onclick="document.formula.laborartesanal.value='1';" <%if laborartesanal="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_laborartesanal" type="radio" onclick="document.formula.laborartesanal.value='0';" <%if laborartesanal="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;Tiene&nbsp;Existencias:</font></td>
				 <td bgcolor="#E9F8FE" width=10%><input name="v_existencias" type="radio" onclick="document.formula.existencias.value='1';" <%if existencias="1" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sí&nbsp;</font><input name="v_existencias" type="radio" onclick="document.formula.existencias.value='0';" <%if existencias="0" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No</font></td>
				 <td bgcolor="#BEE8FB" width=10%><font size=2 face=Arial color=#00529B>&nbsp;N°&nbsp;Personas:</font></td>
				 <td bgcolor="#E9F8FE" colspan=3><input name="nropersonas" type=text maxlength=250 value="<%=obtener("nropersonas")%>" style="text-align: right; font-size: x-small; width: 50px;"></td>				 
			</tr>	
			</table>					
			
            <table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;4. Disposición para el pago:</b></font></td>
			</tr>
			</table>	
			
            <table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Compromiso de Pago para:</font></td>
				 <td bgcolor="#E9F8FE"><input name="fechapdp" type=text maxlength=10 id="sel4" readonly value="<%if IsDate(fechapdp) then%><%=fechapdp%><%else%><%=date()%><%end if%>" style="font-size: x-small; width: 80px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel4', '%d/%m/%Y');"><font size=1 face=Arial color=#00529B>&nbsp;No puede exceder de 5 días</font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Motivos del Atraso:</font></td>
				 <td bgcolor="#E9F8FE">
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Reducción de Ventas - Competencia del Sector';" <%if motivoatraso="Reducción de Ventas - Competencia del Sector" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Reducción de Ventas - Competencia del Sector&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Reducción de Ventas - Falta Capital de Trabajo';" <%if motivoatraso="Reducción de Ventas - Falta Capital de Trabajo" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Reducción de Ventas - Falta Capital de Trabajo&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Reducción de Ventas - Existencias';" <%if motivoatraso="Reducción de Ventas - Existencias" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Reducción de Ventas - Existencias&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Flujo de Caja inadecuado';" <%if motivoatraso="Flujo de Caja inadecuado" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Flujo de Caja inadecuado&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Sobreendeudado';" <%if motivoatraso="Sobreendeudado" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Sobreendeudado&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Retiro de Socios';" <%if motivoatraso="Retiro de Socios" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Retiro de Socios&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Pago a trabajadores';" <%if motivoatraso="Pago a trabajadores" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Pago a trabajadores&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Siniestro';" <%if motivoatraso="Siniestro" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Siniestro&nbsp;</font><BR>
				        <input name="v_motivoatraso" type="radio" onclick="document.formula.motivoatraso.value='Otras causas';" <%if motivoatraso="Otras causas" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Otras Causas&nbsp;(Detallar)&nbsp;:&nbsp;</font><input name="otrascausasatraso" type=text maxlength=250 value="<%=obtener("otrascausasatraso")%>" style="font-size: x-small; width: 220px;">
				 </td>
			</tr>	
            <tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Como afrontará el pago:</font></td>
				 <td bgcolor="#E9F8FE">
				        <input name="v_afrontapago" type="radio" onclick="document.formula.afrontapago.value='Préstamo Familiar';" <%if afrontapago="Préstamo Familiar" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Préstamo Familiar&nbsp;</font><BR>
				        <input name="v_afrontapago" type="radio" onclick="document.formula.afrontapago.value='Préstamo Bancario';" <%if afrontapago="Préstamo Bancario" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Préstamo Bancario&nbsp;</font><BR>
				        <input name="v_afrontapago" type="radio" onclick="document.formula.afrontapago.value='Factura por Cobrar';" <%if afrontapago="Factura por Cobrar" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Factura por Cobrar&nbsp;</font><BR>
				        <input name="v_afrontapago" type="radio" onclick="document.formula.afrontapago.value='Junta';" <%if afrontapago="Junta" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Junta&nbsp;</font><BR>
				        <input name="v_afrontapago" type="radio" onclick="document.formula.afrontapago.value='Otros';" <%if afrontapago="Otros" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Otros&nbsp;(Detallar)&nbsp;:&nbsp;</font><input name="otrosafronta" type=text maxlength=250 value="<%=obtener("otrosafronta")%>" style="font-size: x-small; width: 220px;">
				 </td>
			</tr>						
            <tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Cuestiona cobranza:</font></td>
				 <td bgcolor="#E9F8FE">
				        <input name="v_cuestionacobro" type="radio" onclick="document.formula.cuestionacobro.value='Está al día';" <%if cuestionacobro="Está al día" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Está al día&nbsp;</font><BR>
				        <input name="v_cuestionacobro" type="radio" onclick="document.formula.cuestionacobro.value='Reclamo en trámite';" <%if cuestionacobro="Reclamo en trámite" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Reclamo en trámite&nbsp;</font><BR>
				        <input name="v_cuestionacobro" type="radio" onclick="document.formula.cuestionacobro.value='Refinanciación/Reprogramación/Prórroga en trámite';" <%if cuestionacobro="Refinanciación/Reprogramación/Prórroga en trámite" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Refinanciación/Reprogramación/Prórroga en trámite&nbsp;</font><BR>
				        <input name="v_cuestionacobro" type="radio" onclick="document.formula.cuestionacobro.value='Posible Fraude';" <%if cuestionacobro="Posible Fraude" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Posible Fraude&nbsp;</font>
				 </td>
			</tr>				
			</table>

			
            <table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;5. Observaciones del Especialista de Cobranzas:</b></font></td>
			</tr>
			</table>	


            <table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Motivo del No Contacto:</font></td>
				 <td bgcolor="#E9F8FE">
				        <input name="v_nocontacto" type="radio" onclick="document.formula.nocontacto.value='Se mudó';" <%if nocontacto="Se mudó" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Se mudó&nbsp;</font><BR>
				        <input name="v_nocontacto" type="radio" onclick="document.formula.nocontacto.value='No existe dirección';" <%if nocontacto="No existe dirección" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No existe dirección&nbsp;</font><BR>
				        <input name="v_nocontacto" type="radio" onclick="document.formula.nocontacto.value='Dirección no corresponde';" <%if nocontacto="Dirección no corresponde" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;Dirección no corresponde&nbsp;</font><BR>
				        <input name="v_nocontacto" type="radio" onclick="document.formula.nocontacto.value='No opera hace 3 meses';" <%if nocontacto="No opera hace 3 meses" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No opera hace 3 meses&nbsp;</font><BR>
				        <input name="v_nocontacto" type="radio" onclick="document.formula.nocontacto.value='No opera de 6 a 12 meses';" <%if nocontacto="No opera de 6 a 12 meses" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No opera de 6 a 12 meses&nbsp;</font><BR>
				        <input name="v_nocontacto" type="radio" onclick="document.formula.nocontacto.value='No opera hace más de 1 año';" <%if nocontacto="No opera hace más de 1 año" then%> checked<%end if%>><font size=2 face=Arial color=#00529B>&nbsp;No opera hace más de 1 año&nbsp;</font>
				 </td>
			</tr>	
			</table>
			
            <table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;6. Comentarios del Especialista:</b></font></td>
			</tr>
			</table>	
						

            <table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Comentarios:</font></td>
				 <td bgcolor="#E9F8FE"><textarea name="comentario" onfocus="backspaceactivo=1;" onblur="backspaceactivo=0;" style="font-family: 'Arial';font-size: 11px;" rows=4 cols=100 style="font-size: xx-small; width: 85%;" onchange="if(this.value.length>4000){this.value=this.value.substring(0,4000);alert('El texto se truncó a 4000 caracteres, excedió el máximo');}"><%=obtener("comentario")%></textarea></td>
			</tr>	
			</table>
						
            <table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;7. Estado de la Ficha de Visita de Campo:</b></font></td>
			</tr>
			</table>	
						

            <table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Estado:</font></td>
                 <td bgcolor="#E9F8FE">
                    <select name="estado" style="font-size: x-small; width: 250px;">
                               <%
                               sql="select codestado,descripcion from EstadoVisita where activo=1 order by codestado"
                               consultar sql,RS
                               Do While Not  RS.EOF
                               %>
                                       <option value="<%=RS.Fields("codestado")%>" <% if estado<>"" then%><%if RS.fields("codestado")=int(estado) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
                               <%
                               RS.MoveNext
                               loop
                               RS.Close
                               %>        
                       </select>
                 </td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB" width=25%><font size=2 face=Arial color=#00529B>&nbsp;Adjuntar Archivo (.rar):</font></td>
                 <td bgcolor="#E9F8FE"><input type="file" name="archivo" id="archivo"></td>
			</tr>				
			</table>
			
            <table width=100% cellpadding=0 cellspacing=0 border=0>
			<tr bgcolor="#007DC5">
				<td height="40" align=right><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;&nbsp;&nbsp;</td>
			</tr>
			</table>				
																							
			<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
			<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
			<input type="hidden" name="vistapadre1" value="">
			<input type="hidden" name="paginapadre1" value="">			
			<input type="hidden" name="fdcodcen" value="<%=fdcodcen%>">
			<input type="hidden" name="codigocentral" value="<%=codigocentral%>">
			<input type="hidden" name="contrato" value="<%=contrato%>">
			<input type="hidden" name="fechadatos" value="<%=fechadatos%>">
			<input type="hidden" name="fechagestion" value="<%=fechagestion%>">
			<input type="hidden" name="direccioneliminar" value="">
			<input type="hidden" name="telefonoeliminar" value="">
			<input type="hidden" name="emaileliminar" value="">
			<input type="hidden" name="dirsistema" value="<%=dirsistema%>">		
			<input type="hidden" name="telefsistema" value="<%=telefsistema%>">
            <input type="hidden" name="actprinc" value="<%=actprinc%>">
            <input type="hidden" name="enactividad" value="<%=enactividad%>">
            <input type="hidden" name="localalquilado" value="<%=localalquilado%>">
            <input type="hidden" name="localfinanciado" value="<%=localfinanciado%>">
            <input type="hidden" name="puertacalle" value="<%=puertacalle%>">
            <input type="hidden" name="ofadministrativa" value="<%=ofadministrativa%>">
            <input type="hidden" name="casanegocio" value="<%=casanegocio%>">
            <input type="hidden" name="laborartesanal" value="<%=laborartesanal%>">
            <input type="hidden" name="existencias" value="<%=existencias%>">
            <input type="hidden" name="motivoatraso" value="<%=motivoatraso%>">
            <input type="hidden" name="afrontapago" value="<%=afrontapago%>">
            <input type="hidden" name="cuestionacobro" value="<%=cuestionacobro%>">
			<input type="hidden" name="nocontacto" value="<%=nocontacto%>">
			
			
		</form>
		<script language="javascript">
			inicio();
		</script>
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>							
		</body>
		</html>		
		<%
		''Codigo exp excel
		''Si se pide exportar a excel
				if expimp="1" then

					''se coloca arriba para el enlace si no abre directo el archivo
					''sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaExportar' or descripcion='RutaWebExportar'"
					''consultar sql,RS
					''RS.Filter=" descripcion='RutaFisicaExportar'"
					''RutaFisicaExportar=RS.Fields(1)
					''RS.Filter=" descripcion='RutaWebExportar'"
					''RutaWebExportar=RS.Fields(1)	
					''RS.Filter=""			
					''RS.Close					
					''Para Exportar a Excel
					''Primero Cabecera en temp1_(user).txt
					consulta_exp="select 'Cod.Marca','Descripción' "
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select codmarca,descripcion " & _
								 "from CobranzaCM.dbo.marca " & filtrobuscador & " order by codmarca" 
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
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
		window.open("userexpira.asp","_top");
	</script>
	<%	
	end if
	desconectar
else
%>
<script language="javascript">
	alert("Tiempo Expirado");
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>



