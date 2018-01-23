<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admvisitas.asp") then
		fdcodcen=obtener("fdcodcen")
		
		sql="select *,CASE WHEN clasificacion=0 THEN 'NORMAL' WHEN clasificacion=1 THEN 'CPP' WHEN clasificacion=2 THEN 'DEFICIENTE' WHEN clasificacion=3 THEN 'DUDOSO' WHEN clasificacion=4 THEN 'PERDIDA' END as clasifica from Cliente_FichaVisita A where convert(varchar,A.fechadatos,112) + A.codigocentral='" & fdcodcen & "'"
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
		
		fechavisita=obtener("fechavisita")
		horavisita=iif(obtener("horavisita")="","15",obtener("horavisita"))
		minutovisita=iif(obtener("minutovisita")="","00",obtener("minutovisita"))
		tipocontacto=obtener("tipocontacto")
		dirsistema=iif(obtener("dirsistema")="","1",obtener("dirsistema"))
		coddireccionnueva=obtener("coddireccionnueva")	
		
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
			nuevodir=global_popup_IWTSystem(nuevodir,"adicionardireccion.asp?vistapadre1=" + window.name + "&paginapadre1=verficha.asp&vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&fechadatos=<%=fechadatos%>&fechadatos=<%=fechadatos%>&fechavisita=" + document.formula.fechavisita.value + "&horavisita=" + document.formula.horavisita.value + "&minutovisita=" + document.formula.minutovisita.value + "&dirsistema=" + document.formula.dirsistema.value + "&tipocontacto=" + document.formula.tipocontacto.value + "&coddireccionnueva=" + document.formula.coddireccionnueva.value,"Newdir","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 + 250) + ",left=" + (screen.width/4) + ",resizable=yes");
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
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=clasifica%></b></font></td>				 
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Dirección:</font></td>
				 <td bgcolor="#E9F8FE" colspan=3><font size=2 face=Arial color=#00529B><b>&nbsp;<%=direccion & referencia1 & referencia2%> / <%=distrito%> / <%=provincia%> / <%=departamento%></b></font></td>
				 <!--<td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;</b></font></td>-->
			</tr>															
			<!--<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Direcciones:</font></td>
						<td align=right><a href="javascript:agregardir();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar Dirección" title="Agregar Dirección"></a>&nbsp;</td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
						<script language=javascript>
						function visualizardir()
						{
							var filas = document.getElementById('tabladirecciones').rows.length;
							//if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
							if (document.getElementById('imagendir').title=="Mostrar")
							{
								document.getElementById('imagendir').title="Ocultar";
								document.getElementById('imagendir').alt="Ocultar";
								document.getElementById('imagendir').src="imagenes/ocultar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tabladirecciones').rows[i].style.display = '';
								}
							}
							else
							{
								document.getElementById('imagendir').title="Mostrar";
								document.getElementById('imagendir').alt="Mostrar";
								document.getElementById('imagendir').src="imagenes/mostrar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tabladirecciones').rows[i].style.display = 'none';
								}							
							}
						}
						</script>
				 		<table id="tabladirecciones" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizardir();"><img id="imagendir" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Dirección</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Distrito</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Provincia</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Departamento</b></font></td>
						</tr>					 		
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=direccion%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=distrito%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=provincia%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=departamento%></font></td>
						</tr>	
						<%
						sql="select A.coddireccionnueva,A.direccion,B.departamento,B.provincia,B.distrito from DireccionNueva A left outer join Ubigeo B on A.coddpto=B.coddpto and A.codprov=B.codprov and A.coddist=B.coddist where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.coddireccionnueva desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:eliminardir('<%=RS.Fields("coddireccionnueva")%>');"><img src="imagenes/eliminar.png" border=0 alt="Eliminar Dirección" title="Eliminar Dirección"></a></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("direccion")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("distrito")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("provincia")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("departamento")%></font></td>
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>						
						</table>
						<%if obtener("agreguedir")<>"" then%>
						<script language=javascript>
							visualizardir();
						</script>
						<%end if%>
				 </td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Teléfonos:</font></td>
						<td align=right><a href="javascript:agregartelf();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar Teléfono" title="Agregar Teléfono"></a>&nbsp;</td>
					</tr>
					</table>						
				 </td>
				 <td colspan=3>
						<script language=javascript>
						function visualizartelf()
						{
							var filas = document.getElementById('tablatelefonos').rows.length;
							if (document.getElementById('imagentel').title=="Mostrar")
							{
								document.getElementById('imagentel').title="Ocultar";
								document.getElementById('imagentel').alt="Ocultar";
								document.getElementById('imagentel').src="imagenes/ocultar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablatelefonos').rows[i].style.display = '';
								}
							}
							else
							{
								document.getElementById('imagentel').title="Mostrar";
								document.getElementById('imagentel').alt="Mostrar";
								document.getElementById('imagentel').src="imagenes/mostrar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablatelefonos').rows[i].style.display = 'none';
								}							
							}
						}
						</script>
				 		<table id="tablatelefonos" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizartelf();"><img id="imagentel" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td width=38 bgcolor="#BEE8FB" align="center"><font size=2 face=Arial color=#00529B><b>&nbsp;Tipo</b></font></td>
							<td width=40 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Pref</b></font></td>
							<td width=80 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Teléfono</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Extensión</b></font></td>
						</tr>					 		
						<%if fono1<>"" then%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension1<>"0000" then%><%=extension1%><%end if%></font></td>
						</tr>	
						<%end if%>
						<%if fono2<>"" then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension2<>"0000" then%><%=extension2%><%end if%></font></td>
						</tr>	
						<%end if%>						
						<%if fono3<>"" then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension3<>"0000" then%><%=extension3%><%end if%></font></td>
						</tr>	
						<%end if%>
						<%if fono4<>"" then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension4<>"0000" then%><%=extension4%><%end if%></font></td>
						</tr>	
						<%end if%>
						<%if fono5<>"" then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension5<>"0000" then%><%=extension5%><%end if%></font></td>
						</tr>	
						<%end if%>																		
						<%		
						sql="select codtelefononuevo,codtipotelefono,prefijo,fono,extension from TelefonoNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codtelefononuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:eliminartelf('<%=RS.Fields("codtelefononuevo")%>');"><img src="imagenes/eliminar.png" border=0 alt="Eliminar Teléfono" title="Eliminar Teléfono"></a></font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("codtipotelefono")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("prefijo")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("fono")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("extension")%></font></td>
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>							
						</table>
						<%if obtener("agreguetelf")<>"" then%>
						<script language=javascript>
							visualizartelf();
						</script>
						<%end if%>
				 </td>
			</tr>							
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;E-mail:</font></td>
						<td align=right><a href="javascript:agregaremail();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar E-mail" title="Agregar E-mail"></a>&nbsp;</td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
						<script language=javascript>
						function visualizaremail()
						{
							var filas = document.getElementById('tablaemails').rows.length;
							//if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
							if (document.getElementById('imagenemail').title=="Mostrar")
							{
								document.getElementById('imagenemail').title="Ocultar";
								document.getElementById('imagenemail').alt="Ocultar";
								document.getElementById('imagenemail').src="imagenes/ocultar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablaemails').rows[i].style.display = '';
								}
							}
							else
							{
								document.getElementById('imagenemail').title="Mostrar";
								document.getElementById('imagenemail').alt="Mostrar";
								document.getElementById('imagenemail').src="imagenes/mostrar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablaemails').rows[i].style.display = 'none';
								}							
							}
						}
						</script>
				 		<table id="tablaemails" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizaremail();"><img id="imagenemail" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;E-mail</b></font></td>
						</tr>					 		
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=email%></font></td>
						</tr>	
						<%		
						sql="select A.codemailnuevo,A.email from EmailNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codemailnuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:eliminaremail('<%=RS.Fields("codemailnuevo")%>');"><img src="imagenes/eliminar.png" border=0 alt="Eliminar E-Mail" title="Eliminar E-Mail"></a></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("email")%></font></td>
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>						
						</table>
						<%if obtener("agregueemail")<>"" then%>
						<script language=javascript>
							visualizaremail();
						</script>
						<%end if%>
				 </td>
			</tr>		
			-->		
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
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Total</b></font></td>						
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Vencido</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Prov.Const.</b></font></td>
			</tr>
			<%
			RS1.Close
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
				 <td bgcolor="#E9F8FE"><input name="fechavisita" type=text maxlength=10 id="sel3" readonly value="<%if IsDate(fechavisita) then%><%=fechavisita%><%else%><%=date()%><%end if%>" style="font-size: x-small; width: 80px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel3', '%d/%m/%Y');"><font size=2 face=Arial color="#00529B">&nbsp;Hora:&nbsp;</font><input name="horavisita" type=text maxlength=2 value="<%=horavisita%>" style="text-align: center;font-size: x-small; width: 20px;"><font size=2 face=Arial color="#00529B"><b>&nbsp;:&nbsp;</b></font><input name="minutovisita" type=text maxlength=2 value="<%=minutovisita%>" style="text-align: center;font-size: x-small; width: 20px;"></td>
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
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Dirección Adicional:</font></td>
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
            </table>								
								
			<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
			<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
			<input type="hidden" name="codigocentral" value="<%=codigocentral%>">
			<input type="hidden" name="contrato" value="<%=contrato%>">
			<input type="hidden" name="fechadatos" value="<%=fechadatos%>">
			<input type="hidden" name="fechagestion" value="<%=fechagestion%>">
			<input type="hidden" name="direccioneliminar" value="">
			<input type="hidden" name="telefonoeliminar" value="">
			<input type="hidden" name="emaileliminar" value="">
			<input type="hidden" name="dirsistema" value="<%=dirsistema%>">			
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



