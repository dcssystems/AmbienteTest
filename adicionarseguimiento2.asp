<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admimpagado.asp") then
		codigocentral=obtener("codigocentral")
		
		contrato=obtener("contrato")
		fechadatos=obtener("fechadatos")
		fechagestion=obtener("fechagestion")
		codtipoaccion=obtener("codtipoaccion")
		codtipocontacto=obtener("codtipocontacto")
		codgestion=obtener("codgestion")
		
		sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUpload' or descripcion='RutaWebUpload'"
		consultar sql,RS3
		RS3.Filter=" descripcion='RutaFisicaUpload'"
		RutaFisicaUpload=RS3.Fields(1)
		RS3.Filter=" descripcion='RutaWebUpload'"
		RutaWebUpload=RS3.Fields(1)				
		RS3.Filter=""
		RS3.Close	
		
%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Agregar Gestión</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				function agregar()
				{
				if(formula.codtipoaccion.selectedIndex==0){alert("Debe seleccionar la Acción de la Gestión.");return;}
				if(formula.coddirtel.selectedIndex<0){alert("Debe seleccionar la Dirección o Teléfono de la Gestión.");return;}
				if(formula.codgestion.selectedIndex<0){alert("Debe seleccionar la Respuesta de la Gestión.");return;}
				if(trim(formula.comentario.value)==""){alert("Debe ingresar un comentario de la Gestión.");return;}
				/*if (document.formula.codgestion.options[document.formula.codgestion.selectedIndex].innerText.toUpperCase().replace('PROMESA','').length!=document.formula.codgestion.options[document.formula.codgestion.selectedIndex].innerText.length)
				{
					if(document.formula.fechapromesa.value==""){alert("Debe ingresar una fecha de promesa de pago.");return;}
					var fprom=new Date(parseInt(document.formula.fechapromesa.value.substring(6,10)),eval(document.formula.fechapromesa.value.substring(3,5))-1,parseInt(document.formula.fechapromesa.value.substring(0,2)),0,0,0);
					var fhoy=new Date(<%=Year(Date())%>,<%=Month(Date())-1%>,<%=Day(date())%>,0,0,0);
					var fmax=new Date(<%=Year(Date()+5)%>,<%=Month(Date()+5)-1%>,<%=Day(date()+5)%>,0,0,0);
					var one_day=1000*60*60*24;
					var dif1=Math.ceil((fhoy.getTime()-fprom.getTime())/(one_day));
					var dif2=Math.ceil((fmax.getTime()-fprom.getTime())/(one_day));
					if(dif1>0){alert("La fecha de promesa de pago debe ser igual o mayor que hoy.");return;}
					if(dif2<0){alert("La fecha de promesa no debe superar los 5 días siguientes.");return;}
					if(document.formula.importe1.value==""||isNaN(trim(document.formula.importe1.value.replace(",","")))){alert("Debe ingresar un dato numérico en el primer importe de la promesa.");return;}
					if(document.formula.importe2.value!="")if(isNaN(trim(document.formula.importe2.value.replace(",","")))){alert("Debe ingresar un dato numérico en el segundo importe de la promesa.");return;}
				}	*/				
					document.formuladata.codtipoaccion_data.value=document.formula.codtipoaccion.value;
					document.formuladata.codtipocontacto_data.value=document.formula.codtipocontacto.value;
					document.formuladata.coddirtel_data.value=document.formula.coddirtel.value;
					document.formuladata.codgestion_data.value=document.formula.codgestion.value;
					document.formuladata.comentario_data.value=document.formula.comentario.value;
					document.formuladata.fechapromesa_data.value=document.formula.fechapromesa.value;
					document.formuladata.divisa1_data.value=document.formula.divisa1.value;
					document.formuladata.importe1_data.value=document.formula.importe1.value;
					document.formuladata.divisa2_data.value=document.formula.divisa2.value;
					document.formuladata.importe2_data.value=document.formula.importe2.value;
					document.formuladata.submit();
				}			
				function actualizarcodrespuesta()
				{
				document.formula.action="adicionarseguimiento2.asp";
				document.formula.codrespuestaact.value = 1;
				document.formula.submit();				
				}
				function trim(string)
				{
					while(string.substr(0,1)==" ")
					string = string.substring(1,string.length) ;
					while(string.substr(string.length-1,1)==" ")
					string = string.substring(0,string.length-2) ;
					return string;
				}			
				function isEmailAddress(theElement)
				{
				var s = theElement.value;
				var filter=/^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/ig;
				if (s.length == 0 ) return true;
				   if (filter.test(s))
				      return true;
				   else
					theElement.focus();
					return false;
				}					
			</script>
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
			</style>
			
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
			
						
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
				<form name=formula method=post action="adicionarseguimiento2.asp">
					<table width=100% height=80% border=0 cellspacing=0 cellpadding=0>
					<tr bgcolor="#007DC5">
					<td align="middle" height="22" colspan=2><font size=2 face=Arial color="#FFFFFF"><b>Agregar Gestión</b></font></td>
					</tr>										  
					<tr>
						<td width=22%><font face=Arial size=2 color=#483d8b>&nbsp;Acción realizada:</font></td>
						<td>
						<select name="codtipoaccion" onchange="actualizarcodrespuesta()" style="font-size: xx-small; width: 94%">
						<option value="">Seleccionar Acción</option>
						<%
						sql="select A.codtipoaccion,A.descripcion from TipoAccion A inner join Gestion B on A.codtipoaccion=B.codtipoaccion and A.activo=1 and B.activo=1 group by A.codtipoaccion,A.descripcion order by A.descripcion"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
							<option value="<%=RS.Fields("codtipoaccion")%>" <% if codtipoaccion<>"" then%><%if RS.fields("codtipoaccion")=int(codtipoaccion) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>	
						</select>
						</td>
					</tr>	
					<tr>
						<td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;Tipo de Contacto:</font></td>
						<td bgcolor="#f5f5f5">
						<select name="codtipocontacto" onchange="actualizarcodrespuesta()" style="font-size: xx-small; width: 94%">
						<%
						sql="select codtipocontacto,descripcion from tipocontacto where activo=1 order by codtipocontacto"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
							<option value="<%=RS.Fields("codtipocontacto")%>" <% if codtipocontacto<>"" then%><%if RS.fields("codtipocontacto")=int(codtipocontacto) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>	
						</select>
						</td>
					</tr>						
					<tr>
					<td><font face=Arial size=2 color=#483d8b>&nbsp;Dirección / Teléfono:</font></td>
					<td>
						<select name="coddirtel" style="font-size: xx-small; width: 94%">
						<%
						''1: LLAMADA
						''2: VISITA
						''CARGAR NUMEROS O DIRECCIONES
						
						''NO PUEDE INGRESAR UNA GESTION DE FECHA ANTERIOR
						if codtipoaccion="2" then ''visita
						
							sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
							consultar sql,RS	
							maxfechagestion=rs.fields(0)
							RS.Close	
		
							if mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)=CStr(maxfechagestion) then
								vistabusqueda="VERULTIMPAGADO"
							else
								vistabusqueda="VERIMPAGADO"
							end if								
							
							sql="select top 1 direccion,departamento,provincia,distrito from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "'"
							''Response.Write sql
							consultar sql,RS1	
							if not RS1.EOF then
								direccion=RS1.Fields("direccion")
								departamento=RS1.Fields("departamento")
								provincia=RS1.Fields("provincia")
								distrito=RS1.Fields("distrito")
							end if		
							RS1.Close					
							%>
							<option value="dirprin" <%if obtener("coddirtel")="dirprin" then%> selected<%end if%>><%=direccion & " - " & distrito & " - " & provincia & " - " & departamento%></option>
							<%
						    sql="select A.coddireccionnueva,A.direccion,B.departamento,B.provincia,B.distrito from DireccionNueva A left outer join Ubigeo B on A.coddpto=B.coddpto and A.codprov=B.codprov and A.coddist=B.coddist where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.coddireccionnueva desc"
						    consultar sql,RS
						    Do While Not  RS.EOF
								direccion=RS.Fields("direccion")
								departamento=RS.Fields("departamento")
								provincia=RS.Fields("provincia")
								distrito=RS.Fields("distrito")
						    %>
							<option value="<%=RS.Fields("coddireccionnueva")%>" <%if obtener("coddirtel")=Cstr(RS.Fields("coddireccionnueva")) then%> selected<%end if%>><%=direccion & " - " & distrito & " - " & provincia & " - " & departamento%></option>  
						    <%
						    RS.MoveNext
						    loop
						    RS.Close							
						else
							if codtipoaccion<>"2" then ''no visita
								sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
								consultar sql,RS	
								maxfechagestion=rs.fields(0)
								RS.Close	
		
								if mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)=CStr(maxfechagestion) then
									vistabusqueda="VERULTIMPAGADO"
								else
									vistabusqueda="VERIMPAGADO"
								end if								
							
								sql="select top 1 tipofono1,prefijo1,fono1,extension1,tipofono2,prefijo2,fono2,extension2,tipofono3,prefijo3,fono3,extension3,tipofono4,prefijo4,fono4,extension4,tipofono5,prefijo5,fono5,extension5 from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "'"
								''Response.Write sql
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
								end if	
								RS1.Close
									if fono1<>"" then
									%>
									<option value="fono1" <%if obtener("coddirtel")="fono1" then%> selected<%end if%>><%=tipofono1 & " - " & prefijo1 & " - " & fono1%><%if extension1<>"0000" then%><%=" - " & extension1%><%end if%></option>
									<%				    
									end if
									if fono2<>"" then
									%>
									<option value="fono2" <%if obtener("coddirtel")="fono2" then%> selected<%end if%>><%=tipofono2 & " - " & prefijo2 & " - " & fono2%><%if extension2<>"0000" then%><%=" - " & extension2%><%end if%></option>
									<%				    
									end if
									if fono3<>"" then
									%>
									<option value="fono3" <%if obtener("coddirtel")="fono3" then%> selected<%end if%>><%=tipofono3 & " - " & prefijo3 & " - " & fono3%><%if extension3<>"0000" then%><%=" - " & extension3%><%end if%></option>
									<%				    
									end if
									if fono4<>"" then
									%>
									<option value="fono4" <%if obtener("coddirtel")="fono4" then%> selected<%end if%>><%=tipofono4 & " - " & prefijo4 & " - " & fono4%><%if extension4<>"0000" then%><%=" - " & extension4%><%end if%></option>
									<%				    
									end if
									if fono5<>"" then
									%>
									<option value="fono5" <%if obtener("coddirtel")="fono5" then%> selected<%end if%>><%=tipofono5 & " - " & prefijo5 & " - " & fono5%><%if extension5<>"0000" then%><%=" - " & extension5%><%end if%></option>
									<%				    
									end if
							
								sql="select codtelefononuevo,codtipotelefono,prefijo,fono,extension from TelefonoNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codtelefononuevo desc"
								consultar sql,RS
								Do While Not  RS.EOF
									tipofono=RS.Fields("codtipotelefono")
									prefijo=RS.Fields("prefijo")
									fono=RS.Fields("fono")
									extension=RS.Fields("extension")
								%>
								    <option value="<%=RS.Fields("codtelefononuevo")%>" <%if obtener("coddirtel")=CStr(RS.Fields("codtelefononuevo")) then%> selected<%end if%>><%=tipofono & " - " & prefijo & " - " & fono%><%if extension<>"0000" then%><%=" - " & extension%><%end if%></option>
								<%
								RS.MoveNext
								loop
								RS.Close
							end if	
						end if
						%>
						</select>
						</td>
					</tr>					
					<tr>
					<td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;Respuesta Gestión:</font></td>
						<td bgcolor="#f5f5f5">
							<table width=100% border=0 cellspacing=0 cellpadding=0 id="tablapromesa">
							<tr>
								<td>
								<select name="codgestion" style="font-size: xx-small; width: 94%" onchange="validarpromesa();">
								<%
								if codtipocontacto<>"" and codtipoaccion<>"" then
									sql="select codgestion,Descripcion from Gestion where activo=1 and codtipocontacto=" & codtipocontacto & " and codtipoaccion=" & codtipoaccion & " order by codgestion"
									consultar sql,RS
									Do While Not  RS.EOF
									%>
										<option value="<%=RS.Fields("codgestion")%>" <% if codgestion<>"" then%><%if RS.fields("codgestion")=int(codgestion) then%> selected<%end if%><%end if%>><%=RS.Fields("codgestion") & " - " & RS.Fields("Descripcion")%></option>
									<%
									RS.MoveNext
									loop
									RS.Close
								end if
								%>							
								</select>
								</td>
							</tr>
							<tr style="display: none">
								<td><font face=Arial size=2 color=#483d8b>Fecha de Promesa de Pago:&nbsp;<input name="fechapromesa" type=text readonly maxlength=10 id="selfp"  value="<%if IsDate(obtener("fechapromesa")) then%><%=obtener("fechapromesa")%><%end if%>" style="font-size: xx-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('selfp', '%d/%m/%Y');"></font>							
							</tr>
							<tr style="display: none">
								<td><font face=Arial size=2 color=#483d8b>Monto a Pagar:&nbsp;<select name="divisa1" style="font-size: xx-small;"><option value="PEN" <%if obtener("divisa1")="PEN" or obtener("divisa1")="" then%> selected<%end if%>>PEN</option><option value="USD" <%if obtener("divisa1")="USD" then%> selected<%end if%>>USD</option><option value="EUR" <%if obtener("divisa1")="EUR" then%> selected<%end if%>>EUR</option></select>&nbsp;<input name="importe1" type=text value="<%=obtener("importe1")%>" style="font-size: xx-small; width: 60px;text-align: right;"></font>
							</tr>							
							<tr style="display: none">
								<td><font face=Arial size=2 color=#483d8b>Monto a Pagar:&nbsp;<select name="divisa2" style="font-size: xx-small;"><option value="PEN" <%if obtener("divisa2")="PEN" then%> selected<%end if%>>PEN</option><option value="USD" <%if obtener("divisa2")="USD" or obtener("divisa2")="" then%> selected<%end if%>>USD</option><option value="EUR" <%if obtener("divisa2")="EUR" then%> selected<%end if%>>EUR</option></select>&nbsp;<input name="importe2" type=text value="<%=obtener("importe2")%>" style="font-size: xx-small; width: 60px;text-align: right;"></font>
							</tr>														
							</table>
							<script language=javascript>
								function validarpromesa()
								{
									var filas = document.getElementById('tablapromesa').rows.length;
									if(document.formula.codgestion.selectedIndex>=0)
									{
										/* if (document.formula.codgestion.options[document.formula.codgestion.selectedIndex].innerText.toUpperCase().replace('PROMESA','').length!=document.formula.codgestion.options[document.formula.codgestion.selectedIndex].innerText.length)
										{
											for (i = 1; i < filas; i++)
											{
												document.getElementById('tablapromesa').rows[i].style.display = '';
											}
										}
										else
										{
											for (i = 1; i < filas; i++)
											{
												document.getElementById('tablapromesa').rows[i].style.display = 'none';
											}
										}	*/								
									}
									else
									{
										for (i = 1; i < filas; i++)
										{
											document.getElementById('tablapromesa').rows[i].style.display = 'none';
										}									
									}
								}						
								validarpromesa();		
							</script>
						</td>
					</tr>						
					<tr>
						<td><font face=Arial size=2 color=#483d8b>&nbsp;Comentario:</font></td>
						<td><textarea name="comentario" onfocus="backspaceactivo=1;" onblur="backspaceactivo=0;" style="font-family: 'Arial';font-size: 11px;" rows=4 cols=100 style="font-size: xx-small; width: 85%;" onchange="if(this.value.length>4000){this.value=this.value.substring(0,4000);alert('El texto se truncó a 4000 caracteres, excedió el máximo');}"><%=obtener("comentario")%></textarea></td>
					</tr>
					<input type=hidden name="codrespuestaact" value="">
					<input type=hidden name="vistapadre1" value="<%=obtener("vistapadre1")%>">
					<input type=hidden name="paginapadre1" value="<%=obtener("paginapadre1")%>">
					<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
					<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
					<input type="hidden" name="codigocentral" value="<%=codigocentral%>">
					<input type="hidden" name="contrato" value="<%=contrato%>">
					<input type="hidden" name="fechadatos" value="<%=fechadatos%>">
					<input type="hidden" name="fechagestion" value="<%=fechagestion%>">					
				</form>						
				</table>
				<table width=100% height=20% border=0 cellspacing=0 cellpadding=0>
				<form name="formuladata" method=post enctype="multipart/form-data" action="uploadfileseguimiento.asp">
					<tr>
						<td align=center valign=top>
							<font face=Arial size=2 color=#483d8b>Adjuntar archivo:&nbsp;</font><input type="file" name="archivo" id="archivo">
						</td>
					</tr>
					<tr>					
						<td bgcolor="#F5F5F5" align=right><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>
					</tr>
					<input type=hidden name="codtipoaccion_data" value="">
					<input type=hidden name="codtipocontacto_data" value="">
					<input type=hidden name="coddirtel_data" value="">
					<input type=hidden name="codgestion_data" value="">
					<input type=hidden name="comentario_data" value="">
					<input type=hidden name="fechapromesa_data" value="">
					<input type=hidden name="divisa1_data" value="">
					<input type=hidden name="importe1_data" value="">
					<input type=hidden name="divisa2_data" value="">
					<input type=hidden name="importe2_data" value="">					
					<input type=hidden name="vistapadre1_data" value="<%=obtener("vistapadre1")%>">
					<input type=hidden name="paginapadre1_data" value="<%=obtener("paginapadre1")%>">
					<input type="hidden" name="vistapadre_data" value="<%=obtener("vistapadre")%>">
					<input type="hidden" name="paginapadre_data" value="<%=obtener("paginapadre")%>">
					<input type="hidden" name="codigocentral_data" value="<%=codigocentral%>">
					<input type="hidden" name="contrato_data" value="<%=contrato%>">
					<input type="hidden" name="fechadatos_data" value="<%=fechadatos%>">
					<input type="hidden" name="fechagestion_data" value="<%=fechagestion%>">									
				</form>					
				</table>
			</body>
		</html>	
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
	//window.open("index.html","_top");
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>


