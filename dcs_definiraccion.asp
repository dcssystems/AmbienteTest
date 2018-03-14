<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admfacultad.asp") then
	buscador=obtener("buscador")	
	codfacultad=obtener("codfacultad")
		if obtener("agregardato")<>"" then
		codgrupofacultad=obtener("codgrupofacultad")
		descripcion=obtener("descripcion")
		pagina=obtener("pagina")
		orden=obtener("orden")
		if not isNumeric(orden) then
			orden="0"
		end if						
										
									
			existefacultad=0
			
			if codfacultad<>"" then
			sql="select count(*) from facultad where descripcion='" & descripcion & "' and codfacultad<>" & codfacultad & " and codgrupofacultad=" & codgrupofacultad 
			else
			sql="select count(*) from facultad where descripcion='" & descripcion & "' and codgrupofacultad=" & codgrupofacultad 
			end if
			consultar sql,RS
			existefacultad=RS.Fields(0)
			RS.Close			
			if existefacultad=0 then			
				if obtener("agregardato")="1" then		
				sql="insert into facultad (codgrupofacultad,descripcion,pagina,orden,usuarioregistra,fecharegistra) values (" & codgrupofacultad & ",'" & descripcion & "','" & pagina & "','" & orden & "'," & session("codusuario") & ",getdate())"
				else
					sql="update facultad set codgrupofacultad=" & codgrupofacultad & ",descripcion='" & descripcion & "',pagina='" & pagina & "',orden=" & orden & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codfacultad=" & codfacultad
				end if
				''Response.Write sql
				conn.execute sql
									
				%>
				<script language="javascript">
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre")="dcs_admfacultad.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language="javascript">
					alert("El usuario ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codfacultad<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from facultad A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where a.codfacultad = " & codfacultad
					consultar sql,RS
					descripcion=rs.Fields("descripcion")
					codgrupofacultad=rs.Fields("codgrupofacultad")		
					pagina=rs.Fields("pagina")		
					orden=rs.Fields("orden")
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					RS.Close
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		<head>
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
function showCalendar2(id, format, showsTime, showsOtherMonths) {
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
			<title><%if codfacultad="" then%>Nuevo <%end if%>Privilegio</title>
			
			<link rel="stylesheet" href="assets/css/css/animation.css"/>
			<link rel="stylesheet" href="assets/css/custom.css" />
			<link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"/>
			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->
			
			<script language="javascript">
				var limpioclave=0;
				<%if codfacultad="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(trim(formula.pagina.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
																		
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(trim(formula.pagina.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
					
					document.formula.agregardato.value=2;
					document.formula.submit();
				}				
				<%end if%>
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
		</head>
		<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
				<form name="formula" method="post" action="dcs_definiraccion.asp">					
					<tr class="fondo-red">	
						<td class="text-withe" >			
							<font size="2"><b>&nbsp;<b><%if codfacultad="" then%><%end if%>Definir Acción</b></b></font>	
						</td>
						<td align="right" height="40">
							<a href="javascript:agregar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
							<a href="javascript:window.close();"><i class="logout demo-icon icon-logout">&#xe800;</i></a>&nbsp;
						</td>	
									
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange"><font size="2">Tipo Acción:</font></td>
						<td>
							<select name="idtipoaccion" style="font-size: xx-small; width: 200px;">
							<OPTION value="0">Seleccione una acción</OPTION>
							<%
							sql = "select IDTipoAccion, descripcion from TipoAccion where Activo = 1 order by IDTipoAccion"
							consultar sql,RS
							Do While Not  RS.EOF
							%>
								<option value="<%=RS.Fields("idtipoaccion")%>" <% if idtipoaccion<>"" then%><% if RS.fields("idtipoaccion")=int(idtipoaccion) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
							<%
							RS.MoveNext
							loop
							RS.Close
							%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="text-orange" width="30%"><font size="2" >Fecha / Hora de inicio:</font></td>
						<td><input name="fechaasigini"  id="sel1" readonly  type=text maxlength=10 size=10 value="<%if IsDate(fechaasigini) then%><%=fechaasigini%><%else%><%=maxfechagestion%><%end if%>"  style="font-size: x-small; width: 60px;"><i class="demo-icon2 icon-calendar" style="padding-left: 1px; vertical-align: center; cursor: pointer;" onclick="return showCalendar('sel1', '%d/%m/%Y');" >&#xe821;</i>	
						</td>

					</tr>
					<tr class="fondo-gris">
						<td class="text-orange" width="30%"><font size="2">Fecha / Hora de ejecución:</font></td>
						<td><input name="fechaexec"  id="sel2" readonly  type=text maxlength=10 size=10 value="<%if IsDate(fechaexec) then%><%=fechaexec%><%else%><%=maxfechagestion2%><%end if%>"  style="font-size: x-small; width: 60px;"><i class="demo-icon2 icon-calendar" style="padding-left: 1px; vertical-align: center; cursor: pointer;" onclick="return showCalendar('sel2', '%d/%m/%Y');" >&#xe821;</i>						
						</td>
					</tr>			
					<tr class="fondo-red">					
						<td><font size="2" >&nbsp;</font></td>
						<td align="right" height="40">
						</td>	
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange"><font size="2">Tipo Acción:</font></td>
						<td>
							<select name="idtipoaccion" style="font-size: xx-small; width: 200px;">
							<OPTION value="0">Seleccione una acción</OPTION>
							<%
							sql = "select IDTipoAccion, descripcion from TipoAccion where Activo = 1 order by IDTipoAccion"
							consultar sql,RS
							Do While Not  RS.EOF
							%>
								<option value="<%=RS.Fields("idtipoaccion")%>" <% if idtipoaccion<>"" then%><% if RS.fields("idtipoaccion")=int(idtipoaccion) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
							<%
							RS.MoveNext
							loop
							RS.Close
							%>
							</select>
						</td>
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange"><font size="2">Tipo Acción:</font></td>
						<td>
							<select name="idtipoaccion" style="font-size: xx-small; width: 200px;">
							<OPTION value="0">Seleccione una acción</OPTION>
							<%
							sql = "select IDTipoAccion, descripcion from TipoAccion where Activo = 1 order by IDTipoAccion"
							consultar sql,RS
							Do While Not  RS.EOF
							%>
								<option value="<%=RS.Fields("idtipoaccion")%>" <% if idtipoaccion<>"" then%><% if RS.fields("idtipoaccion")=int(idtipoaccion) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
							<%
							RS.MoveNext
							loop
							RS.Close
							%>
							</select>
						</td>
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange"><font size="2">Tipo Acción:</font></td>
						<td>
							<select name="idtipoaccion" style="font-size: xx-small; width: 200px;">
							<OPTION value="0">Seleccione una acción</OPTION>
							<%
							sql = "select IDTipoAccion, descripcion from TipoAccion where Activo = 1 order by IDTipoAccion"
							consultar sql,RS
							Do While Not  RS.EOF
							%>
								<option value="<%=RS.Fields("idtipoaccion")%>" <% if idtipoaccion<>"" then%><% if RS.fields("idtipoaccion")=int(idtipoaccion) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
							<%
							RS.MoveNext
							loop
							RS.Close
							%>
							</select>
						</td>
					</tr>
							<input type="hidden" name="agregardato" value="">
							<input type="hidden" name="codfacultad" value="<%=codfacultad%>">
							<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
							<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
						</form>	
					</table>
			</body>
		</html>	
		<%		
		end if
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

