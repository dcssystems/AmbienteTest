<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admimpagado.asp") or permisofacultad("admvisitas.asp") then
		codigocentral=obtener("codigocentral")
		contrato=obtener("contrato")
		fechadatos=obtener("fechadatos")
		fechagestion=obtener("fechagestion")
		
		'''Datos formulario visita
		fdcodcen=obtener("fdcodcen")
		fechavisita=obtener("fechavisita")
		horavisita=obtener("horavisita")
		minutovisita=obtener("minutovisita")
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
		
		if obtener("agregardato")<>"" then
			coddpto = obtener("coddpto")
			codprov = mid(obtener("codprov"),3,2)
			coddist = mid(obtener("coddist"),5,3)
            dirnuevo=trim(obtener("dirnuevo"))
			
			existedir=0
			
			if dirnuevo<>"" then
			    sql="select count(*) from direccionnueva where codigocentral = '" & codigocentral & "' and direccion='" & dirnuevo & "'"
                consultar sql,RS
			    existedir=RS.Fields(0)
			    RS.Close			
			end if
		
			
			sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
			consultar sql,RS	
			maxfechagestion=rs.fields(0)
			RS.Close				
				
			if mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)=CStr(maxfechagestion) then
				vistabusqueda="VERULTIMPAGADO"
			else
				vistabusqueda="VERIMPAGADO"
			end if			
		
			if dirnuevo<>"" then
			    sql="select count(*) from " & vistabusqueda & " where codigocentral = '" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "' and direccion='" & dirnuevo & "'"
			    consultar sql,RS
			    existedir=existedir + RS.Fields(0)
			    RS.Close			
			end if
						
			
			existeinactivo=0
			if existedir>0 then
				sql="select top 1 activo from direccionnueva where codigocentral = '" & codigocentral & "' and  direccion='" & dirnuevo & "' order by activo"
				consultar sql,RS
				if not RS.EOF then
					if RS.Fields(0)=0 then
						existeinactivo=1
					end if
				end if
				RS.Close			
			end if
			
			if existedir=0 or existeinactivo=1 then			
				if existeinactivo=0 then		
				sql="insert into direccionnueva (codigocentral,direccion,coddpto,codprov,coddist,activo,usuarioregistra,fecharegistra) values ('" & codigocentral & "','" & dirnuevo & "','" & coddpto & "','" & codprov & "','" & coddist & "',1," & session("codusuario") & ",getdate())"
				else
				sql="update direccionnueva set activo=1,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codigocentral = '" & codigocentral & "' and  direccion='" & dirnuevo & "'"
				end if
				conn.execute sql									
				%>
				<!--<script language=javascript>
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre1")= "verimpagado.asp" then%>window.open("<%=obtener("paginapadre1")%>?vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>&agreguedir=1","<%=obtener("vistapadre1")%>");<%end if%>
					<%if obtener("paginapadre1")= "verficha.asp" then%>window.open("<%=obtener("paginapadre1")%>?vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&fdcodcen=<%=mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & codigocentral%>&fechavisita=<%=fechavisita%>&horavisita=<%=horavisita%>&minutovisita=<%=minutovisita%>&tipocontacto=<%=tipocontacto%>&dirsistema=<%=dirsistema%>&coddireccionnueva=<%=coddireccionnueva%>&agreguedir=1","<%=obtener("vistapadre1")%>");<%end if%>
					window.close();
				</script>			-->
				<%if obtener("paginapadre1")= "verimpagado.asp" then%>
				    <script language=javascript>
				        window.open("<%=obtener("paginapadre1")%>?vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>&agreguedir=1","<%=obtener("vistapadre1")%>");
				        window.close();
				    </script>
				<%end if%>
				<%if obtener("paginapadre1")= "verficha.asp" then%>
				    <form name=formula method=post action="<%=obtener("paginapadre1")%>" target="<%=obtener("vistapadre1")%>">
					    <input type=hidden name="vistapadre1" value="<%=obtener("vistapadre1")%>">
					    <input type=hidden name="paginapadre1" value="<%=obtener("paginapadre1")%>">
					    <input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
					    <input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
					    <input type="hidden" name="codigocentral" value="<%=codigocentral%>">
					    <input type="hidden" name="contrato" value="<%=contrato%>">
					    <input type="hidden" name="fechadatos" value="<%=fechadatos%>">
					    <input type="hidden" name="fechagestion" value="<%=fechagestion%>">					    
					    <!--formulario de visita-->
					    <input type="hidden" name="fdcodcen" value="<%=fdcodcen%>">
					    <input type="hidden" name="fechavisita" value="<%=fechavisita%>">
					    <input type="hidden" name="horavisita" value="<%=horavisita%>">
					    <input type="hidden" name="minutovisita" value="<%=minutovisita%>">
					    <input type="hidden" name="tipocontacto" value="<%=tipocontacto%>">							    
					    <input type="hidden" name="dirsistema" value="<%=dirsistema%>">
					    <input type="hidden" name="coddireccionnueva" value=""> <!--Agregue nueva<%=coddireccionnueva%>">-->
		                <input type="hidden" name="situaciontel1" value="<%=situaciontel1%>">
		                <input type="hidden" name="situaciontel2" value="<%=situaciontel2%>">
		                <input type="hidden" name="situaciontel3" value="<%=situaciontel3%>">
		                <input type="hidden" name="situaciontel4" value="<%=situaciontel4%>">
		                <input type="hidden" name="situaciontel5" value="<%=situaciontel5%>">				    	
						<%
						sql="select codtelefononuevo from TelefonoNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codtelefononuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<input type="hidden" name="situaciontelnuevo<%=RS.fields("codtelefononuevo")%>" value="<%=obtener("situaciontelnuevo" & RS.fields("codtelefononuevo"))%>">
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>		
					    <input type="hidden" name="fechapdp" value="<%=fechapdp%>">	
						<input type="hidden" name="actprinc" value="<%=actprinc%>">
						<input type="hidden" name="actividadotros" value="<%=actividadotros%>">
                        <input type="hidden" name="actividadadicional" value="<%=actividadadicional%>">		
		                <input type="hidden" name="enactividad" value="<%=enactividad%>">		
		                <input type="hidden" name="localalquilado" value="<%=localalquilado%>">
		                <input type="hidden" name="localfinanciado" value="<%=localfinanciado%>">
		                <input type="hidden" name="facturacionprincipal" value="<%=facturacionprincipal%>">
		                <input type="hidden" name="facturacionadicional" value="<%=facturacionadicional%>">
		                <input type="hidden" name="puertacalle" value="<%=puertacalle%>">
		                <input type="hidden" name="ofadministrativa" value="<%=ofadministrativa%>">
		                <input type="hidden" name="casanegocio" value="<%=casanegocio%>">
		                <input type="hidden" name="laborartesanal" value="<%=laborartesanal%>">
		                <input type="hidden" name="existencias" value="<%=existencias%>">
		                <input type="hidden" name="nropersonas" value="<%=nropersonas%>">
		                <input type="hidden" name="motivoatraso" value="<%=motivoatraso%>">
		                <input type="hidden" name="otrascausasatraso" value="<%=otrascausasatraso%>">
		                <input type="hidden" name="afrontapago" value="<%=afrontapago%>">
		                <input type="hidden" name="otrosafronta" value="<%=otrosafronta%>">
		                <input type="hidden" name="cuestionacobro" value="<%=cuestionacobro%>">
		                <input type="hidden" name="nocontacto" value="<%=nocontacto%>">
		                <input type="hidden" name="comentario" value="<%=comentario%>">
		                <input type="hidden" name="estado" value="<%=estado%>">
					    <input type="hidden" name="agreguedir" value="1">
				    </form>
                    <script language=javascript>
                        document.formula.submit();
                        window.close();
                    </script>
				<%end if%>
			<%
			else
			%>
				<script language=javascript>
					alert("La Dirección ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			    dirnuevo=obtener("dirnuevo")
			    codterritorio=obtener("codterritorio")
			    descripcion=obtener("descripcion")
				coddpto = obtener("coddpto")
				codprov = obtener("codprov")
				coddist = obtener("coddist")	
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Nueva Dirección</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				function agregar()
				{
					
					if(trim(formula.dirnuevo.value)==""){alert("Debe ingresar una Dirección para ingreso al Sistema.");return;}
				
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				function actualizaubigeo()
				{
				document.formula.ubigeoact.value = 1;
				document.formula.agregardato.value="";
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
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
				<form name=formula method=post action="adicionardireccion.asp">
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>		
					<tr bgcolor="#007DC5">
					<td colspan="2" align="left" height="18"><font size=2 face=Arial color="#FFFFFF"><b>Agregar Dirección</b></font>
					</td>
					</tr>										  
					<tr>
					<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Departamento:</font></td>
					<td>
						<select name="coddpto" onchange="actualizaubigeo()" style="font-size: x-small; width: 200px;">
						<%
						sql = "select coddpto,departamento from ubigeo where departamento<>'' and coddpto<>'00' group by coddpto,departamento order by departamento"
						consultar sql,RS
						seleccion=0
						primercoddpto=""
						Do While Not  RS.EOF
							if primercoddpto="" then
								primercoddpto=RS.Fields("coddpto")
							end if
						%>
						<option value="<%=RS.Fields("coddpto")%>" <%if RS.fields("coddpto")=coddpto then%> selected<%seleccion=1%><%end if%>><%=RS.Fields("departamento")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						if seleccion=0 then
							coddpto=primercoddpto
						end if
						%>
						</select>
						</td>
					</tr>					
					<tr>
					<td bgcolor ="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Provincia:</font></td>
					<td bgcolor ="#f5f5f5">
						<select name = "codprov" onchange="actualizaubigeo()" style="font-size: x-small; width: 200px;">
						<%
						sql = "select codprov,provincia from ubigeo where coddpto = '" & coddpto & "' and provincia<>'' and codprov<>'00' group by codprov,provincia order by provincia"
						consultar sql,RS
						seleccion=0
						primercodprov=""
						Do While Not RS.EOF
							if primercodprov="" then
								primercodprov=coddpto & RS.Fields("codprov")
							end if
						%>
						<option value="<%=coddpto & RS.Fields("codprov")%>" <%if coddpto & RS.fields("codprov")=codprov then%> selected<%seleccion=1%><%end if%>><%=RS.Fields("provincia")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						if seleccion=0 then
							codprov=primercodprov
						end if						
						%>
						</select>
						</td>
					</tr>
					<tr>
					<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Distrito:</font></td>
					<td>
						<select name = "coddist" onchange="actualizaubigeo()" style="font-size: x-small; width: 200px;">
						<%
						sql = "select coddist,distrito from ubigeo where coddpto = '" & coddpto & "' and coddpto + codprov = '" & codprov & "' and distrito<>'' and coddist<>'000' group by coddist,distrito order by distrito"
						consultar sql,RS
						seleccion=0
						primercoddist=""
						Do While Not RS.EOF
							if primercoddist="" then
								primercoddist=codprov & RS.Fields("coddist")
							end if
						%>
						<option value="<%=codprov & RS.Fields("coddist")%>" <%if codprov & RS.fields("coddist")=coddist then%> selected<%seleccion=1%><%end if%>><%=RS.Fields("distrito")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						if seleccion=0 then
							coddist=primercoddist
						end if							
						%>
						</select>
						</td>
					</tr>
					<tr>
						<td width=120  bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Dirección:</font></td>
						<td bgcolor ="#f5f5f5"><input name="dirnuevo" type=text maxlength=200 value="" style="font-size: x-small; width: 500px;"></td>
					</tr>
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
					</table>
					<input type=hidden name="agregardato" value="">
					<input type=hidden name="ubigeoact" value="">
					<input type=hidden name="vistapadre1" value="<%=obtener("vistapadre1")%>">
					<input type=hidden name="paginapadre1" value="<%=obtener("paginapadre1")%>">
					<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
					<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
					<input type="hidden" name="codigocentral" value="<%=codigocentral%>">
					<input type="hidden" name="contrato" value="<%=contrato%>">
					<input type="hidden" name="fechadatos" value="<%=fechadatos%>">
					<input type="hidden" name="fechagestion" value="<%=fechagestion%>">
					<!--formulario de visita-->
					<input type="hidden" name="fdcodcen" value="<%=fdcodcen%>">
					<input type="hidden" name="fechavisita" value="<%=fechavisita%>">
					<input type="hidden" name="horavisita" value="<%=horavisita%>">
					<input type="hidden" name="minutovisita" value="<%=minutovisita%>">
					<input type="hidden" name="tipocontacto" value="<%=tipocontacto%>">					
					<input type="hidden" name="dirsistema" value="<%=dirsistema%>">
					<input type="hidden" name="coddireccionnueva" value=""> <!--Agregue nueva<%=coddireccionnueva%>">-->
	                <input type="hidden" name="situaciontel1" value="<%=situaciontel1%>">
	                <input type="hidden" name="situaciontel2" value="<%=situaciontel2%>">
	                <input type="hidden" name="situaciontel3" value="<%=situaciontel3%>">
	                <input type="hidden" name="situaciontel4" value="<%=situaciontel4%>">
	                <input type="hidden" name="situaciontel5" value="<%=situaciontel5%>">						
					<%
					sql="select codtelefononuevo from TelefonoNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codtelefononuevo desc"
					''Response.Write sql
					consultar sql,RS
					Do While Not RS.EOF						
					%>
					<input type="hidden" name="situaciontelnuevo<%=RS.fields("codtelefononuevo")%>" value="<%=obtener("situaciontelnuevo" & RS.fields("codtelefononuevo"))%>">
					<%
					RS.MoveNext
					Loop
					RS.Close
					%>	
					<input type="hidden" name="fechapdp" value="<%=fechapdp%>">
					<input type="hidden" name="actprinc" value="<%=actprinc%>">		
					<input type="hidden" name="actividadotros" value="<%=actividadotros%>">
					<input type="hidden" name="actividadadicional" value="<%=actividadadicional%>">
				    <input type="hidden" name="enactividad" value="<%=enactividad%>">		
	                <input type="hidden" name="localalquilado" value="<%=localalquilado%>">
	                <input type="hidden" name="localfinanciado" value="<%=localfinanciado%>">
	                <input type="hidden" name="facturacionprincipal" value="<%=facturacionprincipal%>">
	                <input type="hidden" name="facturacionadicional" value="<%=facturacionadicional%>">
	                <input type="hidden" name="puertacalle" value="<%=puertacalle%>">
	                <input type="hidden" name="ofadministrativa" value="<%=ofadministrativa%>">
	                <input type="hidden" name="casanegocio" value="<%=casanegocio%>">
	                <input type="hidden" name="laborartesanal" value="<%=laborartesanal%>">
	                <input type="hidden" name="existencias" value="<%=existencias%>">
	                <input type="hidden" name="nropersonas" value="<%=nropersonas%>">
	                <input type="hidden" name="motivoatraso" value="<%=motivoatraso%>">
	                <input type="hidden" name="otrascausasatraso" value="<%=otrascausasatraso%>">
	                <input type="hidden" name="afrontapago" value="<%=afrontapago%>">
	                <input type="hidden" name="otrosafronta" value="<%=otrosafronta%>">
	                <input type="hidden" name="cuestionacobro" value="<%=cuestionacobro%>">
	                <input type="hidden" name="nocontacto" value="<%=nocontacto%>">
	                <input type="hidden" name="comentario" value="<%=comentario%>">
	                <input type="hidden" name="estado" value="<%=estado%>">
				</form>	
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

