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
			codtipotelefono=obtener("tipotelf")
			prefijo=right("000" & trim(obtener("prefijo")),3)
			if prefijo="000" then
			    prefijo="001"
			end if
			''telefono=trim(obtener("telefono"))
			telefono=replace(replace(replace(replace(replace(replace(trim(obtener("telefono")),",",""),".",""),"-",""),"+",""),"\",""),"$","")
			
			''X es otros
			if codtipotelefono<>"X" then
			    if IsNumeric(telefono) then 
				    telefono=int(telefono) 
			    else
				    telefono=0
			    end if
			end if
			extension=trim(obtener("extension"))
			if ISNumeric(replace(replace(replace(replace(replace(replace("extension",",",""),".",""),"-",""),"+",""),"\",""),"$","")) then 
				extension=int(extension)
			else
				extension=""
			end if
			
			existetelf=0
			
			if telefono>=0 then
			sql="select count(*) from telefononuevo where codigocentral = '" & codigocentral & "' and  fono='" & telefono & "'"
			end if
			consultar sql,RS
			existetelf=RS.Fields(0)
			RS.Close
			
			sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
			consultar sql,RS	
			maxfechagestion=rs.fields(0)
			RS.Close				
				
			if mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)=CStr(maxfechagestion) then
				vistabusqueda="VERULTIMPAGADO"
			else
				vistabusqueda="VERIMPAGADO"
			end if	
			
			if telefono>=0 then
			sql="select count(*) from " & vistabusqueda & " where codigocentral = '" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "' and (fono1='" & telefono & "' or fono2='" & telefono & "' or fono3='" & telefono & "' or fono4='" & telefono & "' or fono5='" & telefono & "')"
			end if
			consultar sql,RS
			existetelf=existetelf + RS.Fields(0)
			RS.Close				
						
			existeinactivo=0						
			if existetelf>0 then
				sql="select top 1 activo from telefononuevo where codigocentral = '" & codigocentral & "' and  fono='" & telefono & "' order by activo"
				consultar sql,RS
				if not RS.EOF then
					if RS.Fields(0)=0 then
						existeinactivo=1
					end if
				end if
				RS.Close			
			end if			
			
			if existetelf=0 or existeinactivo=1 then			
				if existeinactivo=0 then		
				sql="insert into telefononuevo (codigocentral,codtipotelefono,prefijo,fono,extension,activo,usuarioregistra,fecharegistra) values ('" & codigocentral & "','" & codtipotelefono & "','" & prefijo & "','" & telefono & "','" & extension & "',1," & session("codusuario") & ",getdate())"
				else
				sql="update telefononuevo set activo=1,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codigocentral = '" & codigocentral & "' and fono='" & telefono & "'" 
				end if
				''Response.Write sql
				conn.execute sql									
				%>
				<!--<script language=javascript>
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre1")= "verimpagado.asp" then%>window.open("<%=obtener("paginapadre1")%>?vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>&agreguetelf=1","<%=obtener("vistapadre1")%>");<%end if%>
					<%if obtener("paginapadre1")= "verficha.asp" then%>window.open("<%=obtener("paginapadre1")%>?vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&fdcodcen=<%=mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & codigocentral%>&fechavisita=<%=fechavisita%>&horavisita=<%=horavisita%>&minutovisita=<%=minutovisita%>&tipocontacto=<%=tipocontacto%>&dirsistema=<%=dirsistema%>&coddireccionnueva=<%=coddireccionnueva%>&agreguetelf=1","<%=obtener("vistapadre1")%>");<%end if%>
					window.close();
				</script>-->
                <%if obtener("paginapadre1")= "verimpagado.asp" then%>						
				    <script language=javascript>
				        window.open("<%=obtener("paginapadre1")%>?vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>&agreguetelf=1","<%=obtener("vistapadre1")%>");
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
					    <input type="hidden" name="coddireccionnueva" value="<%=coddireccionnueva%>">
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
					    <input type="hidden" name="agreguetelf" value="1">
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
					alert("El Teléfono ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
		%>				
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Nuevo Teléfono</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				function agregar()
				{
					if(trim(formula.prefijo.value)==""){alert("Debe ingresar el Código de la Ciudad del Télefono.");return;}
					if(trim(formula.telefono.value)==""){alert("Debe ingresar un Número de Teléfono.");return;}

					document.formula.agregardato.value=1;
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
				<form name=formula method=post action="adicionartelefono.asp">
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>		
					<tr bgcolor="#007DC5">
					<td colspan="2" align="left" height="18"><font size=2 face=Arial color="#FFFFFF"><b>Agregar teléfono</b></font>
					</td>
					</tr>										  
					<tr>
					<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Tipo:</font></td>
						<td>
						<select name="tipotelf" style="font-size: x-small; width: 200px;">
						<%
						sql = "select codtipotelefono,descripcion from TipoTelefono where activo=1 order by descripcion"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
						<option value="<%=RS.Fields("codtipotelefono")%>"><%=RS.Fields("codtipotelefono") & " - " & RS.Fields("descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>							
						</select>
						</td>
					</tr>					
					<tr>
						<td width=30% bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Prefijo:</font></td>
						<td bgcolor ="#f5f5f5"><input name="prefijo" type=text maxlength=3 value="<%=prefijo%>" style="font-size: x-small; width: 30px;"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;(Código de Ciudad)</font></td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Teléfono:</font></td>
						<td><input name="telefono" type=text maxlength=9 value="<%=Telefono%>" style="font-size: x-small; width: 80px;"></td>
					</tr>
					<tr>
						<td width=30% bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Anexo:</font></td>
						<td bgcolor ="#f5f5f5"><input name="extension" type=text maxlength=4 value="<%=Anexo%>" style="font-size: x-small; width: 40px;"></td>
					</tr>
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><%if codoficina="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
					</table>
					<input type=hidden name="agregardato" value="">
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
					<input type="hidden" name="coddireccionnueva" value="<%=coddireccionnueva%>">
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

