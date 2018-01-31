<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisoCliente_Contacto("dcs_admContacto.asp") then
	buscador=obtener("buscador")	
	idClienteContacto=obtener("idClienteContacto")
		if obtener("agregardato")<>"" then
		IDCliente=obtener("IDCliente")
		Nombres=obtener("Nombres")
		pagina=obtener("pagina")
		orden=obtener("orden")
		if not isNumeric(orden) then
			orden="0"
		end if						
										
									
			existeCliente_Contacto=0
			
			if idClienteContacto<>"" then
			sql="select count(*) from Cliente_Contacto where Nombres='" & Nombres & "' and idClienteContacto<>" & idClienteContacto & " and IDCliente=" & IDCliente 
			else
			sql="select count(*) from Cliente_Contacto where Nombres='" & Nombres & "' and IDCliente=" & IDCliente 
			end if
			consultar sql,RS
			existeCliente_Contacto=RS.Fields(0)
			RS.Close			
			if existeCliente_Contacto=0 then			
				if obtener("agregardato")="1" then		
				sql="insert into Cliente_Contacto (IDCliente,Nombres,Cargo,Telefono,Email,usuarioregistra,fecharegistra) values (" & IDCliente & ",'" & Nombres & "','" & Cargo & "','" & Telefono & "','" & Email & "'," & session("codusuario") & ",getdate())"
				else
					sql="update Cliente_Contacto set IDCliente=" & IDCliente & ",Nombres='" & Nombres & "',Cargo='" & Cargo & "',Telefono=" & Telefono & ", Email=" & Email & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where idClienteContacto=" & idClienteContacto
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
					<%if obtener("paginapadre")="dcs_admContacto.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
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
			if idClienteContacto<>"" then
					sql="select cc.IDClienteContacto,c.IDCliente,cc.Nombres,cc.Cargo,cc.Telefono,cc.Email,cc.Activo from Cliente_Contacto cc inner join Cliente c on cc.IDCliente = c.IDCliente where cc.IDClienteContacto " & idClienteContacto
					consultar sql,RS
					Nombres=rs.Fields("Nombres")
					IDCliente=rs.Fields("IDCliente")		
					Cargo=rs.Fields("Cargo")		
					Telefono=rs.Fields("Telefono")
					Email=rs.Fields("Email")
					Activo=rs.Fields("Activo")
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
			<title><%if idClienteContacto="" then%>Nuevo <%end if%>Contacto</title>
			
			<link rel="stylesheet" href="assets/css/css/animation.css"/>
			<link rel="stylesheet" href="assets/css/custom.css" />
			<link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"/>
			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->
			
			<script language="javascript" src="scripts/popcalendar.js"></script> 
			<script language="javascript">
				var limpioclave=0;
				<%if idClienteContacto="" then%>
				function agregar()
				{
					if(trim(formula.Nombres.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(trim(formula.Cargo.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.Telefono.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
																		
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.Nombres.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(trim(formula.Cargo.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.Telefono.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
					
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
				<form name="formula" method="post" action="dcs_nuevoCliente_Contacto.asp">					
					<tr class="fondo-red">	
						<td class="text-withe" colspan="2">			
							<font size="2"><b>&nbsp;<b><%if idClienteContacto="" then%>Nuevo <%end if%>Contacto</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr class="fondo-gris" height="25">
						<td class="text-orange label-registra" colspan="2" align="right"><font size="1">Registró:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modificó:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<tr>
						<td class="text-orange" width="20%"><font  size="2">Nombres:</font></td>
						<td><input name="Nombres" type=text maxlength=200 value="<%=Nombres%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange"><font size="2">Cliente:</font></td>
						<td>
							<select name="IDCliente" style="font-size: xx-small; width: 200px;">
							<%
							sql = "select IDCliente, RazonSocial from Cliente order by IDCliente"
							consultar sql,RS
							Do While Not  RS.EOF
							%>
								<option value="<%=RS.Fields("IDCliente")%>" <% if IDCliente<>"" then%><% if RS.fields("IDCliente")=int(IDCliente) then%> selected<%end if%><%end if%>><%=RS.Fields("RazonSocial")%></option>
							<%
							RS.MoveNext
							loop
							RS.Close
							%>
							</select>
						</td>
					</tr>
					<tr>
						<td class="text-orange" width="30%"><font size="2" >P&aacute;gina:</font></td>
						<td><input name="pagina" type="text" maxlength="200" value="<%=pagina%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange" width="30%"><font size="2">Orden:</font></td>
						<td><input name="orden" type="text" maxlength="50" value="<%=orden%>" style="font-size: xx-small; width: 60px; text-align: right"></td>
					</tr>			
					<tr class="fondo-red">					
						<td><font size="2" >&nbsp;</font></td>
						<td align="right" height="40">
							<%if idClienteContacto="" then%>
							<a href="javascript:agregar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
							<%else%>
							<a href="javascript:modificar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
							<%end if%>
							<a href="javascript:window.close();"><i class="logout demo-icon icon-logout">&#xe800;</i></a>&nbsp;
						</td>					
					</tr>

							<input type="hidden" name="agregardato" value="">
							<input type="hidden" name="idClienteContacto" value="<%=idClienteContacto%>">
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

